VERSION 5.00
Begin VB.Form frmSousaTakuVerKanri 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�o�[�W�����Ǘ��i�����j"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.CommandButton cmdClear 
      Caption         =   "���[�N�N���A"
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
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyBaitai_Work 
      Caption         =   " �}�� �� ���[�N�@�R�s�["
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
      Top             =   3360
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
      TabIndex        =   16
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyOld_Jikko 
      Caption         =   "   �� �� ���s   �R�s�["
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
      TabIndex        =   15
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton CmdRemove 
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
      TabIndex        =   13
      Top             =   6960
      Width           =   2415
   End
   Begin VB.ListBox lstTaku 
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
      TabIndex        =   6
      Top             =   2520
      Width           =   8655
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
      TabIndex        =   5
      Top             =   6240
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
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Frame fraVersion 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   9360
      TabIndex        =   7
      Top             =   480
      Width           =   2055
      Begin VB.CheckBox chkFolder 
         Caption         =   "W ���[�N"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Value           =   1  '����
         Width           =   1380
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "N ���s"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Value           =   1  '����
         Width           =   1380
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "O ��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Value           =   1  '����
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " �o�[�W�����Ǘ�  ��ʂ֖߂�"
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
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Left            =   9000
      Top             =   7200
   End
   Begin VB.Label lblZenVer 
      Caption         =   "�S�̃o�[�W����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   19
      Top             =   600
      Width           =   8895
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�����o�[�W�����Ǘ�"
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
      TabIndex        =   14
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblKan 
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
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "̫���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   11
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label lblKan 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "����(�޲�)"
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
      Index           =   2
      Left            =   3600
      TabIndex        =   10
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblKan 
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
      Index           =   3
      Left            =   5040
      TabIndex        =   9
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblKan 
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
      Index           =   4
      Left            =   7080
      TabIndex        =   8
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmSousaTakuVerKanri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmSousaTakuVerKanri.frm
'//  �p�b�P�[�W���F�o�[�W�����Ǘ�(�Ď���)���
'//
'//  �T�v�F�o�[�W�����Ǘ�(�Ď���)���
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.100�֘A�z
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.273�C���Ή��z
'//     REVISIONS :(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                 �y�w�E����No.02�C���Ή��z
'//                 �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//     REVISIONS :(EG20 5.8.0.1) 2012-04-17   REVISED BY [TCC] T.Furuya
'//                 EG20 �t�F�[�Y2,3�����Ή�
'//     REVISIONS :(EG20 V5.9.0.1) 2012-05-02   REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή�
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

'�t�H���_��ʕ�
Public mlngChkFolderType        As Long

'Dim uVersion() As MN_VERSION_LIST       '�o�[�W�������i�[�G���A

Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l

Private Const HEADERTITLE_WRK = "�����o�[�W�����i���[�N�j�F"
Private Const HEADERTITLE_NOW = "�@�@�@�@�@�@�@�@�i���s�j�@�F"
Private Const HEADERTITLE_OLD = "�@�@�@�@�@�@�@�@�i���j�@�@�F"
Private Const HEADERVERSION_NON = "--.--.--.--"


'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdClear_Click
'//  �@�\����    �F���[�N�N���A
'//  �@�\�T�v    �F
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F
'//  �߂�l      �F
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-11-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                 �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//  REVISIONS   �F (EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                 �y�^���\�����P�Ή��z
'//  REVISIONS   �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub CmdClear_Click()

    Dim iResponse As Integer         ' MsgBox�{�^���R�[�h
    Dim bResult As Boolean           ' ��������

    On Error Resume Next

    '�u�o�[�W�����Ǘ���ʁF���[�N�N���A�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_WRK_CREA_BUTTOM, 0)

    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("�u���[�N�v�t�H���_���̃t�@�C�����A" _
           & Chr(vbKeyReturn) & "�S�č폜���܂��B    ��낵���ł����H", _
           vbOKCancel + vbExclamation, _
           "���[�N �N���A")

    If iResponse <> vbCancel Then
        sCmdBtnEnabled False                        ' ��ʑ���s��
        
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        '[�͂�] �{�^����I�������ꍇ
        '���[�N�t�H���_���̃t�@�C�����폜����
        bResult = sWrkFolderRemove
        sCmdBtnEnabled True                         ' ��ʑ����
        If bResult = True Then
            ' �����̃o�[�W��������\������
            Call psVersionDisp
        
' EG20 V5.8.0.1�폜�J�n
'            ' �^����ԍX�V                                              ' EG20 V5.5.0.1�ǉ�
'            Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1�ǉ�
' EG20 V5.8.0.1�폜�I��
' EG20 V5.8.0.1�ǉ��J�n
            ' �^����ԍX�V
'            Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_TAKU, BOOTINFO_UNKAI_NASHI)    ' EG20 V5.11.0.1�폜
            Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_TAKU, BOOTINFO_UNKAI_CLEAR)     ' EG20 V5.11.0.1�ǉ�
            Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMETAKU, BOOTINFO_UNKAI_NASHI)
' EG20 V5.8.0.1�ǉ��I��
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            '�u����I���v�|�b�v�A�b�v��ʕ\��
            MsgBox "����I�����܂����B", _
                   vbOKOnly + vbInformation, _
                   "���s����"
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

        End If
    End If

' EG20 V5.6.0.1�ǉ��J�n
    pubSubCreateFolder (PATH_OPERATE_APL)
    pubSubCreateFolder (PATH_OPERATE_APLNEW)
    pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1�ǉ��I��
' EG20 V6.9.0.1ADD START
    pubSubCreateFolder (FLD_OPERATEPROGNOW)
    pubSubCreateFolder (FLD_OPERATEPROGWRK)
    pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END

End Sub

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
'//  REVISIONS   �F(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyBaitai_Work_Click()
    Dim iResponse As Integer         'MsgBox�{�^���R�[�h

    On Error Resume Next

    '�u�}�́����[�N�R�s�[�v�{�^���̏ꍇ�B
    '�u�o�[�W�����Ǘ���ʁF�}�́����[�N�R�s�[�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_USB_COPY_WRK_BUTTOM, 0)

    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("�C���X�g�[���}�̂����[�N�t�H���_��" _
           & Chr(vbKeyReturn) & "�R�s�[���܂��B��낵���ł����H", _
           vbOKCancel + vbExclamation, _
           "�}�́����[�N�R�s�[")
    If iResponse <> vbCancel Then
        '[�͂�] �{�^����I�������ꍇ
        sCmdBtnEnabled False                        ' ��ʑ���s��
        '�C���X�g�[���}�̂����[�N�t�H���_���ɃR�s�[����
        Call sFDInstall
        sCmdBtnEnabled True                         ' ��ʑ����

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�J�n
'        ' �����̃o�[�W��������\������
'        Call psVersionDisp
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�I��
    End If

' EG20 V5.6.0.1�ǉ��J�n
    pubSubCreateFolder (PATH_OPERATE_APL)
    pubSubCreateFolder (PATH_OPERATE_APLNEW)
    pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1�ǉ��I��
' EG20 V6.9.0.1ADD START
    pubSubCreateFolder (FLD_OPERATEPROGNOW)
    pubSubCreateFolder (FLD_OPERATEPROGWRK)
    pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END

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
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.372�C���Ή��z
'//  REVISIONS   �F(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY[TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyOld_Jikko_Click()
    Dim iResponse As Integer         'MsgBox�{�^���R�[�h
    Dim bRet As Boolean              ' ��������

    On Error Resume Next

    '�u�o�[�W�����Ǘ���ʁF�������s�R�s�[�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OLD_COPY_NOW_BUTTOM, 0)

    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("���s�t�H���_���N���A�����t�H���_��" _
           & Chr(vbKeyReturn) & "�t�@�C�����R�s�[���܂�����낵���ł����H", _
           vbOKCancel + vbExclamation, _
           "�������s�R�s�[")
    If iResponse <> vbCancel Then
        
'EG20 V3.6.0.1�y03����TR-No.372�C���Ή��z�ǉ��J�n
        ' ���o�[�W�����t�H���_�ɑ�\�o�[�W�����t�@�C�������݂��Ȃ��ꍇ�ُ͈�Ƃ���B
        ' ���o�[�W�����EOPERATE�E�����E
        bRet = dllCheckAplVersion(4, PATH_OPERATE_APL, 3)
        If bRet = False Then
'            MsgBox "�ُ�I�����܂����B", vbCritical, "�������s�@�R�s�["        ' EG20 V5.8.0.1�폜
            MsgBox "�ُ�I�����܂����B", vbCritical, "���s����"                 ' EG20 V5.8.0.1�ǉ�
' EG20 V5.6.0.1�ǉ��J�n
            pubSubCreateFolder (PATH_OPERATE_APL)
            pubSubCreateFolder (PATH_OPERATE_APLNEW)
            pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1�ǉ��I��
' EG20 V6.9.0.1ADD START
            pubSubCreateFolder (FLD_OPERATEPROGNOW)
            pubSubCreateFolder (FLD_OPERATEPROGWRK)
            pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END
            Exit Sub
        End If
'EG20 V3.6.0.1�y03����TR-No.372�C���Ή��z�ǉ��I��
        
        '[�͂�] �{�^����I�������ꍇ
        sCmdBtnEnabled False                        ' ��ʑ���s��
        '�C���X�g�[���}�̂����[�N�t�H���_���ɃR�s�[����
        Call sVersionRollBack
        sCmdBtnEnabled True                         ' ��ʑ����
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�J�n
'        ' �����̃o�[�W��������\������
'        Call psVersionDisp
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�I��
    End If

' EG20 V5.6.0.1�ǉ��J�n
    pubSubCreateFolder (PATH_OPERATE_APL)
    pubSubCreateFolder (PATH_OPERATE_APLNEW)
    pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1�ǉ��I��
' EG20 V6.9.0.1ADD START
    pubSubCreateFolder (FLD_OPERATEPROGNOW)
    pubSubCreateFolder (FLD_OPERATEPROGWRK)
    pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END

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
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.372�C���Ή��z
'//  REVISIONS   �F(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY[TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyWork_Jikko_Click()
    Dim iResponse As Integer         'MsgBox�{�^���R�[�h
    Dim bRet As Boolean              ' ��������

    On Error Resume Next

    '�u�o�[�W�����Ǘ���ʁF�������s�R�s�[�t�����v���O�o��
'    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OLD_COPY_NOW_BUTTOM, 0)
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_WRK_COPY_NOW_BUTTOM, 0)

    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("���s�t�H���_���N���A�����[�N�t�H���_��" _
            & Chr(vbKeyReturn) & "�t�@�C�����R�s�[���܂�����낵���ł����H", _
           vbOKCancel + vbExclamation, _
           "���[�N�����s�R�s�[")
    If iResponse <> vbCancel Then
        
'EG20 V3.6.0.1�y03����TR-No.372�C���Ή��z�ǉ��J�n
        ' ���o�[�W�����t�H���_�ɑ�\�o�[�W�����t�@�C�������݂��Ȃ��ꍇ�ُ͈�Ƃ���B
        ' ���o�[�W�����EOPERATE�E�����E
        bRet = dllCheckAplVersion(1, PATH_OPERATE_APL, 3)
        If bRet = False Then
'            MsgBox "�ُ�I�����܂����B", vbCritical, "���[�N�����s �R�s�["     ' EG20 V5.8.0.1�폜
            MsgBox "�ُ�I�����܂����B", vbCritical, "���s����"                 ' EG20 V5.8.0.1�ǉ�
' EG20 V5.6.0.1�ǉ��J�n
            pubSubCreateFolder (PATH_OPERATE_APL)
            pubSubCreateFolder (PATH_OPERATE_APLNEW)
            pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1�ǉ��I��
' EG20 V6.9.0.1ADD START
            pubSubCreateFolder (FLD_OPERATEPROGNOW)
            pubSubCreateFolder (FLD_OPERATEPROGWRK)
            pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END
            Exit Sub
        End If
'EG20 V3.6.0.1�y03����TR-No.372�C���Ή��z�ǉ��I��
        
        '[�͂�] �{�^����I�������ꍇ
        sCmdBtnEnabled False                        ' ��ʑ���s��
        '�C���X�g�[���}�̂����[�N�t�H���_���ɃR�s�[����
        Call sVersionUpdate
        sCmdBtnEnabled True                         ' ��ʑ����

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�J�n
'        ' �����̃o�[�W��������\������
'        Call psVersionDisp
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�I��
    End If

' EG20 V5.6.0.1�ǉ��J�n
    pubSubCreateFolder (PATH_OPERATE_APL)
    pubSubCreateFolder (PATH_OPERATE_APLNEW)
    pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1�ǉ��I��
' EG20 V6.9.0.1ADD START
    pubSubCreateFolder (FLD_OPERATEPROGNOW)
    pubSubCreateFolder (FLD_OPERATEPROGWRK)
    pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �o�[�W�����Ǘ�(�Ď���)���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}�N��
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
Private Sub Form_Activate()
   On Error Resume Next
 
    '�^�C�}���N������
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �o�[�W�����Ǘ�(�Ď���)���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}��~
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
Private Sub Form_Deactivate()
   On Error Resume Next
    
    '�^�C�}���~����
    tmrMail.Enabled = False
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FForm_Load
'//  �@�\����    �F�o�[�W�����Ǘ�(�Ď���)���(���[�h��)
'//  �@�\�T�v    �F�����������s���B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F�Ȃ�
'//  �߂�l      �F�Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    On Error Resume Next
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
   
   '�u�����o�[�W�����Ǘ���ʁF�\���v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_VERASION_TAKU_GAMEN_START, 0)

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    '������
    lstTaku.Clear
    mlngChkFolderType = 0

    '�t�H���_�I�𕔁F�I��L��
    chkFolder(0).Value = 1
    chkFolder(1).Value = 1
    chkFolder(2).Value = 1
    
    mlngChkFolderType = 7
    
    ' �����̃o�[�W��������\������
    Call psVersionDisp
   
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

'   ���[����M�p�̃��[����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False

End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FchkFolder_Click
'//  �@�\����    �F�u�t�H���_�`�F�b�N�v�`�F�b�N��������
'//  �@�\�T�v    �F�t�H���_�I�𕔃`�F�b�N���s���B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F�Ȃ�
'//  �߂�l      �F�Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
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

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdRefresh_Click
'//  �@�\����    �F�u�\���X�V�v�t��������
'//  �@�\�T�v    �F�ŐV�̏�Ԃ�\������B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F�Ȃ�
'//  �߂�l      �F�Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdRefresh_Click()
    Dim i As Integer        '�J�E���^�[
    Dim bFlag As Boolean    '�\���t�H���_�I���`�F�b�N(TRUE�F�`�F�b�N�L�BFALSE�F�`�F�b�N��)
   
    On Error Resume Next
    
    '�u�����o�[�W�����Ǘ���ʁF�\���X�V�t�����v
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
                "�����o�[�W�����Ǘ�"
        Exit Sub
    End If
   
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
   
    ' �����̃o�[�W��������\������
    Call psVersionDisp

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
End Sub


'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdOutPut_Click
'//  �@�\����    �F�u�o�[�W�������}�̏o�́v�t��������
'//  �@�\�T�v    �F�\�����ꂽ�o�[�W��������}�̂ɏo�͂���B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F�Ȃ�
'//  �߂�l      �F�Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.100�֘A�z
'//  REVISIONS   �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdOutPut_Click()
'*******************************
'VB�G���[����
On Error GoTo Error_cmdOutPut_Click
'*******************************
    Dim strCopySaki    As String        ' �o�͐�t�@�C���p�X
    Dim strWriteDir    As String        ' �o�͐�t�H���_
    Dim fso            As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim lngErrCode     As Long          '�G���[�R�[�h
    
    Dim strStationName As String        ' �w����
    Dim szCornerName   As String        ' �R�[�i����
    Dim nNullIndex     As Integer       ' ���������[�N
    Dim strWork        As String        ' ���[�N
    Dim strFileName    As String        ' �t�@�C����
    Dim bRet           As Boolean  '�߂�l

   '�u�Ď��Ճo�[�W�����Ǘ���ʁF�o�[�W�������}�̏o�͖t�����v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, VERASION_INFO_OUTPUT, 0)
    
' EG20 V3.3.0.1 �y����TR-No.100�֘A�z�ǉ��J�n
    ' ���X�g�ɂP�����f�[�^���Ȃ��ꍇ�ُ͈�I��
    If lstTaku.ListCount = 0 Then
        '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
        Exit Sub
    End If
' EG20 V3.3.0.1 �y����TR-No.100�֘A�z�ǉ��I��
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�t�H���_�I���|�b�v�A�b�v��ʕ\��
    strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)

    '�w��t�H���_�Ȃ�
    If Len(strWriteDir) = 0 Then
        Exit Sub
    End If
    
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    strStationName = gsGetStationEkiName
    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ �ێ��p�֐�:�����o�[�W�����t�@�C���i��ʕ\���p�j�쐬
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllCreateTakuVersionFile(mlngChkFolderType, TAKUVERLIST_REPORTFILE, VERLISTKIND_REPORT)
    
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
       Exit Sub
    End If
    
    '�t�@�C���̗L���m�F
    If fso.FileExists(TAKUVERLIST_REPORTFILE) = False Then
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
        Exit Sub
    End If
    strFileName = Dir(TAKUVERLIST_REPORTFILE)
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�J�n
'    '�t�H���_�I���|�b�v�A�b�v��ʕ\��
'    strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)
'
'    '�w��t�H���_�Ȃ�
'    If Len(strWriteDir) = 0 Then
'        Exit Sub
'    End If
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�I��
    
    '�R�s�[��t�H���_�̗L���m�F
    If fso.FolderExists(strWriteDir) = False Then
        '�R�s�[��t�H���_�쐬
        fso.CreateFolder (strWriteDir)
    End If
    
    '�R�s�[��t�@�C�����쐬
    strCopySaki = strWriteDir & "\" & strStationName & "_" & strFileName

    '�t�@�C���R�s�[�i���ɑ��݂����ꍇ�͏㏑�����邷��j
    fso.CopyFile TAKUVERLIST_REPORTFILE, strCopySaki, True
  
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '�o�͌��ʃ|�b�v�A�b�v(����)�\��
    MsgBox "����I�����܂����B", vbInformation + vbOKOnly, "�o�͌���"
    '�u�����o�[�W�����Ǘ���ʁF�o�[�W�������}�̏o�͏�������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, VERASION_INFO_OUTPUT_OK, 0)
    
    Set fso = Nothing
    
    Exit Sub
'*******************************
'VB�G���[����
Error_cmdOutPut_Click:
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '�����ُ�̏ꍇ�A�o�͌��ʃ|�b�v�A�b�v(�ُ�)�\��
    MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
    '�u�����o�[�W�����Ǘ���ʁF�o�[�W�������}�̏o�͏����ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_OUTPUT_ERROR, lngErrCode)
    Set fso = Nothing
'*******************************
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdRemove_Click
'//  �@�\����    �F�u�}�̎�O�v�t��������
'//  �@�\�T�v    �F�}�̂̎��O�����s���B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F�Ȃ�
'//  �߂�l      �F�Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdRemove_Click()
   
   On Error Resume Next
       
   '�u�}�̎�O�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '�}�̎�O����
    Call pfRemove(Me)
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FcmdReturn_Click
'//  �@�\����    �F�u���j���[��ʂ֖߂�v�t��������
'//  �@�\�T�v    �F����ʂ���������B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F�Ȃ�
'//  �߂�l      �F�Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V5.6.0.1) 2012-04-07  CODED BY  [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    On Error Resume Next
    
    '�u�����o�[�W�����Ǘ���ʁF�����v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_VERASION_TAKU_GAMEN_END, 0)
 
 ' EG20 V5.6.0.1�ǉ��J�n
    pubSubCreateFolder (PATH_OPERATE_APL)
    pubSubCreateFolder (PATH_OPERATE_APLNEW)
    pubSubCreateFolder (PATH_OPERATE_APLOLD)
' EG20 V5.6.0.1�ǉ��I��
' EG20 V6.9.0.1ADD START
    pubSubCreateFolder (FLD_OPERATEPROGNOW)
    pubSubCreateFolder (FLD_OPERATEPROGWRK)
    pubSubCreateFolder (FLD_OPERATEPROGOLD)
' EG20 V6.9.0.1ADD END

    '�o�[�W�����Ǘ��i�����j��ʂ����
    Unload Me
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FpsVersionDisp
'//  �@�\����    �F�o�[�W�������쐬����
'//  �@�\�T�v    �F�o�[�W�������t�@�C���쐬/��ʕ\�����s���B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F�Ȃ�
'//  �߂�l      �F�Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����� ����No.36�֘A�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub psVersionDisp()
    Dim bRet            As Boolean  '�߂�l
    Dim intFileNo       As Integer  '�t�@�C���ԍ�
    Dim strWork         As String   '��ƃG���A
    Dim strVerData      As String   '�S�̃o�[�W����
    Dim lngErrCode      As Long     '�G���[�R�[�h
    Dim strList         As String
    Dim strVer          As String
    Dim strWork1        As String
    Dim strWork2        As String
    Dim strWork3        As String
    Dim sFileName       As String


'*******************************
'VB�G���[����
On Error GoTo Error_psVersionDisp
'*******************************

    '�}�̏o�͖t�����s��
    cmdOutPut.Enabled = False

    '���X�g������
    lstTaku.Clear
    
    '��ƃG���A������
    strWork = ""

    '�S�̃o�[�W����������
    strVerData = ""

    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '/ �ێ��p�֐�:�����o�[�W�����t�@�C���i��ʕ\���p�j�쐬
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllCreateTakuVersionFile(mlngChkFolderType, TAKUVERLIST_DISPFILE, VERLISTKIND_DISP)

    ' �o�[�W�����t�@�C������
    If bRet Then
       '�u�o�[�W�������t�@�C���쐬����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
    ' �o�[�W�����t�@�C�����s
    Else
       '�u�o�[�W�������t�@�C���쐬�ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
       Exit Sub
    End If
    
    ' �o�[�W�����t�@�C���̗L���m�F
    If Len(Trim(Dir(TAKUVERLIST_DISPFILE))) = 0 Then
        Exit Sub
    End If

    ' �o�[�W�����t�@�C���̃t�@�C���ԍ����擾����B
    intFileNo = FreeFile

    ' �o�[�W�����t�@�C���I�[�v��
    Open TAKUVERLIST_DISPFILE For Input As #intFileNo
    
        '���[�N
        Line Input #intFileNo, strWork
        
        If (Trim(strWork) = "") Then
            strVerData = HEADERTITLE_WRK & HEADERVERSION_NON & vbCrLf
        Else
            '�S�̃o�[�W����������쐬
            strVerData = strWork & vbCrLf
        End If

        '���s
        Line Input #intFileNo, strWork
        If (Trim(strWork) = "") Then
            strVerData = strVerData & HEADERTITLE_NOW & HEADERVERSION_NON & vbCrLf
        Else
            strVerData = strVerData & strWork & vbCrLf
        End If

        '��
        Line Input #intFileNo, strWork
        If (Trim(strWork) = "") Then
            strVerData = strVerData & HEADERTITLE_OLD & HEADERVERSION_NON & vbCrLf
        Else
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

                strWork1 = Right(strWork, 42)
                strWork2 = Mid(strWork1, 1, 12)   '�T�C�Y�̂ݒ��o
                strWork3 = Mid(strWork1, 13, 30)
                strVer = Format(strWork2, "#,##0")
                strVer = Format(strVer, "@@@@@@@@@@@@")
                sFileName = StrConv(MidB(StrConv(Mid(strWork, 1, 27) & Space(20), vbFromUnicode), 1, 27), vbUnicode)
                strList = sFileName & strVer & strWork3
                '���X�g�ɏo��
                lstTaku.AddItem (strList)

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
    '�o�[�W�������t�@�C���쐬�ُ탍�O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
    '�t�@�C���N���[�Y
    Close #intFileNo
'*******************************
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FtmrMail_Timer
'//  �@�\����    �F���[����M�^�C�}�A�^�C���A�b�v����
'//  �@�\�T�v    �F���[������M����B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F�Ȃ�
'//  �߂�l      �F�Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    '�ėp���[����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmSousaTakuVerKanri.Caption, False
        pfFormActive (frmSousaTakuVerKanri.hwnd)            ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    End If
End Sub

'/////////////////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'//
'//  �֐�����    �FsWrkFolderRemove
'//  �@�\����    �F���[�N�t�H���_���t�@�C���폜����
'//  �@�\�T�v    �F���[�N�t�H���_���̃t�@�C�����폜����B
'//
'//                   �^          ����            �Ӗ�
'//  ����        �F�Ȃ�
'//  �߂�l      �F�Ȃ�
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'/////////////////////////////////////////////////////////////////////////////
Private Function sWrkFolderRemove() As Boolean
    Dim stringWorkFolder As String      ' �t�H���_��
    Dim MyName As String                '�t�@�C����
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                     '�t�@�C���I�u�W�F�N�g
    Dim objFolder As Folder               '�t�H���_�I�u�W�F�N�g         ' EG20 V6.9.0.1 ADD
    
    On Error GoTo ErrorHandler          '�G���[�n���h���̓o�^

    '�����l�ݒ�
    sWrkFolderRemove = True
   
    '//////////////////////////////////////////////////////////////////////////
    '// �Ď��Ճt�H���_���̑���샏�[�N�t�H���_������
    ' ���[�N�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
    stringWorkFolder = FLD_OPERATEPROGWRK & "\"
    For Each objFi In objFso.GetFolder(stringWorkFolder).files  ' ���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then            ' �t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            '�t�@�C�����폜����
            Kill stringWorkFolder & MyName
        End If
    Next

' EG20 V6.9.0.1 ADD START
    For Each objFolder In objFso.GetFolder(stringWorkFolder).SubFolders  ' ���[�v���J�n
        If objFso.FolderExists(objFolder.Path) = True Then               ' �t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�����폜
            Call objFso.DeleteFolder(objFolder.Path)
        End If
    Next
' EG20 V6.9.0.1 ADD END

    '//////////////////////////////////////////////////////////////////////////
    '// ���[�N�t�H���_���̑����t�H���_������
    ' ���[�N�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
    stringWorkFolder = PATH_OPERATE_APLNEW
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
    
    Set objFso = Nothing
    Set objFi = Nothing
    Set objFolder = Nothing         ' EG20 V6.9.0.1 ADD

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�J�n
'    '�u����I���v�|�b�v�A�b�v��ʕ\��
'    MsgBox "����I�����܂����B", _
'           vbOKOnly + vbInformation, _
'           "���s����"
'
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�폜�I��
    
    Exit Function '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '�u���[�N�N���A�ُ�I���v�|�b�v�A�b�v��ʕ\��
     MsgBox "�ُ�I�����܂����B", _
           vbOKOnly + vbCritical, _
           "���s����"
           
   '�u�����ް�ޮ݁Fܰ�̫���̧�ٍ폜�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_WRKFILE_DELETE_ERROR, lngErrCode)
           
    sWrkFolderRemove = False
    Set objFso = Nothing
    Set objFi = Nothing
    Set objFolder = Nothing         ' EG20 V6.9.0.1 ADD
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//  REVISIONS   �F(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                 �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//  REVISIONS   �F(EG20 5.8.0.1) 2012-04-17   REVISED BY [TCC] T.Furuya
'//                 EG20 �t�F�[�Y2,3�����Ή�
'//  REVISIONS   �F(EG20 V5.9.0.1) 2012-05-02  REVISED BY [TCC] H.Sugimoto
'//  REVISIONS   �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F���D�@�o�[�W�����Ǘ���ʂ�sFDInstall���p
'///////////////////////////////////////////////////////////////////
Private Sub sFDInstall()

    Dim sInputPass As String                ' �C���X�g�[�����f�B���N�g����
    Dim sInputFolder As String              ' �C���X�g�[�����t�H���_��
    Dim objFso As New FileSystemObject      ' �t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                       ' �t�@�C���I�u�W�F�N�g
    Dim MyName As String                    ' �t�@�C���t���p�X��
    Dim sSrcFileName As String              ' �R�s�[���t�@�C����
    Dim sDstFileName As String              ' �R�s�[��t�@�C����
    Dim lngErrCode     As Long              ' �G���[�R�[�h
    Dim lngProcId As Long                   ' �v���Z�XID
    Dim hProc As Variant                    ' �v���Z�X�n���h��
    Dim objFolder As Folder                 ' �t�H���_�I�u�W�F�N�g          ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim FileName As String                  ' ���o�t�@�C����                ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim FileKaku As String                  ' �g���q                        ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim ExecCommand As String               ' ���s������                    ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim CurrentDirectory As String          ' �J�����g�f�B���N�g��          ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�
    Dim ExecDirectory As String             ' ���s�t�@�C���f�B���N�g��      ' EG20 V5.9.0.1�ǉ�

    On Error GoTo ErrorHandler              ' �G���[�n���h���̓o�^

    ' /////////////////////////////////////////////////////////////////////////
    ' // �C���X�g�[�����ނ̃R�s�[
    sInputPass = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)
    '�w��t�H���_�Ȃ�
    If Len(sInputPass) = 0 Then
        Set objFso = Nothing
        Set objFi = Nothing
        Exit Sub
    End If
 
    sInputFolder = sInputPass

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�폜�J�n
'    For Each objFi In objFso.GetFolder(sInputFolder).files      '���[�v���J�n
'        If objFso.FileExists(objFi.Path) = True Then            '�t�@�C�����̎擾�`�F�b�N
'            '�f�B���N�g�������擾
'            MyName = objFi.Name
'            '�}�̓��t�@�C�������쐬
'            sSrcFileName = sInputFolder & "\" & MyName
'            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
'            If (GetAttr(sSrcFileName) And vbDirectory) <> vbDirectory Then
'                '���[�N�t�H���_���t�@�C�������쐬����
'                sDstFileName = FLD_OPERATEPROGWRK & "\" & MyName
'                '�t�@�C���R�s�[�i���ɑ��݂����ꍇ�͏㏑�����邷��j
'                objFso.CopyFile sSrcFileName, sDstFileName, True
'            End If
'        End If
'    Next
' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�폜�I��
' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ��J�n
    '�t�H���_�I�u�W�F�N�g���擾
    Set objFolder = objFso.GetFolder(sInputFolder)

    '//////////////////////////////////////////////////////
    '// ���[�N�t�H���_������
    If objFso.FolderExists(FLD_OPERATEPROGWRK) Then
        Call objFso.DeleteFolder(FLD_OPERATEPROGWRK)
    End If

    objFolder.Copy FLD_OPERATEPROGWRK
' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ��I��

    
    ' /////////////////////////////////////////////////////////////////////////
    ' // �C���X�g�[�����ނ̒�����w��t�@�C���̎��s
    sSrcFileName = pfGetFileNameTakuExec
    ' �t�@�C���ݒ�Ȃ�
    If Len(sSrcFileName) = 0 Then
        Set objFso = Nothing
        Set objFi = Nothing
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        Call psVersionDisp
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
        Exit Sub
    End If
   
    sDstFileName = FLD_OPERATEPROGWRK & "\" & sSrcFileName
    If objFso.FileExists(sDstFileName) = False Then
        Set objFso = Nothing
        Set objFi = Nothing
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        Call psVersionDisp
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
        Exit Sub
    End If
    
    Set objFso = Nothing
    Set objFi = Nothing
    
'    lngProcId = Shell(sDstFileName, vbNormalFocus)             ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�폜
' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ��J�n
    ' �J�����g�f�B���N�g���擾
    CurrentDirectory = CurDir$()
'    Call ChDir(FLD_OPERATEPROGWRK)                             ' EG20 V5.9.0.1�폜
' EG20 V5.9.0.1�ǉ��J�n
    Call psFolderPathGet(sDstFileName, ExecDirectory)
    Call ChDrive("D")
    Call ChDir(ExecDirectory)
' EG20 V5.9.0.1�ǉ��I��

    ' �t�@�C�����O�擾
    psFileNameGet sDstFileName, FileName, FileKaku
    If UCase(FileKaku) = "VBS" Then
        ExecCommand = "wscript.exe " & sDstFileName
    Else
        ExecCommand = sDstFileName
    End If
    lngProcId = Shell(ExecCommand, vbNormalFocus)
' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ��I��

    hProc = OpenProcess(PROCESS_ALL_ACCESS, False, lngProcId)   ' �v���Z�X�n���h�����擾���܂��B
    If hProc > 0 Then                                           ' �v���Z�X�n���h�����擾�ł����ꍇ
        Call dllWaitForSingleObject(hProc)                      ' �v���Z�X���V�O�i����ԂɂȂ�܂ő҂��܂��B
        CloseHandle hProc                                       ' �v���Z�X�n���h����������܂��B
    End If
    
'    Call ChDir(CurrentDirectory)                ' EG20 V3.6.0.1�y����TR-No.273�C���Ή��z�ǉ�   ' EG20 V5.9.0.1�폜
    Call ChDir("D:\")                                           ' EG20 V5.9.0.1�ǉ�
    
' EG20 V5.8.0.1�폜�J�n
'    ' �^����ԍX�V                                              ' EG20 V5.5.0.1�ǉ�
'    Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI)           ' EG20 V5.5.0.1�ǉ�
' EG20 V5.8.0.1�폜�I��
' EG20 V5.8.0.1�ǉ��J�n
    ' �^����ԍX�V
    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_TAKU, BOOTINFO_UNKAI_ARI)
    Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMETAKU, BOOTINFO_UNKAI_ARI)
' EG20 V5.8.0.1�ǉ��I��
' EG20 V5.8.0.1 ADD START
    '�ǂݎ��O���̊֐������s
    dllChangeAttributeContents (PATH_OPERATE_APLNEW)
    '�ǂݎ��O���̊֐������s
    dllChangeAttributeContents (FLD_OPERATEPROGWRK)
' EG20 V5.8.0.1 ADD END
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    Call psVersionDisp
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    '�u����I���v�|�b�v�A�b�v��ʕ\��
    MsgBox "����I�����܂����B", _
           vbOKOnly + vbInformation, _
           "���s����"
    
    Exit Sub    '�������I������

ErrorHandler:   ' �G���[�����B
    Set objFso = Nothing
    Set objFi = Nothing

' EG20 V5.8.0.1 ADD START
    '�ǂݎ��O���̊֐������s
    dllChangeAttributeContents (PATH_OPERATE_APLNEW)
    '�ǂݎ��O���̊֐������s
    dllChangeAttributeContents (FLD_OPERATEPROGWRK)
' EG20 V5.8.0.1 ADD END

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    Call psVersionDisp
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
    '�u�����ް�ޮ݁F�}�́�ܰ���߰�����ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_USB_COPY_WRK_ERROR, lngErrCode)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : sVersionRollBack
'//  �@�\����  : �o�[�W�����߂�����
'//  �@�\�T�v  : ���s�o�[�W���������o�[�W�����֖߂�
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : String    �t�@�C����
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.273�C���Ή��z
'//  REVISIONS   �F(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                �y�w�E����No.02�C���Ή��z
'//                �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//  REVISIONS   �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Sub sVersionRollBack()

    Dim objFso As New FileSystemObject      ' �t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                       ' �t�@�C���I�u�W�F�N�g
    Dim objFolder As Folder                 ' �t�H���_�I�u�W�F�N�g
    Dim stringWorkFolder As String          ' �t�H���_��
    Dim MyName As String                    ' �t�@�C����
    Dim lngErrCode     As Long              ' �G���[�R�[�h
    Dim strSrcFile As String                ' �R�s�[��
    Dim strDstFile As String                ' �R�s�[��
    Dim bResult As Boolean                  ' ��������      ' EG20 V3.6.0.1�ǉ�
    Dim sSrcFileName As String              ' �R�s�[���t�@�C����    ' EG20 V5.8.0.1�ǉ�
    Dim sDstFileName As String              ' �R�s�[��t�@�C����    ' EG20 V5.8.0.1�ǉ�

    On Error GoTo ErrorHandler          '�G���[�n���h���̓o�^

' EG20 V5.8.0.1�ǉ��J�n
    ' /////////////////////////////////////////////////////////////////////////
    ' // ���[�N�t�H���_�̃t�@�C�����݃`�F�b�N
    stringWorkFolder = FLD_OPERATEPROGOLD
    If objFso.FolderExists(stringWorkFolder) <> True Then
        Set objFso = Nothing
        Set objFi = Nothing
        Set objFolder = Nothing
        ' �t�H���_�������������݂��Ȃ�
        MsgBox "�ُ�I�����܂����B", vbCritical, "���s����"
        Exit Sub                        ' �����I��
    End If
    
    strSrcFile = pfGetFileNameTakuExec
    ' �t�@�C���ݒ�Ȃ�
    If Len(strSrcFile) = 0 Then
        Set objFso = Nothing
        Set objFi = Nothing
        Set objFolder = Nothing
            MsgBox "�ُ�I�����܂����B", vbCritical, "���s����"
        Exit Sub
    End If

    strDstFile = stringWorkFolder & "\" & strSrcFile
    If objFso.FileExists(strDstFile) = False Then
        Set objFso = Nothing
        Set objFi = Nothing
        Set objFolder = Nothing
        ' �t�@�C�������݂��Ȃ�
        MsgBox "�ُ�I�����܂����B", vbCritical, "���s����"
        Exit Sub                        ' �����I��
    End If
    
' EG20 V5.8.0.1�ǉ��I��
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    ' /////////////////////////////////////////////////////////////////////////
    ' // �C���X�g�[�����ނ̃R�s�[
    
    '//////////////////////////////////////////////////////
    '// �Ď��Ճt�H���_���̑������s�t�H���_������
    stringWorkFolder = FLD_OPERATEPROGNOW
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If

    '//////////////////////////////////////////////////////
    '// �������s�R�s�[
    strSrcFile = FLD_OPERATEPROGOLD
    strDstFile = FLD_OPERATEPROGNOW

    '�t�H���_�I�u�W�F�N�g���擾
    Set objFolder = objFso.GetFolder(strSrcFile)
'    If objFolder.Size <> 0 Then                    ' EG20 V6.9.0.1 DEL
    If objFso.FolderExists(strSrcFile) Then         ' EG20 V6.9.0.1 ADD
        objFolder.Copy strDstFile
    End If

    ' /////////////////////////////////////////////////////////////////////////
    ' // �S�̂��R�s�[
    
    '//////////////////////////////////////////////////////
    '// �����t�H���_������
    stringWorkFolder = PATH_OPERATE_APL
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
    
    strSrcFile = PATH_OPERATE_APLOLD
    strDstFile = PATH_OPERATE_APL

    '�t�H���_�I�u�W�F�N�g���擾
    Set objFolder = objFso.GetFolder(strSrcFile)
'    If objFolder.Size <> 0 Then                    ' EG20 V6.9.0.1 DEL
    If objFso.FolderExists(strSrcFile) Then         ' EG20 V6.9.0.1 ADD
        objFolder.Copy strDstFile
    End If

    Set objFso = Nothing
    Set objFi = Nothing
    Set objFolder = Nothing

' EG20 V3.6.0.1�ǉ��J�n
    ' �����v���O�����f�[�^�쐬����
    bResult = pfTakuProgramVersionCreateProc
' EG20 V3.6.0.1�ǉ��I��

    If bResult = True Then
' EG20 V5.8.0.1�ǉ��J�n
        ' �^����ԍX�V
        Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_TAKU, BOOTINFO_UNKAI_KIRIKAE)
        Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMETAKU, BOOTINFO_UNKAI_KIRIKAE)
' EG20 V5.8.0.1�ǉ��I��
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        ' �����̃o�[�W��������\������
        Call psVersionDisp
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        '�u����I���v�|�b�v�A�b�v��ʕ\��
        MsgBox "����I�����܂����B", _
               vbOKOnly + vbInformation, _
               "���s����"
    Else
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        ' �����̃o�[�W��������\������
        Call psVersionDisp
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        '�u�ُ�I���v�|�b�v�A�b�v��ʕ\��
        MsgBox "�ُ�I�����܂����B", _
               vbOKOnly + vbCritical, _
               "���s����"
    End If
    
    Exit Sub '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    ' �����̃o�[�W��������\������
    Call psVersionDisp
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    '�u���[�N�N���A�ُ�I���v�|�b�v�A�b�v��ʕ\��
     MsgBox "�ُ�I�����܂����B", _
           vbOKOnly + vbCritical, _
           "���s����"
           
   '�u�o�[�W�����߂������ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_OLD_COPY_NOW_ERROR, lngErrCode)
           
    Set objFso = Nothing
    Set objFi = Nothing
    Set objFolder = Nothing
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : sVersionUpdate
'//  �@�\����  : �o�[�W�����A�b�v����
'//  �@�\�T�v  : ���[�N�o�[�W���������s�o�[�W�����֍X�V����B
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
'//  REVISIONS   �F(EG20 V5.8.0.1) 2012-04-15  CODED BY  [TCC] H.Sugimoto
'//                �y�w�E����No.02�C���Ή��z
'//                �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//  REVISIONS   �F(EG20 V5.13.0.1) 2012-06-02 REVISED BY[TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS   �F(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή�
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Sub sVersionUpdate()

    Dim objFso As New FileSystemObject      ' �t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                       ' �t�@�C���I�u�W�F�N�g
    Dim objFolder As Folder                 ' �t�H���_�I�u�W�F�N�g
    Dim stringWorkFolder As String          ' �t�H���_��
    
    Dim lngErrCode     As Long              ' �G���[�R�[�h
    Dim strSrcFile As String                ' �R�s�[��
    Dim strDstFile As String                ' �R�s�[��
    Dim bResult As Boolean                  ' ��������      ' EG20 V3.6.0.1�ǉ�

    On Error GoTo ErrorHandler          '�G���[�n���h���̓o�^

    ' /////////////////////////////////////////////////////////////////////////
    ' // ���[�N�t�H���_�̃t�@�C�����݃`�F�b�N
    stringWorkFolder = FLD_OPERATEPROGWRK
    If objFso.FolderExists(stringWorkFolder) <> True Then
        Set objFso = Nothing
        Set objFi = Nothing
        Set objFolder = Nothing
        ' �t�H���_�������������݂��Ȃ�
' EG20 V5.8.0.1�폜�J�n
'        MsgBox "���[�N�t�H���_���ɁA" _
'               & Chr(vbKeyReturn) & "�t�@�C�������݂��܂���B", _
'               vbOKOnly + vbExclamation, _
'               "���[�N�����s�R�s�["
' EG20 V5.8.0.1�폜�I��
' EG20 V5.8.0.1�ǉ��J�n
        MsgBox "�ُ�I�����܂����B", vbCritical, "���s����"
' EG20 V5.8.0.1�ǉ��I��
        Exit Sub                        ' �����I��
    End If

    strSrcFile = pfGetFileNameTakuExec
    ' �t�@�C���ݒ�Ȃ�
    If Len(strSrcFile) = 0 Then
        Set objFso = Nothing
        Set objFi = Nothing
        Set objFolder = Nothing
        MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
        Exit Sub
    End If

    strDstFile = stringWorkFolder & "\" & strSrcFile
    If objFso.FileExists(strDstFile) = False Then
        Set objFso = Nothing
        Set objFi = Nothing
        Set objFolder = Nothing
        ' �t�@�C�������݂��Ȃ�
' EG20 V5.8.0.1�폜�J�n
'        MsgBox "���[�N�t�H���_���ɁA" _
'               & Chr(vbKeyReturn) & "�t�@�C�������݂��܂���B", _
'               vbOKOnly + vbExclamation, _
'               "���[�N�����s�R�s�["
' EG20 V5.8.0.1�폜�I��
' EG20 V5.8.0.1�ǉ��J�n
        MsgBox "�ُ�I�����܂����B", vbCritical, "���s����"
' EG20 V5.8.0.1�ǉ��I��
        Exit Sub                        ' �����I��
    End If
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // ���t�H���_�̍폜
    '//////////////////////////////////////////////////////
    '// �Ď��Ճt�H���_���̑���싌�t�H���_������
    stringWorkFolder = FLD_OPERATEPROGOLD
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
    
    '//////////////////////////////////////////////////////
    '// �����t�H���_������
    stringWorkFolder = PATH_OPERATE_APLOLD
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
   
    ' /////////////////////////////////////////////////////////////////////////
    ' // ���[�N�����s�R�s�[

    '//////////////////////////////////////////////////////
    '// ���[�N�����s�R�s�[
    strSrcFile = FLD_OPERATEPROGNOW
    strDstFile = FLD_OPERATEPROGOLD

    '�t�H���_�I�u�W�F�N�g���擾
    Set objFolder = objFso.GetFolder(strSrcFile)
'    If objFolder.Size <> 0 Then                    ' EG20 V6.9.0.1 DEL
    If objFso.FolderExists(strSrcFile) Then         ' EG20 V6.9.0.1 ADD
        objFolder.Copy strDstFile
    End If

    ' /////////////////////////////////////////////////////
    ' // �S�̂��R�s�[
    strSrcFile = PATH_OPERATE_APL
    strDstFile = PATH_OPERATE_APLOLD

    '�t�H���_�I�u�W�F�N�g���擾
    Set objFolder = objFso.GetFolder(strSrcFile)
'    If objFolder.Size <> 0 Then                    ' EG20 V6.9.0.1 DEL
    If objFso.FolderExists(strSrcFile) Then         ' EG20 V6.9.0.1 ADD
        objFolder.Copy strDstFile
    End If
   
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // ���s�t�H���_�̍폜
    '//////////////////////////////////////////////////////
    '// �Ď��Ճt�H���_���̑������s�t�H���_������
    stringWorkFolder = FLD_OPERATEPROGNOW
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
    
    '//////////////////////////////////////////////////////
    '// �����t�H���_������
    stringWorkFolder = PATH_OPERATE_APL
    If objFso.FolderExists(stringWorkFolder) Then
        Call objFso.DeleteFolder(stringWorkFolder)
    End If
   
    ' /////////////////////////////////////////////////////////////////////////
    ' // ���[�N�����s�R�s�[

    '//////////////////////////////////////////////////////
    '// ���[�N�����s�R�s�[
    strSrcFile = FLD_OPERATEPROGWRK
    strDstFile = FLD_OPERATEPROGNOW

    '�t�H���_�I�u�W�F�N�g���擾
    Set objFolder = objFso.GetFolder(strSrcFile)
    objFolder.Copy strDstFile

    ' /////////////////////////////////////////////////////
    ' // �S�̂��R�s�[
    strSrcFile = PATH_OPERATE_APLNEW
    strDstFile = PATH_OPERATE_APL

    '�t�H���_�I�u�W�F�N�g���擾
    Set objFolder = objFso.GetFolder(strSrcFile)
    objFolder.Copy strDstFile

    Set objFso = Nothing
    Set objFi = Nothing
    Set objFolder = Nothing

' EG20 V3.6.0.1�ǉ��J�n
    ' �����v���O�����f�[�^�쐬����
    bResult = pfTakuProgramVersionCreateProc
' EG20 V3.6.0.1�ǉ��I��
    
    If bResult = True Then
' EG20 V5.8.0.1�폜�J�n
'        ' �^����ԍX�V                                              ' EG20 V5.5.0.1�ǉ�
'        Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1�ǉ�
' EG20 V5.8.0.1�폜�I��
' EG20 V5.8.0.1�ǉ��J�n
        ' �^����ԍX�V
        Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_TAKU, BOOTINFO_UNKAI_KIRIKAE)
        Call pubFuncAplUpdateUnkaiStatus(BOOTINFO_KEYNAMETAKU, BOOTINFO_UNKAI_KIRIKAE)
' EG20 V5.8.0.1�ǉ��I��
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        ' �����̃o�[�W��������\������
        Call psVersionDisp
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        '�u����I���v�|�b�v�A�b�v��ʕ\��
        MsgBox "����I�����܂����B", _
               vbOKOnly + vbInformation, _
               "���s����"
    Else
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        ' �����̃o�[�W��������\������
        Call psVersionDisp
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        '�u����I���v�|�b�v�A�b�v��ʕ\��
        MsgBox "�ُ�I�����܂����B", _
               vbOKOnly + vbInformation, _
               "���s����"
    End If
    
    Exit Sub '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    ' �����̃o�[�W��������\������
    Call psVersionDisp
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    '�u���[�N�N���A�ُ�I���v�|�b�v�A�b�v��ʕ\��
     MsgBox "�ُ�I�����܂����B", _
           vbOKOnly + vbCritical, _
           "���s����"
           
   '�u�o�[�W�����߂������ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, VERASION_INFO_WRK_COPY_NOW_BUTTOM, lngErrCode)
           
    Set objFso = Nothing
    Set objFi = Nothing
    Set objFolder = Nothing
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : pfGetFileNameTakuExec
'//  �@�\����  : �C���X�g�[�����s�t�@�C�����擾����
'//  �@�\�T�v  : �C���X�g�[�����s�t�@�C�������擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : String    �t�@�C����
'//
'//  ORIGINAL    �F(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS   �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  ���l        �F
'///////////////////////////////////////////////////////////////////
Private Function pfGetFileNameTakuExec() As String

    Const lngBufSize = MAX_PATH         ' �擾������̕������FID�A�f�[�^�p
    Dim strRet As String * MAX_PATH     ' �擾������
    Dim lngRet As Long                  ' �߂�l
    Dim szFileName As String            ' �t�@�C������
    Dim nNullIndex As Integer           ' ���������[�N
    
    pfGetFileNameTakuExec = ""
        
    'Ini�t�@�C��������s�t�@�C�������擾
    lngRet = GetPrivateProfileString(HOSHUINI_SECTION_OPERATE, HOSHUINI_OPERATEKEY_INSTEXEC, _
                                        "", strRet, lngBufSize, HOSHU_FILE)
    
    nNullIndex = InStr(strRet, Chr(0))
    If nNullIndex <> 0 Then
        szFileName = Left(strRet, nNullIndex - 1)
    Else
        szFileName = ""                 ' EG20 V3.3.0.1�폜
        szFileName = strRet             ' EG20 V3.3.0.1�ǉ�
    End If
    pfGetFileNameTakuExec = szFileName
    
End Function

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
    cmdClear.Enabled = blnFlg                       ' ���[�N�N���A
    cmdCopyBaitai_Work.Enabled = blnFlg             ' �}�́����[�N�R�s�[
    cmdCopyWork_Jikko.Enabled = blnFlg              ' ���[�N�����s�R�s�[
    cmdCopyOld_Jikko.Enabled = blnFlg               ' �������s�R�s�[
    cmdOutPut.Enabled = blnFlg                      ' �o�[�W�������}�̏o��
    CmdRemove.Enabled = blnFlg                      ' �}�̎�O
    cmdReturn.Enabled = blnFlg                      ' �o�[�W�����Ǘ���ʂ֖߂�

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'/
'/  �֐�����     : pfTakuProgramVersionCreateProc
'/  �@�\����     : �����v���O�����f�[�^�쐬����
'/  �@�\�T�v     : �����̎��s�o�[�W���������k���đ����f�[�^���쐬����B
'/
'/                 �^          ����            �Ӗ�
'/  ����         : �Ȃ�
'/  �߂�l       : Boolean     False           �ُ�I��
'/               :             True            ����I��
'/
'//  ORIGINAL    :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����TR-No.273�C���Ή��z
'/  REVISIONS    :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Function pfTakuProgramVersionCreateProc() As Boolean

    Dim sInputFolder As String                  ' �C���X�g�[�����t�H���_��
    Dim objFso As New FileSystemObject          ' �t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                           ' �t�@�C���I�u�W�F�N�g
    Dim MyName As String                        ' �t�@�C���t���p�X��
    Dim sSrcFileName As String                  ' �R�s�[���t�@�C����
    Dim strCabTarget As String                  ' ���k�Ώۃt�@�C��
    Dim lngRetZip As Long                       ' ���k����
    Dim objFolder As Folder                 ' �t�H���_�I�u�W�F�N�g

    Dim bResult As Long                     ' ��������

    On Error GoTo ErrorHandler                  ' �G���[�n���h���̓o�^

    pfTakuProgramVersionCreateProc = True
    sInputFolder = FLD_OPERATEPROGNOW
    strCabTarget = ""
    For Each objFi In objFso.GetFolder(sInputFolder).files      '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then            '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            '�}�̓��t�@�C�������쐬
            sSrcFileName = sInputFolder & "\" & MyName
            strCabTarget = strCabTarget & sSrcFileName & " "
        End If
    Next

   ' ���ׂẴf�B���N�g����񋓂���
    For Each objFolder In objFso.GetFolder(sInputFolder).SubFolders
        MyName = objFolder.Path
        strCabTarget = strCabTarget & MyName & " "
    Next


    lngRetZip = gsubCabZip(MELTED_TAKUVERSION, strCabTarget)
    
    If (lngRetZip <> 0) Then   '���k���ʂ�����(0)�ȊO
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LZH_ERROR, 0)
        Set objFso = Nothing
        Set objFi = Nothing
        pfTakuProgramVersionCreateProc = False
        Exit Function
    End If

    Set objFso = Nothing
    Set objFi = Nothing

    ' /////////////////////////////////////////////////////
    ' // �����v���O�����f�[�^�̍쐬
    bResult = dllCreateFile_TakuProgramData(1, MELTED_TAKUVERSION)
    If bResult = False Then
       pfTakuProgramVersionCreateProc = False
       Exit Function
    End If
    Exit Function

ErrorHandler:   ' �G���[�����B
    Set objFso = Nothing
    Set objFi = Nothing
    pfTakuProgramVersionCreateProc = False

End Function


