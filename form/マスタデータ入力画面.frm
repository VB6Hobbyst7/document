VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInputMstData 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExtMstInput 
      Caption         =   "   �O���}�X�^   ����      "
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   12
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton cmdUSBRemove 
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
      Height          =   975
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   4
      Top             =   3495
      Width           =   2415
   End
   Begin VB.CommandButton cmdMasterInput 
      Caption         =   "�}�̓���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   3
      Top             =   2415
      Width           =   2415
   End
   Begin VB.CommandButton cmdKoshin 
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
      Height          =   975
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdModoru_Menu 
      Caption         =   " �f�[�^���W�E�o�� ��ʂ֖߂�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   0
      Top             =   7080
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   15266
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   794
      TabMaxWidth     =   3475
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "   �������������@ ������������"
      TabPicture(0)   =   "�}�X�^�f�[�^���͉��.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tmrMail"
      Tab(0).Control(1)=   "dlgSelectFile"
      Tab(0).Control(2)=   "grdData(0)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "   �������������@ ������������"
      TabPicture(1)   =   "�}�X�^�f�[�^���͉��.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdData(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "   �������������@ ������������"
      TabPicture(2)   =   "�}�X�^�f�[�^���͉��.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdData(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "   �������������@ ������������"
      TabPicture(3)   =   "�}�X�^�f�[�^���͉��.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grdData(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "   �������������@ ������������"
      TabPicture(4)   =   "�}�X�^�f�[�^���͉��.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grdData(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "   �������������@ ������������"
      TabPicture(5)   =   "�}�X�^�f�[�^���͉��.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "grdData(5)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Timer tmrMail 
         Left            =   -74520
         Top             =   240
      End
      Begin MSComDlg.CommonDialog dlgSelectFile 
         Left            =   -73800
         Top             =   6240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
         Bindings        =   "�}�X�^�f�[�^���͉��.frx":00A8
         Height          =   4770
         Index           =   0
         Left            =   -74640
         TabIndex        =   6
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8414
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   8421504
         GridColor       =   12632256
         GridColorFixed  =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
         Bindings        =   "�}�X�^�f�[�^���͉��.frx":00BE
         Height          =   4770
         Index           =   1
         Left            =   -74640
         TabIndex        =   7
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8414
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   8421504
         GridColor       =   12632256
         GridColorFixed  =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
         Bindings        =   "�}�X�^�f�[�^���͉��.frx":00D4
         Height          =   4770
         Index           =   2
         Left            =   -74640
         TabIndex        =   8
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8414
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   8421504
         GridColor       =   12632256
         GridColorFixed  =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
         Bindings        =   "�}�X�^�f�[�^���͉��.frx":00EA
         Height          =   4770
         Index           =   3
         Left            =   -74640
         TabIndex        =   9
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8414
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   8421504
         GridColor       =   12632256
         GridColorFixed  =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
         Bindings        =   "�}�X�^�f�[�^���͉��.frx":0100
         Height          =   4770
         Index           =   4
         Left            =   -74640
         TabIndex        =   10
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8414
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   8421504
         GridColor       =   12632256
         GridColorFixed  =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
         Bindings        =   "�}�X�^�f�[�^���͉��.frx":0116
         Height          =   4770
         Index           =   5
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8414
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   8421504
         GridColor       =   12632256
         GridColorFixed  =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�}�X�^�f�[�^����"
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
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "frmInputMstData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 ALL Rights Reserved
'//
'//  �t�@�C����  �F�}�X�^�f�[�^����.frm
'//  �p�b�P�[�W���F�}�X�^�f�[�^���͉�ʂ̃t�H�[�����W���[��
'//
'//  �T�v�F�p�X���[�h���͉��
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-24  CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(EG20 V30.2.0.1) 2014-06-25  CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ��Q
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const DispKensu = 20                '�O���b�h�\���s��
Private Const GRID_TITLE = "<�@�@ �@�@|�@�@�@�@�@ �@ Ͻ����� �@�@�@�@�@�@|�@�ް�ޮ݁@|�@�@�@�@��M�����@�@�@�@"
Private Const MN_MAIL_INTERVAL = 1000   '���C���^�C�}�̃C���^�[�o���l
Private mlngHandle(19)          As Long     'EG20 V30.0.1.1 ADD


'EG20 V30.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : cmdExtMstInput_Click
'//  �@�\����  : �u�O���}�X�^���́v�t����������
'//  �@�\�T�v  : �O���}�X�^���͉�ʂ�\������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-02-24   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdExtMstInput_Click()

    '�u�}�X�^�f�[�^���͉�ʁF�O���}�X�^���͉����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, MASTER_INPUT_EXTMST_BUTTON, 0)
    
    '�O���}�X�^���͉�ʂ�\��
    Load frmExMasterInput
    frmExMasterInput.Show 1

End Sub
'EG20 V30.1.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdKoshin_Click
'//  �@�\����  : �u�\���X�V�v�t����������
'//  �@�\�T�v  : �}�X�^�f�[�^���ĕ\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-26  CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdKoshin_Click()

    '�u�}�X�^���͉�ʁF�\���X�V�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, MASTER_INPUT_KOUSHIN_BUTTOM, 0)
   
    Call sDisp_MasterData(SSTab1.Tab)
    Call sDisp_ParaData(SSTab1.Tab)     'EG20 V30.1.0.1 ADD
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdMasterInput_Click
'//  �@�\����  : �u�}�̓��́v�t����������
'//  �@�\�T�v  : �}�X�^�f�[�^���C���X�g�[������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V6.2.0.1) 2012-06-15 REVISED BY  [TCC] H.Sugimoto
'//                 �y���b�Z�[�W�{�b�N�X�{�^���R�[�h�s���Ή��z
'//     REVISIONS :(EG20 V6.5.0.1) 2012-06-18 REVISED BY  [TCC] H.Sugimoto
'//                 �y�t�@�C���̑I����@�����P�z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-05-09 REVISED BY  [TCC] T.Nakajima
'//                  �k���V�����J�ƑΉ�
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdMasterInput_Click()

    Dim iResponse As Integer            '���b�Z�[�W��������
    Dim strToPath As String             '�R�s�[��t�@�C���p�X
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim fso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    
    Dim szInputPath As String           ' �R�s�[���t�H���_�p�X      ' EG20 V6.5.0.1�ǉ�
    
    On Error Resume Next
    
    '�u�}�X�^���͉�ʁF�\���X�V�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, MASTER_INPUT_INSTALL_BUTTOM, 0)
    
    iResponse = MsgBox("���ݑI�𒆂̃R�[�i�̃}�X�^�f�[�^����͂��܂��B" & vbCrLf & "��낵���ł����H", _
                        vbOKCancel + vbQuestion, "�}�X�^�f�[�^���͊m�F")
    
    If iResponse = vbCancel Then
        Exit Sub
    End If
    
' EG20 V6.5.0.1�y�t�@�C���̑I����@�����P�z�폜�J�n
'    '�t�H���_�I�����
'    dlgSelectFile.FileName = ""
'    dlgSelectFile.Filter = "MST �t�@�C�� (*.MST)|*.MST|"
'    dlgSelectFile.DialogTitle = "�t�H���_���w�肵�Ă�������"
'    dlgSelectFile.InitDir = SHOWFOLDER_DEFAULTFOLDER
'    dlgSelectFile.Flags = dlgSelectFile.Flags Or cdlOFNNoChangeDir
'    dlgSelectFile.ShowOpen
'
'    '�w��t�H���_�Ȃ�
'    If Len(dlgSelectFile.FileName) = 0 Then
'         Exit Sub
'    End If
' EG20 V6.5.0.1�y�t�@�C���̑I����@�����P�z�폜�I��
' EG20 V6.5.0.1�y�t�@�C���̑I����@�����P�z�ǉ��J�n
    ' �t�@�C���I���������t�H���_�I������֕ύX���A
    ' �Œ�t�@�C�����ɑ΂��ď������s���B
    szInputPath = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)
    ' �w��t�H���_�Ȃ�
    If Len(szInputPath) = 0 Then
        Exit Sub
    End If
    'szInputPath = szInputPath & USB_MASTER_FILE    'EG20 V30.0.1.1 DEL
    'EG20 V30.1.0.1 ADD START
    If gintCornerType(SSTab1.Tab) = CORNER_TYPE_KANSEN Then
        szInputPath = szInputPath & USB_MASTER_FILE_KAN
    Else
        szInputPath = szInputPath & USB_MASTER_FILE
    End If
    'EG20 V30.1.0.1 ADD END
    If fso.FileExists(szInputPath) = False Then
        ' �R�s�[���Ƀ}�X�^�t�@�C�������݂��Ȃ��ꍇ�ُ͈�
        Call MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�X�^�f�[�^�X�V����")
        Set fso = Nothing
        Exit Sub
    End If
' EG20 V6.5.0.1�y�t�@�C���̑I����@�����P�z�ǉ��I��
        
    '��ʂ����b�N����
    Call sSetEnable(False)
    
    On Error GoTo Err_Handler
    
    strToPath = PATH_KANSI & "DESHU" & Format(SSTab1.Tab + 1, "00") & DIR_MASTER_V
    
    '�R�s�[��t�H���_�̗L���m�F
    If fso.FolderExists(strToPath) = False Then
        '�R�s�[��t�H���_�쐬
        fso.CreateFolder (strToPath)
    End If
    strToPath = strToPath & USB_MASTER_FILE
' EG20 V6.5.0.1�y�t�@�C���̑I����@�����P�z�폜�J�n
'    fso.CopyFile dlgSelectFile.FileName, strToPath, True
'    dlgSelectFile.InitDir = ""
'    dlgSelectFile.FileName = ""
' EG20 V6.5.0.1�y�t�@�C���̑I����@�����P�z�폜�I��
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z DEL Start
'' EG20 V6.5.0.1�y�t�@�C���̑I����@�����P�z�ǉ��J�n
'    fso.CopyFile szInputPath, strToPath, True
'' EG20 V6.5.0.1�y�t�@�C���̑I����@�����P�z�ǉ��I��
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z DEL End
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
    '�ꎞ�ۑ��t�H���_�Ƀf�[�^���R�s�[���ǎ��p����������
    If pfChangeAttrNormal(szInputPath, PATH_HOSHUTMP_MST_DATA, strToPath) = False Then
        '�ُ폈����
        GoTo Err_Handler
    End If
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End

    ChDir$ "C:\"
    Set fso = Nothing
    
    On Error Resume Next
    
    iResponse = MsgBox("���͂��ꂽ�}�X�^�f�[�^��K�p���܂��B" & vbCrLf & "��낵���ł����H", _
                        vbOKCancel + vbQuestion, "�}�X�^�f�[�^�K�p�m�F")
    
    If iResponse = vbCancel Then
        Call sSetEnable(True)
        Exit Sub
    End If
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_SHIMEKIRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '�}�X�^�X�V�v�����M
    If fCDATAMailSend = False Then
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

'        iResponse = MsgBox("�ُ�I�����܂����B", vbOK + vbCritical, "�}�X�^�f�[�^�X�V����")     ' EG20 V6.2.0.1�폜
        iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�X�^�f�[�^�X�V����")  ' EG20 V6.2.0.1�ǉ�
        Call sSetEnable(True)
        Exit Sub
    End If
    
    '��M�҂��B
    
    Exit Sub

Err_Handler:
    Set fso = Nothing
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
    '�ꎞ�ۑ��t�H���_���폜����
    psDeleteFolder PATH_HOSHUTMP
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End
    Call sSetEnable(True)
    '�u�}�X�^���͉�ʁF�}�X�^�f�[�^���ُ͈�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FWRITE
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, MASTER_INPUT_INSTALL_ERROR, lngErrCode)

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
'    iResponse = MsgBox("�ُ�I�����܂����B", vbOK + vbCritical, "�}�X�^�f�[�^���͌���")         ' EG20 V6.2.0.1�폜
    iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�X�^�f�[�^���͌���")      ' EG20 V6.2.0.1�ǉ�
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdUSBRemove_Click
'//  �@�\����  : �u�}�̎�O�v�t����������
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdUSBRemove_Click()

    On Error Resume Next
    
    '�u�}�X�^���͉�ʁF�}�̎�O�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, MASTER_INPUT_EJECT_BUTTOM, 0)
    
   '�}�̎�O����
    Call pfRemove(Me)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u���j���[��ʂɖ߂�v�t����������
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdModoru_Menu_Click()

    On Error Resume Next
    
   '�u�}�X�^�f�[�^���͉�ʁF�I���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, MASTER_INPUT_GAMEN_END, 0)
 
    '����ʂ������B
    Unload Me
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �}�X�^���͉��(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-26  CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    '�^�C�}���N������
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = True
    
    'EG20  V30.1.0.1 ADD START
    '�O���}�X�^���͉�ʂ���߂��Ă����Ƃ��ɂ��\���ł���悤��Activate�C�x���g�ŕ\���������s���悤�ɂ����B
    Call sDisp_MasterData(SSTab1.Tab)
    Call sDisp_ParaData(SSTab1.Tab)
    'EG20 V30.1.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �}�X�^���͉��(�f�B�A�N�e�B�u��:�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()

    '�^�C�}���~����
    tmrMail.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �}�X�^���͉��(���[�h���F�C�x���g�v���V�[�W��)
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-20  CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim intCount As Integer
    Dim bySyoAssort As Byte             '���O�p������
    Dim strCorner1 As String
    Dim strCorner2 As String
    
    On Error Resume Next
    
    Call gsGetSettiCorner
    Call gsGetCornerName
    Call gsGetCornerType        'EG20 V30.1.0.1 ADD

    '�^�u����ݒu�R�[�i���Ƃ���
    SSTab1.Tab = 0

    '�����t�@�C���G���[�̃g���b�v
    On Error GoTo Err_LOG
    bySyoAssort = L3AN_FILE
    
    For intCount = 0 To UBound(gblnCornerSet)
    
        '�ݒ肠��̃R�[�i�������ɂ���
        If gblnCornerSet(intCount) = True Then
            '�R�[�i�[���̕\��
            strCorner1 = MidB(gstrCornerName(intCount), 1, 12)
            strCorner2 = MidB(gstrCornerName(intCount), 13, 24)
            SSTab1.TabCaption(intCount) = strCorner1 & vbCrLf & strCorner2
        Else
            SSTab1.TabVisible(intCount) = False
        End If
    
    Next intCount
    
    'Call sDisp_MasterData(SSTab1.Tab)  'EG20 V30.1.0.1 DEL
                                        '����ǉ������O���}�X�^���͉�ʂ���߂��Ă����Ƃ��ɍĕ\���ł���悤Activate�C�x���g�Ɉړ�
    
    
    
Exit Sub

'�G���[����
Err_LOG:

    '�G���[���O�̏o��
     Call sLogTraceReq(LTYP_ERROR, bySyoAssort, MASTER_INPUT_GAMEN_START, 0)
     
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : SSTab1_Click
'//  �@�\����  : �^�u����������
'//  �@�\�T�v  : ���[����M�^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-02-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub SSTab1_Click(PreviousTab As Integer)
    '�I�𒆂̃^�u�C���f�b�N�X���Z�b�g�i�O���}�X�^���͉�ʂŕK�v�̂��߁j
    gintSelectedCornerTabIdx = SSTab1.Tab
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//
'//     �T�v      : �u���[����M�p�^�C�}�v���^�C���A�b�v�������̃C�x���g�v���V�[�W��
'//     ����      : ���[����M�������s���B
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF  '���[����M�G���A
    Dim lngLength As Long            '��M���[���o�C�g�T�C�Y
    Dim intStatus As Integer         '��M���[���`�F�b�N����
    Dim iResponse As Integer
    
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
                '�v���Z�X�̏I���������s��
                pfAbortProc
            Case ML_ID_HOSHU_ACTIVE_REQ
                '�u�ێ��ʃA�N�e�B�u�\���v����M�����ꍇ
                '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '�\������ʁi�ێ�f�[�^���W��ʁj���A�N�e�B�u�\������B
                AppActivate frmInputMstData.Caption, False
                pfFormActive (frmInputMstData.hwnd)           ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
            Case ML_ID_MASTER_UPDATE_RES
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
                '�v���O���X�o�[����������
                Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
                
                '�u�}�X�^�X�V�����ʒm�v����M�����ꍇ
                If fReadMailCheck(udtReadMail) = False Then
                    iResponse = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "�}�X�^�f�[�^�X�V����")
                Else
                    iResponse = MsgBox("����I�����܂����B", vbOKOnly + vbInformation, "�}�X�^�f�[�^�X�V����")
                    Call sDisp_MasterData(SSTab1.Tab)
                    Call sDisp_ParaData(SSTab1.Tab)     'EG20 V30.1.0.1 ADD
                End If
                Call sSetEnable(True)
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
'//  �֐�����  : sDisp_MasterData
'//  �@�\����  : �}�X�^�f�[�^�\������
'//  �@�\�T�v  : ���ݑI�𒆂̃R�[�i�̃}�X�^�t�@�C���f�[�^��\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   intTab    �I���^�u
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDisp_MasterData(ByVal intTab As Integer)

    Dim bySyoAssort         As Byte                 '���O�p������
    Dim intCorner           As Integer              '�R�[�i�ԍ��J�E���^
    Dim intCnt, intCnt2     As Integer              '�J�E���^
    Dim cFso                As FileSystemObject
    Dim cFile               As File
    Dim dtUpdate            As Date                 '�X�V����
    Dim strFilePath         As String               '�t�@�C���p�X
    Dim strFileName         As String               '�t�@�C����
    Dim intFileNo           As Integer              '�t�@�C���ԍ�
    Dim strNum              As String               '�}�X�^��
    Dim strNo()             As String               '�}�X�^�ԍ�
    Dim strMasterName()     As String               '�}�X�^����
    Dim strVer              As String               '�o�[�W����
    Dim intDataCnt          As Integer              '�f�[�^�J�E���^
    Dim intFileNumber       As Integer
    Dim intItemNum          As Integer
    Dim strDateTime         As String
    Dim byBuf()             As Byte
    Dim lngFileSize         As Long
    
    '�����t�@�C���G���[�̃g���b�v
    On Error GoTo Err_LOG
    bySyoAssort = L3AN_FILE
    
    '���g�p�̃t�@�C���ԍ����擾
    intFileNumber = FreeFile

    '�ݒ���t�@�C�����I�[�v������
    Open MASTER_DATA_NAME_FILE For Input As #intFileNumber
    
    For intCnt = 0 To 1
        Input #intFileNumber, strNum, strMasterName

        '�}�X�^����ݒ肷��
        If intCnt = 1 Then
            intItemNum = CInt(strNum)
        End If
    Next
    
    ReDim strNo(intItemNum - 1)
    ReDim strMasterName(intItemNum - 1)
    
    For intCnt = 0 To intItemNum - 1
        Input #intFileNumber, strNo(intCnt), strMasterName(intCnt)
    Next intCnt
    
    Close #intFileNumber
    
    grdData(intTab).Redraw = False      '�����ĕ`�����
    
    
    '�R�[�i�����[�v
    For intCorner = 0 To 5
        '�ݒu�R�[�i�̃f�[�^���擾����
        'If SSTab1.TabVisible(intCorner) = True Then    'EG20 V30.1.0.1 DEL
        'EG20 V30.1.0.1 ADD START
        '�ݒu�R�[�i���ݗ��R�[�i�̂݃f�[�^���擾����
        If (SSTab1.TabVisible(intCorner) = True) And (gintCornerType(intCorner) = CORNER_TYPE_ZAIRAI) Then
        'EG20 V30.1.0.1 ADD END
        
            '�t�H���_���w��
            strFilePath = PATH_KANSI & "DESHU" & Format(intCorner + 1, "00") & DIR_MASTER_V
            intDataCnt = grdData(intCorner).FixedRows
            
            '�O���b�h��������
            For intCnt = grdData(intCorner).FixedRows To grdData(intCorner).Rows - 2
                Call grdData(intCorner).RemoveItem(1)
            Next
            
            For intCnt = 0 To grdData(intCorner).Cols - 1
                grdData(intCorner).TextMatrix(1, intCnt) = ""
            Next
    
            intDataCnt = 1
             grdData(intCorner).FormatString = GRID_TITLE
            Set cFile = Nothing
            Set cFso = New FileSystemObject
            
            For intCnt = 0 To intItemNum - 1
                strFileName = Dir(strFilePath & "MASTER" & Format(strNo(intCnt), "00") & ".dat")
    
                If strFileName = Empty Then
                    strVer = ""
                    strDateTime = ""
                Else
                    '�t�@�C���̍X�V�������擾
                    Set cFile = cFso.GetFile(strFilePath & strFileName)
                    dtUpdate = cFile.DateLastModified
                    strDateTime = Format(dtUpdate, "yyyy�Nm��d��h��nn��")
            
                    lngFileSize = cFile.Size
                    ReDim byBuf(lngFileSize - 1)
            
                    intFileNo = FreeFile
                    '�t�@�C���I�[�v��
                    
                    Open strFilePath & strFileName For Binary As intFileNo Len = lngFileSize
            
                    '�f�[�^���t�@�C������ǂݍ���
                    Get #intFileNo, , byBuf
                
                    '�o�[�W�������擾
                    strVer = CStr(byBuf(3))
                    strVer = Format(strVer, "000")
                
                    Close #intFileNo
                End If
                    
                '�f�[�^�\��
                If intDataCnt > 0 Then
                    grdData(intCorner).AddItem ""
                End If
                grdData(intCorner).TextMatrix(intDataCnt, 0) = strNo(intCnt)
                grdData(intCorner).TextMatrix(intDataCnt, 1) = strMasterName(intCnt)
                grdData(intCorner).TextMatrix(intDataCnt, 2) = strVer
                grdData(intCorner).TextMatrix(intDataCnt, 3) = strDateTime
                intDataCnt = intDataCnt + 1
    
            Next intCnt
                    
            Call sSetRowFill(intCorner)
            '�\���ʒu�������ʒu��
            grdData(intCorner).TopRow = grdData(intCorner).FixedRows
        End If
        
        
    Next intCorner
    
    Set cFso = Nothing
    grdData(intTab).Redraw = True
    Me.Refresh
    
Exit Sub

Err_LOG:

    'EG20 V30.1.0.1 ADD START
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    'EG20 V30.1.0.1 ADD END
    
    If intFileNo > 0 Then
        Close #intFileNo
    End If
    
    Set cFso = Nothing
    grdData(intTab).Redraw = True
    '�G���[���O�̏o��
     Call sLogTraceReq(LTYP_ERROR, bySyoAssort, MASTER_INPUT_DISP_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : sDisp_ParaData
'//  �@�\����  : �p�����[�^�f�[�^�\������
'//  �@�\�T�v  : ���ݑI�𒆂̃R�[�i�̃}�X�^�t�@�C���f�[�^��\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   intTab    �I���^�u
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-02-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sDisp_ParaData(ByVal intTab As Integer)

    Dim bySyoAssort         As Byte                 '���O�p������
    Dim intCorner           As Integer              '�R�[�i�ԍ��J�E���^
    Dim intCnt, intCnt2     As Integer              '�J�E���^
    Dim cFso                As FileSystemObject
    Dim cFile               As File
    Dim dtUpdate            As Date                 '�X�V����
    Dim strFilePath         As String               '�t�@�C���p�X
    Dim strFileName         As String               '�t�@�C����
    Dim intFileNo           As Integer              '�t�@�C���ԍ�
    Dim strNum              As String               '�}�X�^��
    Dim strNo()             As String               '�}�X�^�ԍ�
    Dim strParaName()       As String               '�}�X�^����
    Dim strParaFile()       As String               '�p�����[�^�f�[�^�t�@�C����
    Dim strVer              As String               '�o�[�W����
    Dim intDataCnt          As Integer              '�f�[�^�J�E���^
    Dim intFileNumber       As Integer
    Dim intItemNum          As Integer
    Dim strDateTime         As String
    Dim byBuf()             As Byte
    Dim lngFileSize         As Long
    Dim uParaFoot           As PARA_FOOT            '�p�����[�^�f�[�^�̃t�b�^��
    Dim i                   As Integer
    Dim intMuIdx            As Integer
    Dim strMutexFile        As String
    
    '�����t�@�C���G���[�̃g���b�v
    On Error GoTo Err_LOG
    bySyoAssort = L3AN_FILE
    
    '���g�p�̃t�@�C���ԍ����擾
    intFileNumber = FreeFile

    '�ݒ���t�@�C�����I�[�v������
    Open PARA_DATA_NAME_FILE For Input As #intFileNumber
    
    For intCnt = 0 To 1
        Input #intFileNumber, strNum, strParaName, strParaFile

        '�}�X�^����ݒ肷��
        If intCnt = 1 Then
            intItemNum = CInt(strNum)
        End If
    Next
    
    intMuIdx = 0
    Erase mlngHandle
    
    ReDim strNo(intItemNum - 1)
    ReDim strParaName(intItemNum - 1)
    ReDim strParaFile(intItemNum - 1)
    
    For intCnt = 0 To intItemNum - 1
        Input #intFileNumber, strNo(intCnt), strParaName(intCnt), strParaFile(intCnt)
    Next intCnt
    
    Close #intFileNumber
    
    grdData(intTab).Redraw = False      '�����ĕ`�����
    
    
    '�R�[�i�����[�v
    For intCorner = 0 To 5
        '�ݒu�R�[�i�������R�[�i�̃f�[�^���擾����
        intMuIdx = 0
        If (SSTab1.TabVisible(intCorner) = True) And (gintCornerType(intCorner) = CORNER_TYPE_KANSEN) Then
        
            '�t�H���_���w��
            strFilePath = PATH_KANSI & "N_GATE" & Format(intCorner + 1, "00") & DIR_NPARA_V
            intDataCnt = grdData(intCorner).FixedRows
            
            '�O���b�h��������
            For intCnt = grdData(intCorner).FixedRows To grdData(intCorner).Rows - 2
                Call grdData(intCorner).RemoveItem(1)
            Next
            
            For intCnt = 0 To grdData(intCorner).Cols - 1
                grdData(intCorner).TextMatrix(1, intCnt) = ""
            Next
    
            intDataCnt = 1
             grdData(intCorner).FormatString = GRID_TITLE
            Set cFile = Nothing
            Set cFso = New FileSystemObject
            
            For intCnt = 0 To intItemNum - 1
                strFileName = Dir(strFilePath & strParaFile(intCnt))
    
                If strFileName = Empty Then
                    strVer = ""
                    strDateTime = ""
                Else
                    '�r������(OPEN)
                    strMutexFile = "MU_PARAMETER" & Format(intCorner + 1, "00")
                    mlngHandle(intMuIdx) = dllOpenMutex(strMutexFile)
                    If mlngHandle(intMuIdx) <> 0 Then
                        dllWaitForSingleObject (mlngHandle(intMuIdx))     '�r������(GET)
                    End If
                    
                    '�t�@�C���̍X�V�������擾
                    Set cFile = cFso.GetFile(strFilePath & strFileName)
                    dtUpdate = cFile.DateLastModified
                    strDateTime = Format(dtUpdate, "yyyy�Nm��d��h��nn��")
            
                    lngFileSize = cFile.Size
                    ReDim byBuf(lngFileSize - 1)
            
                    intFileNo = FreeFile
                    '�t�@�C���I�[�v��
                    'Open strFilePath & strFileName For Binary As intFileNo Len = lngFileSize
                    'Binary�ŊJ���ꍇ��Len�߂͈Ӗ��Ȃ��B�i�T�C�Y��32,767 �o�C�g�ȉ��ł���K�v������B�������p�����[�^�͂���ȏ�j
                    Open strFilePath & strFileName For Binary As intFileNo
            
                    '�p�����[�^�f�[�^�̃t�b�^�����擾����
                    Get #intFileNo, lngFileSize - Len(uParaFoot) + 1, uParaFoot
            
                    '�o�[�W�������擾
                    strVer = ""
                    For i = 0 To UBound(uParaFoot.byVersion)
                        strVer = strVer & Right$("0" & Hex(uParaFoot.byVersion(i)), 2)
                    Next i
                    strVer = Format(strVer, "000")

                    Close #intFileNo
                    
                    If mlngHandle(intMuIdx) <> 0 Then
                        '�r������(FREE)
                        Call dllReleaseMutex(mlngHandle(intMuIdx))
                        '�r������(CLOSE)
                        Call dllCloseHandle(mlngHandle(intMuIdx))
                    End If
                    intMuIdx = intMuIdx + 1
                End If
                    
                '�f�[�^�\��
                If intDataCnt > 0 Then
                    grdData(intCorner).AddItem ""
                End If
                grdData(intCorner).TextMatrix(intDataCnt, 0) = strNo(intCnt)
                grdData(intCorner).TextMatrix(intDataCnt, 1) = strParaName(intCnt)
                grdData(intCorner).TextMatrix(intDataCnt, 2) = strVer
                grdData(intCorner).TextMatrix(intDataCnt, 3) = strDateTime
                intDataCnt = intDataCnt + 1
    
            Next intCnt
                    
            Call sSetRowFill(intCorner)
            '�\���ʒu�������ʒu��
            grdData(intCorner).TopRow = grdData(intCorner).FixedRows
        End If
        
        
    Next intCorner
    
    Set cFso = Nothing
    grdData(intTab).Redraw = True
    Me.Refresh
    
Exit Sub

Err_LOG:

    If intFileNumber > 0 Then
        Close #intFileNumber
    End If

    If intFileNo > 0 Then
        Close #intFileNo
    End If
    
    For intCnt = 0 To intMuIdx - 1
        '�r�������iFREE)
        Call dllReleaseMutex(mlngHandle(intCnt))
        '�r������(CLOSE)
        Call dllCloseHandle(mlngHandle(intCnt))
    Next
  
    
    Set cFso = Nothing
    grdData(intTab).Redraw = True
    '�G���[���O�̏o��
     Call sLogTraceReq(LTYP_ERROR, bySyoAssort, MASTER_INPUT_DISP_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : sSetRowFill
'//  �@�\����  : �O���b�h�s���ߏ���
'//  �@�\�T�v  : �O���b�h�̍s���Ɣw�i�F��ݒ�
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   intTab    �I���^�u
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sSetRowFill(ByVal intTab As Integer)

    Dim intCnt, intCnt2 As Integer
    
    '�s�����P�y�[�W�̕\�������ɂȂ�悤�ɁA�󔒍s���쐬����B
    If (grdData(intTab).Rows - grdData(intTab).FixedRows) Mod DispKensu > 0 Then
        grdData(intTab).Rows = grdData(intTab).Rows + (DispKensu - (grdData(intTab).Rows - grdData(intTab).FixedRows) Mod DispKensu)
    End If
    
    '�O���b�h�̍s�w�i�F��ݒ�
    grdData(intTab).RowHeight(0) = 232
    For intCnt = 1 To (grdData(intTab).Rows - 1)
        grdData(intTab).Row = intCnt
        grdData(intTab).RowHeight(intCnt) = 232
        For intCnt2 = 0 To grdData(intTab).Cols - 1
            grdData(intTab).Col = intCnt2
            If (intCnt Mod 2) = 0 Then
            '�����s�̔w�i�F�́uFFFFFF�v
                grdData(intTab).CellBackColor = "&H00FFFFFF"
             Else
                '��s�̔w�i�F�́uDFDDDE�v
                grdData(intTab).CellBackColor = "&H00DFDDDE"
            End If
        Next intCnt2
    Next intCnt
        
    grdData(intTab).Redraw = True               '�����ĕ`��ĊJ
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : fCDATAMailSend
'//  �@�\����  : �}�X�^�X�V�v�����M����
'//  �@�\�T�v  : �����������F���[���𑗐M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-04   CODED   BY [TCC] M.Matsumoto
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fCDATAMailSend() As Boolean

    Dim udtMail As MAIL_INFO_RES    '�}�X�^�X�V�v�����[�����M�G���A
    Dim lngRet As Long              '�֐��߂�l
    Dim lngErrCode As Long          '�G���[�R�[�h
    
    On Error Resume Next
 
    '�}�X�^�X�V�v�����ă}�ɑ��M����B
    udtMail.mlHeader.dwId = ML_ID_MASTER_UPDATE_CMD
    udtMail.mlHeader.dwSize = MlSize.MASTER_UPDATE_REQ
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    udtMail.dwRequestType = ML_ID_MASTER_UPDATE_H       '�f�[�^���
    udtMail.dwSts = SSTab1.Tab + 1                      '�R�[�i�ԍ�
    
    lngRet = DssSendMail(MAIL_SLOT_KANMA, Len(udtMail), udtMail.mlHeader)
    If lngRet = False Then
       '�u�}�X�^�f�[�^���͉�ʁF�}�X�^�X�V�v�����M�ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, MASTER_UPDATE_REQ_SEND, lngErrCode)
       fCDATAMailSend = False
       Exit Function
    Else
       '�u�}�X�^�f�[�^���͉�ʁF�}�X�^�X�V�v�����M����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, MASTER_UPDATE_REQ_SEND, 0)
       fCDATAMailSend = True
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : fReadMailCheck
'//  �@�\����  : �}�X�^�X�V�����ʒm���[���`�F�b�N����
'//  �@�\�T�v  : ���[����M���F���[������M����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-20   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fReadMailCheck(udtReadMail As ML_KYOTU_INF) As Boolean

    Dim lngErrCode As Long
    
    On Error Resume Next
    
    '�f�[�^��ʃ`�F�b�N
    If udtReadMail.lngData(0) <> ML_ID_MASTER_UPDATE_H Then
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MRECEIVE + 1
        Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MASTER_UPDATE_REQ_RECV, lngErrCode)
        fReadMailCheck = False
        Exit Function
    End If
    
    '�R�[�i�`�F�b�N
    If udtReadMail.lngData(1) <> (SSTab1.Tab + 1) Then
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MRECEIVE + 2
        Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MASTER_UPDATE_REQ_RECV, lngErrCode)
        fReadMailCheck = False
        Exit Function
    End If
    
    '�������ʃ`�F�b�N
    If udtReadMail.lngData(2) > 0 Then
        fReadMailCheck = False
        Exit Function
    End If
    
    fReadMailCheck = True
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : sSetEnable
'//  �@�\����  : ������Ԑ���
'//  �@�\�T�v  : �R�}���h�{�^���̊����E�񊈐��𐧌䂷��
'//
'//              �^        ����      �Ӗ�
'//  ����      : Boolean   blnEnable [IN]�����^�s�����t���O
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-10-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V30.2.0.1) 2014-06-25  REVISED BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ��Q
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sSetEnable(ByVal blnEnable As Boolean)

    Dim lngErrCode As Long
    
    On Error Resume Next
    
    cmdKoshin.Enabled = blnEnable
    cmdMasterInput.Enabled = blnEnable
    cmdUSBRemove.Enabled = blnEnable
    cmdModoru_Menu.Enabled = blnEnable
    cmdExtMstInput.Enabled = blnEnable      'EG20 V30.2.0.1 ADD
    SSTab1.Enabled = blnEnable
    
End Sub
