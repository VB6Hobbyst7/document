VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmKansenGateVerKanri 
   BackColor       =   &H00800000&
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
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
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
      Height          =   550
      Left            =   9360
      TabIndex        =   70
      Top             =   3240
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrMail 
      Left            =   8640
      Top             =   1560
   End
   Begin VB.CommandButton cmdGateVerUpdate 
      Caption         =   "�ꊇ�X�V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   62
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "�o�[�W�������  �}�̏o��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Index           =   0
      Left            =   9360
      TabIndex        =   20
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton cmdInstall 
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
      Height          =   550
      Index           =   1
      Left            =   9360
      TabIndex        =   19
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton cmdModoru_Menu 
      Caption         =   "  �o�[�W�����Ǘ�  ��ʂ֖߂�"
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
      TabIndex        =   1
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Frame fraDataSelect 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   7080
      Width           =   6255
      Begin VB.OptionButton optData 
         Caption         =   "�\���R"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   3960
         TabIndex        =   61
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "�\���Q"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   3960
         TabIndex        =   60
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "�\���P"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   3960
         TabIndex        =   59
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "�\���P"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2040
         TabIndex        =   17
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "�\���Q"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   2040
         TabIndex        =   18
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "�n�r"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2040
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "�T�u�b�o�t"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "���C���b�o�t"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optData 
         Caption         =   "����b�o�t"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame fraFolderSelect 
      Height          =   1335
      Left            =   6720
      TabIndex        =   8
      Top             =   7080
      Width           =   1935
      Begin VB.CheckBox chkFolder 
         Caption         =   "�n ��"
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "�m ���s"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   615
         Width           =   1575
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "�v ���[�N"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
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
      Height          =   550
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   6
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopyBaitai_Work 
      Caption         =   " ���k�t�@�C�� �� ���[�N�R�s�["
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   5
      Top             =   2640
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
      Height          =   550
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   4
      Top             =   3840
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
      Height          =   550
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   3
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdDLLJikkoGamen 
      Caption         =   " �����؂藣��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   2
      Top             =   5520
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
      Height          =   550
      Left            =   9360
      Style           =   1  '���̨���
      TabIndex        =   7
      Top             =   600
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   21
      Top             =   360
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   15266
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   706
      TabMaxWidth     =   3475
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "   �������������@ ������������"
      TabPicture(0)   =   "�o�[�W�����Ǘ�(��������).frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblKan(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblKan(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblKan(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblKan(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblKan(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblKan(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblZenVer(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lstKan(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "   �������������@ ������������"
      TabPicture(1)   =   "�o�[�W�����Ǘ�(��������).frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstKan(1)"
      Tab(1).Control(1)=   "Command1(1)"
      Tab(1).Control(2)=   "lblKan(22)"
      Tab(1).Control(3)=   "lblKan(16)"
      Tab(1).Control(4)=   "lblKan(14)"
      Tab(1).Control(5)=   "lblKan(8)"
      Tab(1).Control(6)=   "lblKan(7)"
      Tab(1).Control(7)=   "lblKan(5)"
      Tab(1).Control(8)=   "lblZenVer(1)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "   �������������@ ������������"
      TabPicture(2)   =   "�o�[�W�����Ǘ�(��������).frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstKan(2)"
      Tab(2).Control(1)=   "lblZenVer(2)"
      Tab(2).Control(2)=   "lblKan(23)"
      Tab(2).Control(3)=   "lblKan(21)"
      Tab(2).Control(4)=   "lblKan(20)"
      Tab(2).Control(5)=   "lblKan(19)"
      Tab(2).Control(6)=   "lblKan(18)"
      Tab(2).Control(7)=   "lblKan(17)"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "   �������������@ ������������"
      TabPicture(3)   =   "�o�[�W�����Ǘ�(��������).frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstKan(3)"
      Tab(3).Control(1)=   "lblZenVer(3)"
      Tab(3).Control(2)=   "lblKan(31)"
      Tab(3).Control(3)=   "lblKan(29)"
      Tab(3).Control(4)=   "lblKan(28)"
      Tab(3).Control(5)=   "lblKan(27)"
      Tab(3).Control(6)=   "lblKan(26)"
      Tab(3).Control(7)=   "lblKan(25)"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "   �������������@ ������������"
      TabPicture(4)   =   "�o�[�W�����Ǘ�(��������).frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lstKan(4)"
      Tab(4).Control(1)=   "lblZenVer(4)"
      Tab(4).Control(2)=   "lblKan(39)"
      Tab(4).Control(3)=   "lblKan(37)"
      Tab(4).Control(4)=   "lblKan(36)"
      Tab(4).Control(5)=   "lblKan(35)"
      Tab(4).Control(6)=   "lblKan(34)"
      Tab(4).Control(7)=   "lblKan(33)"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "   �������������@ ������������"
      TabPicture(5)   =   "�o�[�W�����Ǘ�(��������).frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lstKan(5)"
      Tab(5).Control(1)=   "lblZenVer(5)"
      Tab(5).Control(2)=   "lblKan(47)"
      Tab(5).Control(3)=   "lblKan(45)"
      Tab(5).Control(4)=   "lblKan(44)"
      Tab(5).Control(5)=   "lblKan(43)"
      Tab(5).Control(6)=   "lblKan(42)"
      Tab(5).Control(7)=   "lblKan(41)"
      Tab(5).ControlCount=   8
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
         Height          =   4380
         Index           =   1
         Left            =   -74640
         TabIndex        =   75
         Top             =   2280
         Width           =   8055
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
         Height          =   4380
         Index           =   2
         Left            =   -74640
         TabIndex        =   74
         Top             =   2280
         Width           =   8055
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
         Height          =   4380
         Index           =   3
         Left            =   -74640
         TabIndex        =   73
         Top             =   2280
         Width           =   8055
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
         Height          =   4380
         Index           =   4
         Left            =   -74640
         TabIndex        =   72
         Top             =   2280
         Width           =   8055
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
         Height          =   4380
         Index           =   5
         Left            =   -74640
         TabIndex        =   71
         Top             =   2280
         Width           =   8055
      End
      Begin VB.CommandButton Command1 
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
         Height          =   550
         Index           =   1
         Left            =   -65640
         Style           =   1  '���̨���
         TabIndex        =   69
         Top             =   2280
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
         Height          =   4380
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   2280
         Width           =   8055
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�R�����g"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   -71880
         TabIndex        =   81
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�쐬���t"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   -68880
         TabIndex        =   80
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ް����{�ް�ޮ�"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   -71400
         TabIndex        =   79
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   -71880
         TabIndex        =   78
         Top             =   1680
         Width           =   495
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
         Height          =   615
         Index           =   7
         Left            =   -72720
         TabIndex        =   77
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�t�@�C����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   -74640
         TabIndex        =   76
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblZenVer 
         Alignment       =   1  '�E����
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
         Height          =   1215
         Index           =   5
         Left            =   -73920
         TabIndex        =   68
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label lblZenVer 
         Alignment       =   1  '�E����
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
         Height          =   1215
         Index           =   4
         Left            =   -73920
         TabIndex        =   67
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label lblZenVer 
         Alignment       =   1  '�E����
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
         Height          =   1215
         Index           =   3
         Left            =   -73920
         TabIndex        =   66
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label lblZenVer 
         Alignment       =   1  '�E����
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
         Height          =   1215
         Index           =   2
         Left            =   -73920
         TabIndex        =   65
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label lblZenVer 
         Alignment       =   1  '�E����
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
         Height          =   1215
         Index           =   1
         Left            =   -73920
         TabIndex        =   64
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label lblZenVer 
         Alignment       =   1  '�E����
         Caption         =   "�������������o�[�W�����i���[�N�j�F99"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   1080
         TabIndex        =   63
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�R�����g"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   47
         Left            =   -71880
         TabIndex        =   58
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�쐬���t"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   45
         Left            =   -68880
         TabIndex        =   57
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ް����{�ް�ޮ�"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   44
         Left            =   -71400
         TabIndex        =   56
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   43
         Left            =   -71880
         TabIndex        =   55
         Top             =   1680
         Width           =   495
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
         Height          =   615
         Index           =   42
         Left            =   -72720
         TabIndex        =   54
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�t�@�C����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   41
         Left            =   -74640
         TabIndex        =   53
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�R�����g"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   -71880
         TabIndex        =   52
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�쐬���t"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   37
         Left            =   -68880
         TabIndex        =   51
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ް����{�ް�ޮ�"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   36
         Left            =   -71400
         TabIndex        =   50
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   35
         Left            =   -71880
         TabIndex        =   49
         Top             =   1680
         Width           =   495
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
         Height          =   615
         Index           =   34
         Left            =   -72720
         TabIndex        =   48
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�t�@�C����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   33
         Left            =   -74640
         TabIndex        =   47
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�R�����g"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   -71880
         TabIndex        =   46
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�쐬���t"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   -68880
         TabIndex        =   45
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ް����{�ް�ޮ�"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   -71400
         TabIndex        =   44
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   -71880
         TabIndex        =   43
         Top             =   1680
         Width           =   495
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
         Height          =   615
         Index           =   26
         Left            =   -72720
         TabIndex        =   42
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�t�@�C����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   25
         Left            =   -74640
         TabIndex        =   41
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�R�����g"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   -71880
         TabIndex        =   40
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�쐬���t"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   -68880
         TabIndex        =   39
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ް����{�ް�ޮ�"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   -71400
         TabIndex        =   38
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   -71880
         TabIndex        =   37
         Top             =   1680
         Width           =   495
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
         Height          =   615
         Index           =   18
         Left            =   -72720
         TabIndex        =   36
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�t�@�C����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   17
         Left            =   -74640
         TabIndex        =   35
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�R�����g"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   -71880
         TabIndex        =   34
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�쐬���t"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   -68880
         TabIndex        =   33
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ް����{�ް�ޮ�"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   -71400
         TabIndex        =   32
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   -71880
         TabIndex        =   31
         Top             =   1680
         Width           =   495
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
         Height          =   615
         Index           =   10
         Left            =   -72720
         TabIndex        =   30
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�t�@�C����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   -74640
         TabIndex        =   29
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�t�@�C����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   1680
         Width           =   1935
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
         Height          =   615
         Index           =   1
         Left            =   2280
         TabIndex        =   27
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3120
         TabIndex        =   26
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ް����{�ް�ޮ�"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3600
         TabIndex        =   25
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�쐬���t"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6120
         TabIndex        =   24
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�R�����g"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   23
         Top             =   2040
         Width           =   5295
      End
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�V�����������D�@�o�[�W�����Ǘ�"
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
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   12120
   End
End
Attribute VB_Name = "frmKansenGateVerKanri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmJGateVerKanri.frm
'//  �p�b�P�[�W���F�o�[�W�����Ǘ�(EG20����)���
'//
'//  �T�v�F�o�[�W�����Ǘ�(EG-R����/NEG����)���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�v�����������`�F�b�N�����ǉ�
'//     REVISIONS :(1.6.0.1) 2009-06-11   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή�
'//                     �E�@�퐳�����`�F�b�N�����ǉ�/�u���[�N�����s�R�s�[�v��
'//                     �E�t�F�[�Y�Q�s��C��
'//                     �E�t�F�[�Y�P�s��C��
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.11.0.1) 2009-10-23  CODED   BY [TCC] D.Yamashita
'//                 �E�t�F�[�Y�R�c�����ڑΉ�
'//     REVISIONS :(1.12.0.1) 2009-11-11  REVISED BY [TCC] C.Terui
'//                 ���̓t�@�C���i�[�f�B���N�g���ʒu�ύX
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//     REVISIONS :(1.16.0.1) 2009-12-21   REVISED BY [TCC] S.Terao
'//                 �s��Ή�
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                �@ �t�H���_�I����ʂ�OS�d�l�ɕύX
'//                �A�u���j���[��ʂ֖߂�v�t�����ɂāA
'//                 �@�o�[�W�����Ǘ���ʂ̃o�[�W�����\���X�V���s��
'//                �B�\�����\�[�X���W�I�t�I���Ń��X�g�̕\���X�V
'//                �C���[�N�����s�R�s�[�ł̋@�퐳�����`�F�b�N�ύX
'//                �D���[�N�����s�R�s�[�ł̐������`�F�b�Nini�t�@�C����
'//                �EDir�֐���FileSystemObject�ɒu������
'//                �F�t�@�C���I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(V2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 EG-R(KK)�@������Ή��@KUK�������`�F�b�N�ύX
'//                 �}�̎�O�s��C��
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 �t�@�C�����`�F�b�N�s��C��
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-16  REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-07  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                �y�^���\�����P�Ή��z
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-03  CODED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y��ʃ`�F�b�N�@�\�ǉ��z�yTOMAS�p�̈�R�s�[�Ή��z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Dim FolderSyubetu As Integer                 '�I�����\�[�X���

Dim FolderName(0 To 2, 0 To 9) As String     '�t�H���_��
Dim TitleBox(0 To 10) As String               '�^�C�g����
Dim LogBox(0 To 10) As String                 '���O�o�͗p�^�C�g����
Dim FileList() As String                     '�t�@�C�������X�g�ꗗ�i�[�G���A
Dim FileListType() As String                 '�t�@�C�����X�g�ꗗ�i�[�G���A�i�����㎩���^�C�v���܂ށj
'Dim uVersion() As MN_VERSION_JIKAI           '�o�[�W�������i�[�G���A      'EG20 V30.1.0.1 DEL
Dim uVersion() As MN_VERSION_KAN_JIKAI       '�o�[�W�������i�[�G���A      'EG20 V30.1.0.1 ADD
Dim gintUnkaiKind(0 To 8) As Integer         ' �^�����    ' EG20 V5.11.0.1�ǉ�
Dim gintProgramJudgeKind(0 To 8) As Integer  ' �v���O����������    ' EG20 V6.9.0.1�y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zADD

'�I�𒆃��\�[�X��� =0=MN_RSOC_PRO�F�v���O�����A=1=MN_RSOC_HAN:����f�[�^
Dim iSelResource As Integer


Private Const MN_MAIL_INTERVAL = 1000       '���[���^�C�}�̃C���^�[�o���l

Private Const MN_FOLD_WRK = 0               '�u���[�N�v�t�H���_
Private Const MN_FOLD_NOW = 1               '�u���s�v�t�H���_
Private Const MN_FOLD_OLD = 2               '�u���v�t�H���_

'�o�[�W�����f�[�^�t�@�C���p�̍\����
Private Type MN_VERSION_FILE
    sFileName As String * 12                '�t�@�C����
    uFooter As MN_FOOT_BYTE                 '�t�b�^���
End Type

Private Type MN_VERSION_DAT
    strFolder(0 To 5) As String * 8         '�t�H���_��
    intFileNum(0 To 5) As Integer           '�t�@�C����
End Type
'�o�[�W�����f�[�^�t�@�C�����(�o�[�W����2)
Private Type MN_FILE_INFO_V2
    udtInfo As MN_VERSION_DAT               '�t�H���_���ƃt�@�C����
    uFileInfo() As MN_VERSION_FILE          '�t�@�C�����ƃt�b�^���
End Type

Dim uVerdataFile As MN_FILE_INFO_V2

Private Const HANKUKA_KUK = "HAN_KUKA.KUK"
Private Const INI_MAX = 5
Dim HAN_KUKA_DATA As HANTEI_DATA
Private Type HANTEI_DATA
    sHederKisyu(0 To 4) As String
    sHederFile(0 To 4) As String
    sFotterKisyu(0 To 4) As String
    sFotterFile(0 To 4) As String
End Type

'V1.4.0.1�@ADD�@START
Private Const FILE_NAME_MAX_SIZE = 12
Private Const FILE_NAME_SIZE = 19
'�y�^���f�[�^�������`�F�b�N�ُ�X�e�[�^�X��`�z
Private sNGSts As String        'NG�ʒu
Private sNGKoumoku As String    'NG����
'�yNG�ʒu�z
Private Const ERROR_HEDER = "�w�b�_"  '�w�b�_
Private Const ERROR_FOTTER = "�t�b�^" '�t�b�^
'�yNG���ځz
Private Const KISHU_NAME_ERROR = "�@�햼"       '�@�햼
Private Const FILE_NAME_ERRORE = "�t�@�C����"   '�t�@�C����
Private Const CREATE_DATA_ERROR = "�쐬���t"    '�쐬���t
Private Const VERSION_ERROR = "�o�[�W����"      '�o�[�W����
Private sJverName As String                     '�\�����b�Z�[�W�{�b�N�X�^�C�g��
'Private Const EG20_JIKAI = "EG20"               'EG20       'EG20 V30.1.0.1 DEL
Private Const EG30_JIKAI = "EG30"               'EG30        'EG20 V30.1.0.1 ADD
'V1.4.0.1�@ADD�@END
'V1.6.0.1 ADD START
Private Const EGR_JIKAI_KISHU = "EG5000"        'EG-R�����@�햼
Private Const NEG_JIKAI_KISHU = "EG2000"        'NEG�����@�햼
Private Const EG20_JIKAI_KISHU = "EG6000"       'EG20 �����@�햼
Private Const EG30_JIKAI_KISHU = "EG7000"       'EG30 �����@�햼
'V1.20.0.1 DEL START
'EG-R����
'Private Const EHANTEI_CPU_CHK_FILE = "ko_gateh.vef"
'Private Const EMAIN_CPU_CHK_FILE = "ko_gatep.vef"
'Private Const ESUB_CPU_CHK_FILE = "ko_gatef.vef"
'Private Const EMAIN_OS_CHK_FILE = "ko_gateo.vef"
''NEG����
'Private Const NHANTEI_CPU_CHK_FILE = "KO_GATEH.VEF"
'Private Const NMAIN_CPU_CHK_FILE = "KO_GATEP.VEF"
'Private Const NSUB_CPU_CHK_FILE = "KO_GATEF.VEF"
'Private Const NMAIN_OS_CHK_FILE = "KO_GATEO.VEF"
'V1.20.0.1 DEL END
'EG20 V30.1.0.1 DEL START
'V1.20.0.1 ADD START
'EG-R����
'Private EHANTEI_CPU_CHK_FILE As String
'Private EMAIN_CPU_CHK_FILE As String
'Private ESUB_CPU_CHK_FILE As String
'Private EMAIN_OS_CHK_FILE As String
''NEG����
'Private NHANTEI_CPU_CHK_FILE As String
'Private NMAIN_CPU_CHK_FILE As String
'Private NSUB_CPU_CHK_FILE As String
'Private NMAIN_OS_CHK_FILE As String
'V1.20.0.1 ADD END
'EG20 V30.1.0.1 DEL END
'EG20 V30.1.0.1 ADD START
'�V��������
Private EG30_HANTEI_CPU_CHK_FILE As String
Private EG30_MAIN_CPU_CHK_FILE As String
Private EG30_SUB_CPU_CHK_FILE As String
Private EG30_MAIN_OS_CHK_FILE As String
'EG20 V30.1.0.1 ADD END
'V1.6.0.1 ADD END
'�f�[�^��ʑI��
' EG20 V2.0.1.1 ADD START
Public mlngOptDataType          As Long

'�t�H���_��ʕ�
Public mlngChkFolderType        As Long

Dim mbVerKanriExecuteFlg                      As Boolean  '�o�͎��s���������ۂ�

Private iTab_index As Integer       '�@�I�𒆂̃R�[�i�[�ԍ�
' EG20 V2.0.1.1 ADD END

' EG20 V3.0.0.2�ǉ��J�n
Private Const TITLEDISP_VERNOTHING = "--"       ' ��ʏ㕔�o�[�W�����Ȃ��\��
Private Const TITLEDISP_FIXEDVERNOW = "                      �i���s�j  �F"
Private Const TITLEDISP_FIXEDVEROLD = "                      �i���j    �F"

Dim DispTitleBox(0 To 10) As String             ' ��ʏ㕔�^�C�g�����i�P�s�ځj
Dim DispTitleVersion(0 To 2) As String          ' ��ʏ㕔�o�[�W����

' EG20 V3.0.0.2�ǉ��I��

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdGateVerUpdate_Click
'//  �@�\����  : �ꊇ�X�V�t��������
'//  �@�\�T�v  : ���D�@�ꊇ�X�V��ʂ�\������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//  ORIGINAL  :(EG20 V2.1.0.1) 2011-09-25  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//                 EG20�����Ď���USDM�Ή��ԍ��yMainte_03_01�z
'//  REVISIONS :(EG20 V30.1.0.1) 2014-02-20  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdGateVerUpdate_Click()

    '�u�����ް�ޮ݁F�����؂藣���t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_IKKATSU_BUTTOM, 0)

    '�ʐM�ڑ��E�ؒf��ʂ�\������B
    'Load frmGateVerUpdate          'EG20 V30.1.0.1 DEL
    Load frmKansenGateVerUpdate     'EG20 V30.1.0.1 ADD
    'frmGateVerUpdate.Show 1        'EG20 V30.1.0.1 DEL
    frmKansenGateVerUpdate.Show 1   'EG20 V30.1.0.1 ADD

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �o�[�W�����Ǘ�(EG20����)���(�A�N�e�B�u��)
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
    
    '���[����M�p�̃^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �o�[�W�����Ǘ�(EG20����)���(�f�B�A�N�e�B�u��)
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
    
    If blnCabfrmOpenFlg = True Then
        Call fnTsbCabCallDiverge
        Exit Sub
    End If

    '���[����M�p�̃^�C�}���~�߂�B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdInstall_Click
'//  �@�\����  : �u�}�̎�O�v�t����������
'//  �@�\�T�v  : �}�̂̎�O�����s��
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
Private Sub cmdInstall_Click(Index As Integer)
   On Error Resume Next
   
   If Index = 1 Then                                ' �}�̎�O ����
       '�u�}�̎�O�t�����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
        '�}�̎�O����
        Call pfRemove(Me)
    Else                                            '�o�[�W�������  �}�̏o�͏���
        '�u�����ް�ޮ݁F�}�̏o�͖t�����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_OUTPUT_BUTTOM, 0)
 
        '�}�̏o�͏���
        fMakeOutPutFile
    End If
    
ErrorHandle:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Command1_Click
'//  �@�\����  : �u�}�́����[�N�R�s�[�v�t����������
'//  �@�\�T�v  : �}�̂����[�N�ɃR�s�[
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] T.koyama
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Command2_Click()

   Dim iResponse As Integer         'MsgBox�{�^���R�[�h
   Dim lngErrCode As Long           '�G���[�R�[�h

   On Error Resume Next

   '�u�����ް�ޮ݁F�}�́����[�N�R�s�[�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_LZH_COPY_WRK_BUTTOM, 0)
    '�C���X�g�[���}�̂����[�N�t�H���_���ɃR�s�[����
    sFDInstall "STD"
        
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2005 All Rights Reserved
'/
'/  �֐�����     : Form_Load
'/  �@�\����     : Form_Load������
'/  �@�\�T�v     : Form_Load���������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/ ORIGINAL  :(3.1.0.1) 2005-11-29   CODED   BY [TCC] A.Mizuno
'/ REVISIONS :(5.1.0.1) 2006-05-10   CODED   BY [TCC] K.Hayashi
'/ REVISIONS :(5.3.0.1) 2006-06-08   CODED   BY [TCC] K.Hayashi
'/ REVISIONS :(EG20 V2.0.1.1) 2011-11-18   CODED   BY [TCC] T.Koyama
'/ REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'/             EG20�t�F�[�Y�Q�Ή�
'/ REVISIONS :(EG20 V3.4.0.1) 2012-02-07  CODED BY  [TCC] H.Sugimoto
'/             EG20�t�F�[�Y�Q�Ή��i1�R�[�i�ݒ�Ő������\�����s���Ȃ��Ή��j
'/ REVISIONS :(EG20 V30.1.0.1) 2014-02-17 CODED BY  [TCC] T.Nakajima
'/             �k���V�����J�ƑΉ�
'/ REVISIONS :(EG20 V30.3.0.1) 2014-10-16 CODED BY  [TCC] T.Nakajima
'/             �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'/ REVISIONS :(x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()

   Dim intCount As Integer
   Dim strCorner1 As String
   Dim strCorner2 As String
   Dim bySelectedFlg    As Byte     'EG20 V30.1.0.1 ADD
   
   On Error Resume Next
 
    'sJverName = EG20_JIKAI     'EG20 V30.1.0.1 DEL
    sJverName = EG30_JIKAI      'EG20 V30.1.0.1 ADD
    
    'EG20 V30.1.0.1 DEL START
    '�uEG-R�������D�@�ް�ޮ݉�ʁF�\���v���O�o��
    'Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EJIKAI_VERASION_KANRI_GAMEN_START, 0)
    'EG20 V30.1.0.1 DEL END
    'EG20 V30.1.0.1 ADD START
    '�uEG-R�������D�@�ް�ޮ݉�ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KJIKAI_VERASION_KANRI_GAMEN_START, 0)
    'EG20 V30.1.0.1 ADD END
  
 ' EG20 V2.0.1.1 ADD START�y�c����60�z
    ' �t�H���_�I���`�F�b�N�{�b�N�X�����l�ݒ�
    For intCount = 0 To chkFolder.UBound
      chkFolder(intCount) = 1
    Next intCount
      
      '���@���擾
    Call gsGetGateInfo
    Call gsGetCornerName
    Call gsGetCornerType        ''EG20 V30.1.0.1 ADD
    
   '�^�u����ݒu�R�[�i���Ƃ���
    SSTab1.Tab = 0
'    SSTab1.Tabs = gintCornerNum            ' EG20 V3.4.0.1 �폜
    bySelectedFlg = False       'EG20 V30.1.0.1 ADD
    For intCount = 0 To UBound(gblnCornerSet)
    
        '�ݒ肠��̃R�[�i�������ɂ���
        If gblnCornerSet(intCount) = True Then
            '�R�[�i�[���̕\��
            strCorner1 = MidB(gstrCornerName(intCount), 1, 12)
            strCorner2 = MidB(gstrCornerName(intCount), 13, 24)
            SSTab1.TabCaption(intCount) = strCorner1 & vbCrLf & strCorner2
            'EG20 V30.0.3.1 �yHKRK_Kansi06_004_02�z DEL START
            'EG20 V30.1.0.1 ADD START
'            If gintCornerType(intCount) = CORNER_TYPE_ZAIRAI Then
'                '�ݗ��R�[�i�Ȃ�Ή����s�ɂ���
'                SSTab1.TabEnabled(intCount) = False
'            Else
'                '��Ԏn�߂̐V�����R�[�i�[�̃^�u��I����Ԃɂ���B
'                If bySelectedFlg = False Then
'                    SSTab1.Tab = intCount
'                    bySelectedFlg = True
'                    '�V�����̐擪�R�[�i�[�Ȃ��GATE00�ɃR�s�[������K�v�����邽�߁A�擪�C���f�b�N�X��ۑ����Ă���
'                    gintKansenFirstCornerIdx = intCount
'                End If
'            End If
            'EG20 V30.1.0.1 ADD END
            'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL END
            
            '���X�g�{�b�N�X������������
            lstKan(intCount).Clear
        
            '��ʃ^�C�g���ݒ�
            'lbltitle(intCount).Caption = "�������D�@�o�[�W�����Ǘ�"    ' EG20 V30.1.0.1 DEL
            lbltitle(intCount).Caption = "�V�����������D�@�o�[�W�����Ǘ�"    ' EG20 V30.1.0.1 ADD
   
' EG20 V3.0.0.2�폜�J�n
'            '��\�o�[�W�����ݒ�
'            lblZenVer(intCount).Caption = "����f�[�^�@�o�[�W�����i���[�N�j�F  " & vbCrLf & _
'                                          "                      �i���s�j  �F  " & vbCrLf & _
'                                          "                      �i���j    �F  "
' EG20 V3.0.0.2�폜�I��
        End If
    Next

    '�ݒ�Ȃ��̃R�[�i�^�u���\���ɐݒ肷��
    For intCount = 0 To UBound(gblnCornerSet)

        If gblnCornerSet(intCount) = False Then
            SSTab1.TabVisible(intCount) = False
        End If
    Next
 ' EG20 V2.0.1.1 ADD END  �y�c����60�z

    '�f�[�^�W�J
    sSetFolderName

    '�ϐ��̏�����
    FolderSyubetu = 0

    '�o�[�W�������̃��X�g�{�b�N�X���쐬����
    fMakeListbox

    '���[����M�p�̃��[����M�p�̃^�C�}�l��ݒ肷��
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000

End Sub

  
'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  �֐�����     : chkFolder_Click
'/  �@�\����     : �u�t�H���_�I�𕔁v�`�F�b�N����
'/  �@�\�T�v     : �u�t�H���_�I�𕔁v�`�F�b�N�������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub chkFolder_Click(Index As Integer)
  
'    Dim ValueCnt                As Integer
'
'    '���O�o��
'    If Index = 0 Then
'        '���[�N
'        Call psPutLog(LOG_frmGateVerKanri_CHKFOLDER0)
'    ElseIf Index = 1 Then
'        '���s
'        Call psPutLog(LOG_frmGateVerKanri_CHKFOLDER1)
'    ElseIf Index = 2 Then
'        '��
'        Call psPutLog(LOG_frmGateVerKanri_CHKFOLDER2)
'    End If
'
'    '��ނɂ���đ����l��ύX����
'    ValueCnt = 0
'    '���[�N
'    If Index = 0 Then
'        ValueCnt = 1
'    '���s
'    ElseIf Index = 1 Then
'        ValueCnt = 2
'    '��
'    ElseIf Index = 2 Then
'        ValueCnt = 4
'    End If
'
'    '�`�F�b�N���͂����ꂽ��
'    If chkFolder(Index).Value = 0 Then
'        mlngChkFolderType = mlngChkFolderType - ValueCnt
'    '�`�F�b�N���ꂽ��
'    ElseIf chkFolder(Index).Value = 1 Then
'        mlngChkFolderType = mlngChkFolderType + ValueCnt
'    End If

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  �֐�����     : cmdClear_Click
'/  �@�\����     : �u���[�N�N���A�v�{�^����������
'/  �@�\�T�v     : �u���[�N�N���A�v�{�^�������������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub CmdClear_Click()

   Dim iResponse As Integer         'MsgBox�{�^���R�[�h
   Dim lngErrCode As Long           '�G���[�R�[�h

   On Error Resume Next

    '�u�����ް�ޮ݊Ǘ��F���[�N�N���A�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_WRK_CREA_BUTTOM, 0)

    '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
    iResponse = MsgBox("�u���[�N�v�t�H���_���� " & TitleBox(FolderSyubetu) & "���A" _
           & Chr(vbKeyReturn) & "�S�č폜���܂��B    ��낵���ł����H", _
           vbYesNo + vbExclamation, _
           TitleBox(FolderSyubetu) & "  ���[�N �N���A")
    If iResponse = vbYes Then
        '[�͂�] �{�^����I�������ꍇ
        '���[�N�t�H���_���̃t�@�C�����폜����
       If sWrkFolderRemove <> True Then
          '�u�����ް�ޮ݊Ǘ��F���[�N�N���A�����ُ�v���O�o��
          lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_CREA_ERROR, lngErrCode)
          Exit Sub
       End If
       '�u�����ް�ޮ݊Ǘ��F���[�N�N���A��������v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_CREA_OK, 0)
       
       '���X�g�{�b�N�X������������
       lstKan(0).Clear
       lstKan(1).Clear
       lstKan(2).Clear
       lstKan(3).Clear
       lstKan(4).Clear
       lstKan(5).Clear
       
       '�o�[�W������񃊃X�g�{�b�N�X���쐬����
       fMakeListbox
    End If
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  �֐�����     : cmdCopyBaitai_Work_Click
'/  �@�\����     : �u�}��(���k)�����[�N �R�s�[�v�{�^����������
'/  �@�\�T�v     : �u�}��(���k)�����[�N �R�s�[�v�{�^�������������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyBaitai_Work_Click()

   On Error Resume Next

    '�u�����ް�ޮ݁F���ķ�ف�ܰ���߰�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_CAB_COPY_WRK_BUTTOM, 0)

    '���k�t�@�C������C���X�g�[������B
    sFDInstall "LZH"
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  �֐�����     : cmdCopyOld_Jikko_Click
'/  �@�\����     : �u�������s �R�s�[�v�{�^����������
'/  �@�\�T�v     : �u�������s �R�s�[�v�{�^�������������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'/                 �y�v���O���X�o�[�\���@�\�������Ή��z
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyOld_Jikko_Click()
    
   Dim iResponse As Integer         'MsgBox�{�^���R�[�h
   Dim lngErrCode As Long           '�G���[�R�[�h

   On Error Resume Next

   '�u�����ް�ޮ݁F�������s�R�s�[�t�����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_OLD_COPY_NOW_BUTTOM, 0)
   '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
   iResponse = MsgBox("�u���v�t�H���_�̓��e���A�u���s�v�t�H���_�ɖ߂����Ƃɂ��A" _
             & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & "�̈ꐢ��O�̃o�[�W�������A" _
             & Chr(vbKeyReturn) & "���s�o�[�W�����Ƃ��܂��B  ��낵���ł����H", _
            vbYesNo + vbExclamation, _
            TitleBox(FolderSyubetu) & "  �������s �R�s�[")
   If iResponse = vbYes Then
   '[�͂�] �{�^����I�������ꍇ
         
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
         
         '�ꐢ��O�̃o�[�W���������s�o�[�W�����ɖ߂�
       If fOldVersion <> True Then
          '�u�����ް�ޮ݁F���[�N�����s�R�s�[�����ُ�v���O�o��
          lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OLD_COPY_NOW_ERROR, lngErrCode)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
           '�v���O���X�o�[����������
           Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
          Exit Sub
       End If
       '�u�����ް�ޮ݁F�������s�R�s�[��������v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_OLD_COPY_NOW_OK, 0)
       
       '���X�g�{�b�N�X������������
       lstKan(0).Clear
       lstKan(1).Clear
       lstKan(2).Clear
       lstKan(3).Clear
       lstKan(4).Clear
       lstKan(5).Clear
      
       '�o�[�W������񃊃X�g�{�b�N�X���쐬����
       fMakeListbox
   
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
   End If
       
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  �֐�����     : cmdCopyWork_Jikko_Click
'/  �@�\����     : �u���[�N�����s �R�s�[�v�{�^����������
'/  �@�\�T�v     : �u���[�N�����s �R�s�[�v�{�^�������������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-27   CODED   BY [TCC] Y.Masuda
'/  REVISIONS    : (3.10.0.1) 2006-02-02  CODED   BY [TCC] K.Inoue
'/  REVISIONS    : (EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'/                 �y�v���O���X�o�[�\���@�\�������Ή��z
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCopyWork_Jikko_Click()
   
   Dim iResponse As Integer         'MsgBox�{�^���R�[�h
   Dim lngErrCode As Long           '�G���[�R�[�h

   On Error Resume Next

   '�u���[�N�����s�R�s�[�v�{�^���̏ꍇ�B
   '�u�����ް�ޮ݁F���[�N�����s�R�s�[�t�����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_WRK_COPY_NOW_BUTTOM, 0)
    
   '�m�F�|�b�v�A�b�v�E�B���h�E��\������B
   iResponse = MsgBox("�u���[�N�v�t�H���_�̓��e���A�u���s�v�t�H���_�ɓo�^���邱�Ƃɂ��A" _
            & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & " �̍ŐV�̃o�[�W�������A���s�o�[�W�����Ƃ��܂��B" _
            & Chr(vbKeyReturn) & "��낵���ł����H", _
           vbYesNo + vbExclamation, _
           TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�[")
   If iResponse = vbYes Then
   '[�͂�] �{�^����I�������ꍇ
            
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[��\������
        Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
            '�ŐV�o�[�W���������s�o�[�W�����Ƃ��ēo�^����
        If fNewVersion <> True Then
           '�u�����ް�ޮ݁F���[�N�����s�R�s�[�����ُ�v���O�o��
           lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_COPY_NOW_ERROR, lngErrCode)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
           '�v���O���X�o�[����������
           Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
           Exit Sub
        End If
        '�u�����ް�ޮ݁F���[�N�����s�R�s�[��������v���O�o��
         Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_WRK_COPY_NOW_OK, 0)
        
        '���X�g�{�b�N�X������������
        lstKan(0).Clear
        lstKan(1).Clear
        lstKan(2).Clear
        lstKan(3).Clear
        lstKan(4).Clear
        lstKan(5).Clear
        
        '�o�[�W������񃊃X�g�{�b�N�X���쐬����
        fMakeListbox
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
   End If
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  �֐�����     : cmdDLLJikkoGamen_Click
'/  �@�\����     : �uDLL���s��ʂցv�{�^����������
'/  �@�\�T�v     : �uDLL���s��ʂցv�{�^�������������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdDLLJikkoGamen_Click()

    Dim i As Integer                '�J�E���^
    Dim bFlag As Boolean            '�t���O
    Dim lRetVal As Long             '�߂�l
    Dim sCommand As String          '�R�}���h������
    Dim sWriteDir As String
    
    On Error GoTo ErrorHandle

    '�u�����ް�ޮ݁F�����؂藣���t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAISATU_VERSION_KANRI_KIRIHANASI_BUTTOM, 0)

    '�ʐM�ڑ��E�ؒf��ʂ�\������B
    Load frmConectSts
    frmConectSts.Show 1

ErrorHandle:
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  �֐�����     : cmdKoshin_Click
'/  �@�\����     : �u�\���X�V�v�{�^����������
'/  �@�\�T�v     : �u�\���X�V�v�{�^�������������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdKoshin_Click()
    
    Dim i As Integer                '�J�E���^
    Dim bFlag As Boolean            '�t���O
    Dim lRetVal As Long             '�߂�l
    Dim sCommand As String          '�R�}���h������
    Dim sWriteDir As String
    
    On Error GoTo ErrorHandle

    '�u�����ް�ޮ݁F�\���X�V�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UPDATE_BUTTOM, 0)

    '�t�H���_�I�𕔂Ɏw��L���`�F�b�N
    bFlag = False                                 '�t���O���u�U�v�ɂ���
    For i = 0 To 2                                '�t�H���_�����J��Ԃ�
        If chkFolder(i).Value = CHECKBOX_ON Then   '�u�H�H�v�t�H���_���w�肳��Ă���
            bFlag = True                            '�t���O���u�^�v�ɂ���
            Exit For                                '���[�v�𔲂���
        End If
    Next
              
    If bFlag = False Then                       '�t�H���_�w�薳��
        '�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
        'EG20 V30.1.0.1 DEL START
'        MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
'                          vbOKOnly + vbExclamation, _
'                          "�������D�@ �o�[�W�����Ǘ�"
        'EG20 V30.1.0.1 DEL END
        'EG20 V30.1.0.1 ADD START
        MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
                          vbOKOnly + vbExclamation, _
                          "�V�����������D�@ �o�[�W�����Ǘ�"
        'EG20 V30.1.0.1 ADD END
        '�����𔲂���
        Exit Sub
    End If
    
    '���X�g�{�b�N�X������������
    lstKan(0).Clear
    lstKan(1).Clear
    lstKan(2).Clear
    lstKan(3).Clear
    lstKan(4).Clear
    lstKan(5).Clear
    
    '�o�[�W������񃊃X�g�{�b�N�X���쐬����
    fMakeListbox
              
ErrorHandle:
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  �֐�����     : cmdModoru_Menu_Click
'/  �@�\����     : ���j���[��ʂɖ߂�{�^����������
'/  �@�\�T�v     : ���j���[��ʂɖ߂�{�^�������������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-07   CODED   BY [TCC] T.Shimizu
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdModoru_Menu_Click()
    
'    '���O�o��
'    Call psPutLog(LOG_frmGateVerKanri_CMDMODORU_MENU)
'
'    '���j���[��ʕ\��
'    frmProgramHanteiData.Show

    '��ʂ�Unload
    Unload Me

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  �֐�����     : sOptDataDisp
'/  �@�\����     : �f�[�^��ʑI�𕔕\������
'/  �@�\�T�v     : �f�[�^��ʑI�𕔂�I�����ꂽ�^�u�ʂɕ\���������s��
'/
'/                 �^          ����                   �Ӗ�
'/  ����         : Long        ����IC-M���[�J�[�I�� �N���b�N�����^�u�C���f�b�N�X(1�`6)
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-26   CODED   BY [TCC] Y.Masuda
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub sOptDataDisp(Index As Long)
'
'    '�f�[�^��ʑI�𕔕\��
'    Dim intCnt                  As Long
'
'    '�f�[�^��ʑI�𕔂��ĕ\������
'    For intCnt = 5 To 0 Step -1
'        If gudtMaker(Index).strType(intCnt) = "" Then
'            Me.optData(intCnt).Caption = ""
'            Me.optData(intCnt).Visible = False
'        Else
'            Me.optData(intCnt).Caption = gudtMaker(Index).strType(intCnt)
'            Me.optData(intCnt).Visible = True
'
'            'Ver1.0.0.6 ADD Start
'            mlngOptDataType = intCnt + 1
'            'Ver1.0.0.6 ADD End
'        End If
'    Next
'
'    'Ver1.0.0.6 UPD Start
'    '�I����Ԃɂ���
'    Me.optData(mlngOptDataType - 1).Value = True
'    'Ver1.0.0.6 UPD End


End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  �֐�����     : sCmdBtnEnabled
'/  �@�\����     : �R�}���h�{�^�������E�s����
'/  �@�\�T�v     : �R�}���h�{�^���������Ɋ�ĉ����E�s�������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-26   CODED   BY [TCC] Y.Masuda
'/  REVISIONS    : (1.0.0.5) 2005-04-06   CODED   BY [TCC] Y.Masuda
'/  REVISIONS    : (3.1.0.1) 2005-12-09   CODED   BY [TCC] A.Mizuno
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub sCmdBtnEnabled(blnFlg As Boolean)
'
'    '���ׂĂ̖t�����\
'    Me.cmdClear.Enabled = blnFlg
'    Me.cmdCopyBaitai_Work.Enabled = blnFlg
'    If blnFlg = vbTrue Then
'      Call CopyBtm_Enabled
'    Else
'      Me.cmdCopyOld_Jikko.Enabled = blnFlg
'      Me.cmdCopyWork_Jikko.Enabled = blnFlg
'    End If
'    Me.cmdDLLJikkoGamen.Enabled = blnFlg
'    Me.cmdKoshin.Enabled = blnFlg
'    Me.cmdModoru_Menu.Enabled = blnFlg
''V3.1.0.1 Add Start
'    'DLL����ʂ̃{�^������ǉ�
'    Me.cmdDLLKyokaGamen.Enabled = blnFlg
''V3.1.0.1 Add End
    
End Sub

Private Sub Form_Paint()
'    glnghwndTabCnt = gudtVerTbsInfo.lnghwndTabCnt
'    glnghwndOwnDrwTab1 = gudtVerTbsInfo.lnghwndOwnDrwTab
'    glngPrevWndProc = gudtVerTbsInfo.lngPrevWndProc
'    tbsICMVersion.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    ' �T�u�N���X���J�n
'    UnSubClass Me, gudtVerTbsInfo.lngPrevWndProc
End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2004 All Right Reserved
'/
'/  �֐�����     : optData_Click
'/  �@�\����     : �u�f�[�^��ʑI�𕔁v��������
'/  �@�\�T�v     : �u�f�[�^��ʑI�𕔁v�����������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.0.0) 2004-12-26   CODED   BY [TCC] Y.Masuda
'/  REVISIONS    : (EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'/                  EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx   CODED   BY [xxx]
'/  ���l         :
'/////////////////////////////////////////////////////////////////////////////
Private Sub optData_Click(Index As Integer)
  
    Dim i As Integer                '�J�E���^
    Dim bFlag As Boolean            '�t���O

    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
    
    '���\�[�X��ʂ�ύX����B'
    FolderSyubetu = Index
    
' EG20 V3.0.0.2�폜�J�n
'    ' EG20 V2.0.1.1 ADD START�y�c����60�z
'    Select Case FolderSyubetu           '���\�[�X���
'        Case 0                              '����f�[�^
'           lblZenVer(iTab_index).Caption = "����f�[�^  �o�[�W�����i���[�N�j�F" & vbCrLf & _
'                                           "                      �i���s�j  �F" & vbCrLf & _
'                                           "                      �i���j    �F"
'        Case 1                              '�v���O����
'           lblZenVer(iTab_index).Caption = "�v���O����  �o�[�W�����i���[�N�j�F" & vbCrLf & _
'                                           "                      �i���s�j  �F" & vbCrLf & _
'                                           "                      �i���j    �F"
'        Case 2                              '���CPU-Pro1
'           lblZenVer(iTab_index).Caption = "���CPU-Pro1 �o�[�W�����i���[�N�j�F" & vbCrLf & _
'                                           "                      �i���s�j  �F" & vbCrLf & _
'                                           "                      �i���j    �F"
'        Case 3                              '���CPU-Pro2
'           lblZenVer(iTab_index).Caption = "���CPU-Pro2 �o�[�W�����i���[�N�j�F" & vbCrLf & _
'                                           "                      �i���s�j  �F" & vbCrLf & _
'                                           "                      �i���j    �F"
'        Case 4                              '���CPU-Pro3
'           lblZenVer(iTab_index).Caption = "���CPU-Pro3 �o�[�W�����i���[�N�j�F" & vbCrLf & _
'                                           "                      �i���s�j  �F" & vbCrLf & _
'                                           "                      �i���j    �F"
'        Case 5                              '�����i�n�r�j
'           lblZenVer(iTab_index).Caption = "�����i�n�r�j�o�[�W�����i���[�N�j�F" & vbCrLf & _
'                                           "                      �i���s�j  �F" & vbCrLf & _
'                                           "                      �i���j    �F"
'        Case 6                              '�\���P
'           lblZenVer(iTab_index).Caption = "�\���P      �o�[�W�����i���[�N�j�F" & vbCrLf & _
'                                           "                      �i���s�j  �F" & vbCrLf & _
'                                           "                      �i���j    �F"
'        Case 7                              '�\���Q
'           lblZenVer(iTab_index).Caption = "�\���Q      �o�[�W�����i���[�N�j�F" & vbCrLf & _
'                                           "                      �i���s�j  �F" & vbCrLf & _
'                                           "                      �i���j    �F"
'        Case 8                              '�\���P
'           lblZenVer(iTab_index).Caption = "�\���R      �o�[�W�����i���[�N�j�F" & vbCrLf & _
'                                           "                      �i���s�j  �F" & vbCrLf & _
'                                           "                      �i���j    �F"
'    End Select
'    ' EG20 V2.0.1.1 ADD START�y�c����60�z
' EG20 V3.0.0.2�폜�I��

    
    
    
    '�t�H���_�I�𕔂Ɏw��L���`�F�b�N
    bFlag = False                                 '�t���O���u�U�v�ɂ���
    For i = 0 To 2                                '�t�H���_�����J��Ԃ�
        If chkFolder(i).Value = CHECKBOX_ON Then   '�u�H�H�v�t�H���_���w�肳��Ă���
            bFlag = True                            '�t���O���u�^�v�ɂ���
            Exit For                                '���[�v�𔲂���
        End If
    Next
    
    If bFlag = False Then                       '�t�H���_�w�薳��
        '�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
        'EG20 V30.1.0.1 DEL START
'        MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
'                    vbOKOnly + vbExclamation, _
'                    "�������D�@ �o�[�W�����Ǘ�"
        'EG20 V30.1.0.1 DEL END
        'EG20 V30.1.0.1 ADD START
        MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
                    vbOKOnly + vbExclamation, _
                    "�V�����������D�@ �o�[�W�����Ǘ�"
        'EG20 V30.1.0.1 ADD END
        '�����𔲂���
        Exit Sub
    End If
    
    '���X�g�{�b�N�X������������
    lstKan(0).Clear
    lstKan(1).Clear
    lstKan(2).Clear
    lstKan(3).Clear
    lstKan(4).Clear
    lstKan(5).Clear
    
    '�o�[�W������񃊃X�g�{�b�N�X���쐬����
    fMakeListbox
    'V1.20.0.1 ADD END

End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sVersionDisp
'//  �@�\����  : �o�[�W������񃊃X�g�{�b�N�X�ǉ�
'//  �@�\�T�v  : �o�[�W���������t�@�C�����P�ʂŃ��X�g�{�b�N�X�ɒǉ�����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.4.0.1) 2012-06-17 REVISED BY [TCC] H.Sugimoto
'//                �y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-18 REVISED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Sub sVersionDisp(uVerData() As MN_VERSION_JIKAI)       'EG20 V30.1.0.1 DEL
Private Sub sVersionDisp(uVerData() As MN_VERSION_KAN_JIKAI)    'EG20 V30.1.0.1 ADD
    Dim sFileName As String         '�t�@�C����������i�����㎩���^�C�v���܂ށj
    Dim sFileSize As String         '�t�@�C���T�C�Y������
    Dim sFileInfo(2) As String      '�o�[�W������񕶎���
    Dim sComment1(2) As String      '�R�����g������
    Dim sComment2(2) As String      '�R�����g������

   On Error Resume Next
    
    If uVerData(0).sFileName <> "" Then     '�u���[�N�v�t�H���_�Ƀt�@�C��������
        '�t�@�C�����i�[
        sFileName = StrConv(MidB(StrConv(uVerData(0).sFileName & Space(12), vbFromUnicode), 1, 16), vbUnicode)
    ElseIf uVerData(1).sFileName <> "" Then '�u���s�v�t�H���_�Ƀt�@�C��������
        '�t�@�C�����i�[
        sFileName = StrConv(MidB(StrConv(uVerData(1).sFileName & Space(12), vbFromUnicode), 1, 16), vbUnicode)
    Else                                    '�u���v�t�H���_�Ƀt�@�C��������
        '�t�@�C�����i�[
        sFileName = StrConv(MidB(StrConv(uVerData(2).sFileName & Space(12), vbFromUnicode), 1, 16), vbUnicode)
    End If
    sFileName = sFileName & " "

    If uVerData(0).sFileName <> "" Then     '�u���[�N�v�t�H���_�Ƀt�@�C��������
        '�o�[�W�������i�[
        'EG20 V30.1.0.1 DEL START
'        sFileInfo(0) = " " & StrConv(MidB(StrConv(uVerData(0).sMachineName & Space(10), vbFromUnicode), 1, 9), vbUnicode)
'        sFileInfo(0) = sFileInfo(0) & StrConv(MidB(StrConv(uVerData(0).sFooterFile & Space(8), vbFromUnicode), 1, 10), vbUnicode)
'        sFileInfo(0) = sFileInfo(0) & StrConv(MidB(StrConv(uVerData(0).sFileDate & Space(15), vbFromUnicode), 1, 14), vbUnicode)
'        sFileInfo(0) = sFileInfo(0) & uVerData(0).sVersion
        'EG20 V30.1.0.1 DEL END
        'EG20 V30.1.0.1 ADD START
        If Len(Trim(uVerData(0).sSyubetsu)) <> 0 Then
            sFileInfo(0) = " " & StrConv(MidB(StrConv(uVerData(0).sSyubetsu & Space(6), vbFromUnicode), 1, 4), vbUnicode)
        Else
            sFileInfo(0) = " " & Left(String(3, "-") & Space(6), 4)
        End If
        If Len(Trim(uVerData(0).sDataVersion)) <> 0 Then
            sFileInfo(0) = sFileInfo(0) & StrConv(MidB(StrConv(uVerData(0).sDataVersion & Space(20), vbFromUnicode), 1, 21), vbUnicode)
        Else
            sFileInfo(0) = sFileInfo(0) & Left(String(20, "-") & Space(20), 21)
        End If
        If Len(Trim(uVerData(0).sFileDate)) <> 0 Then
            sFileInfo(0) = sFileInfo(0) & StrConv(MidB(StrConv(uVerData(0).sFileDate, vbFromUnicode), 1, 16), vbUnicode)
        Else
            sFileInfo(0) = sFileInfo(0) & Left(String(16, "-") & Space(20), 16)
        End If
            
        'EG20 V30.1.0.1 ADD END
        sComment1(0) = " " & StrConv(MidB(StrConv(uVerData(0).sComment, vbFromUnicode), 1, 32), vbUnicode)
        sComment2(0) = " " & StrConv(MidB(StrConv(uVerData(0).sComment, vbFromUnicode), 33, 32), vbUnicode)
        If Len(Trim(sComment1(0))) <> 0 Then
            '���̂܂�
        Else
            sComment1(0) = " " & String(32, "-")
        End If
            
    End If
    If uVerData(1).sFileName <> "" Then     '�u���s�t�H���_�Ƀt�@�C��������
        '�o�[�W�������i�[
        'EG20 V30.1.0.1 DEL START
'        sFileInfo(1) = " " & StrConv(MidB(StrConv(uVerData(1).sMachineName & Space(10), vbFromUnicode), 1, 9), vbUnicode)
'        sFileInfo(1) = sFileInfo(1) & StrConv(MidB(StrConv(uVerData(1).sFooterFile & Space(8), vbFromUnicode), 1, 10), vbUnicode)
'        sFileInfo(1) = sFileInfo(1) & StrConv(MidB(StrConv(uVerData(1).sFileDate & Space(15), vbFromUnicode), 1, 14), vbUnicode)
'        sFileInfo(1) = sFileInfo(1) & uVerData(1).sVersion
        'EG20 V30.1.0.1 DEL END
        'EG20 V30.1.0.1 ADD START
        If Len(Trim(uVerData(1).sSyubetsu)) <> 0 Then
            sFileInfo(1) = " " & StrConv(MidB(StrConv(uVerData(1).sSyubetsu & Space(6), vbFromUnicode), 1, 4), vbUnicode)
        Else
            sFileInfo(1) = " " & Left(String(3, "-") & Space(6), 4)
        End If
        If Len(Trim(uVerData(1).sDataVersion)) <> 0 Then
            sFileInfo(1) = sFileInfo(1) & StrConv(MidB(StrConv(uVerData(1).sDataVersion & Space(20), vbFromUnicode), 1, 21), vbUnicode)
        Else
            sFileInfo(1) = sFileInfo(1) & Left(String(20, "-") & Space(20), 21)
        End If
        If Len(Trim(uVerData(1).sFileDate)) <> 0 Then
            sFileInfo(1) = sFileInfo(1) & StrConv(MidB(StrConv(uVerData(1).sFileDate, vbFromUnicode), 1, 16), vbUnicode)
        Else
            sFileInfo(1) = sFileInfo(1) & Left(String(16, "-") & Space(20), 16)
        End If
        'EG20 V30.1.0.1 ADD END
        sComment1(1) = " " & StrConv(MidB(StrConv(uVerData(1).sComment, vbFromUnicode), 1, 32), vbUnicode)
        sComment2(1) = " " & StrConv(MidB(StrConv(uVerData(1).sComment, vbFromUnicode), 33, 32), vbUnicode)
        If Len(Trim(sComment1(1))) <> 0 Then
            '���̂܂�
        Else
            sComment1(1) = " " & String(32, "-")
        End If
        
    End If
    If uVerData(2).sFileName <> "" Then     '�u���v�t�H���_�Ƀt�@�C��������
        '�o�[�W�������i�[
        'EG20 V30.1.0.1 DEL START
'        sFileInfo(2) = " " & StrConv(MidB(StrConv(uVerData(2).sMachineName & Space(10), vbFromUnicode), 1, 9), vbUnicode)
'        sFileInfo(2) = sFileInfo(2) & StrConv(MidB(StrConv(uVerData(2).sFooterFile & Space(8), vbFromUnicode), 1, 10), vbUnicode)
'        sFileInfo(2) = sFileInfo(2) & StrConv(MidB(StrConv(uVerData(2).sFileDate & Space(15), vbFromUnicode), 1, 14), vbUnicode)
'        sFileInfo(2) = sFileInfo(2) & uVerData(2).sVersion
        'EG20 V30.1.0.1 DEL END
        'EG20 V30.1.0.1 ADD START
        If Len(Trim(uVerData(2).sSyubetsu)) <> 0 Then
            sFileInfo(2) = " " & StrConv(MidB(StrConv(uVerData(2).sSyubetsu & Space(6), vbFromUnicode), 1, 4), vbUnicode)
        Else
            sFileInfo(2) = " " & Left(String(3, "-") & Space(6), 4)
        End If
        If Len(Trim(uVerData(2).sDataVersion)) <> 0 Then
            sFileInfo(2) = sFileInfo(2) & StrConv(MidB(StrConv(uVerData(2).sDataVersion & Space(20), vbFromUnicode), 1, 21), vbUnicode)
        Else
            sFileInfo(2) = sFileInfo(2) & Left(String(20, "-") & Space(20), 21)
        End If
        If Len(Trim(uVerData(2).sFileDate)) <> 0 Then
            sFileInfo(2) = sFileInfo(2) & StrConv(MidB(StrConv(uVerData(2).sFileDate, vbFromUnicode), 1, 16), vbUnicode)
        Else
            sFileInfo(2) = sFileInfo(2) & Left(String(16, "-") & Space(20), 16)
        End If
        'EG20 V30.1.0.1 ADD END
        sComment1(2) = " " & StrConv(MidB(StrConv(uVerData(2).sComment, vbFromUnicode), 1, 32), vbUnicode)
        sComment2(2) = " " & StrConv(MidB(StrConv(uVerData(2).sComment, vbFromUnicode), 33, 32), vbUnicode)
        If Len(Trim(sComment1(2))) <> 0 Then
            '���̂܂�
        Else
            sComment1(2) = " " & String(32, "-")
        End If
    End If


    If chkFolder(0).Value = CHECKBOX_ON Then               '����[�N��t�H���_�\��
        If uVerData(0).sFileName <> "" Then         '����[�N��t�H���_�Ƀt�@�C���͂���
            If chkFolder(1).Value = CHECKBOX_ON Then       '����s��t�H���_�\��
                If uVerData(1).sFileName <> "" Then '����s��t�H���_�Ƀt�@�C���͂���
                    '����[�N��t�H���_�Ƣ���s��t�H���_���r����
                    If sFileInfo(0) = sFileInfo(1) Then
                        If chkFolder(2).Value = CHECKBOX_ON Then   '�����t�H���_�\��
                            If uVerData(2).sFileName <> "" Then
                                '����s��t�H���_�Ƣ����t�H���_���r����
                                If sFileInfo(1) = sFileInfo(2) Then
'                                    lstKan(0).AddItem sFileName & "W N O" & sFileInfo(0)
                                    lstKan(iTab_index).AddItem sFileName & "W N O" & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment1(0)
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                    End If
                                    'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                    'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                    If sComment2(0) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment2(0)
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                    End If
                                Else
'                                    lstKan(0).AddItem sFileName & "W N  " & sFileInfo(0)
                                    lstKan(iTab_index).AddItem sFileName & "W N  " & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment1(0)
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                    End If
                                    'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                    'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                    If sComment2(0) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment2(0)
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                    End If
'                                    lstKan(0).AddItem Space(17) & "    O" & sFileInfo(2)
                                    lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                                    If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment1(2)
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                                    End If
                                    'If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                    'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                    If sComment2(2) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment2(2)
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                                    End If
                                End If
                            Else                            '�����t�H���_�Ƀt�@�C�����Ȃ�
'                                lstKan(0).AddItem sFileName & "W N  " & sFileInfo(1)
                                lstKan(iTab_index).AddItem sFileName & "W N  " & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                End If
                                'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                If sComment2(1) <> "" Then
'                                     lstKan(0).AddItem Space(22) & sComment2(1)
                                     lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                End If
'                                lstKan(0).AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                                'lstKan(iTab_index).AddItem Space(17) & "    O" & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                                'EG20 V30.1.0.1 ADD START
'                                lstKan(iTab_index).AddItem sFileName & "    O" & " ---  ----" & Space(16) & "----/--/-- --:--"
'                                lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                                'EG20 V30.1.0.1 ADD END
                            End If
                        Else                                '�����t�H���_��A�N�e�B�u�\��
'                            lstKan(0).AddItem sFileName & "W N  " & sFileInfo(0)
                            lstKan(iTab_index).AddItem sFileName & "W N  " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                            End If
                            'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                            'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                            If sComment2(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                            End If
                        End If
                    Else                            '����[�N��t�H���_�Ƣ���s��t�H���_�̃o�[�W�������Ⴄ
' EG20 V6.4.0.1�y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z�폜�J�n
''                        lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
'                        lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
'                        If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
''                            lstKan(0).AddItem Space(22) & sComment1(0)
'                            lstKan(iTab_index).AddItem Space(22) & sComment1(0)
'                        End If
'                        If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
''                            lstKan(0).AddItem Space(22) & sComment2(0)
'                            lstKan(iTab_index).AddItem Space(22) & sComment2(0)
'                        End If
' EG20 V6.4.0.1�y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z�폜�I��
                        If chkFolder(2).Value = CHECKBOX_ON Then   '�����t�H���_�\��
                            If uVerData(2).sFileName <> "" Then
                                '����s��t�H���_�Ƣ����t�H���_���r����
                                If sFileInfo(1) = sFileInfo(2) Then
' EG20 V6.4.0.1�y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z�ǉ��J�n
                                    lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                    End If
                                    'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                    'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                    If sComment2(0) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                    End If
' EG20 V6.4.0.1�y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z�ǉ��I��
'                                    lstKan(0).AddItem Space(17) & "  N O" & sFileInfo(1)
                                    lstKan(iTab_index).AddItem sFileName & "  N O" & sFileInfo(1)
                                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment1(1)
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                    End If
                                    'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                    'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                    If sComment2(1) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment2(1)
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                    End If
' EG20 V6.4.0.1�y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z�ǉ��J�n
                                ElseIf sFileInfo(0) = sFileInfo(2) Then
                                    ' �u���[�N�v���u���v�̏ꍇ
                                    lstKan(iTab_index).AddItem sFileName & "W   O" & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                    End If
                                    'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                    'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                    If sComment2(0) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                    End If
                                    lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                    End If
                                    'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                    'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                    If sComment2(1) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                    End If
' EG20 V6.4.0.1�y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z�ǉ��I��
                                Else
' EG20 V6.4.0.1�y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z�ǉ��J�n
                                    ' �u���[�N�v�� �u���s�v ���u���v�̏ꍇ
                                    lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                    End If
                                    'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                    'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                    If sComment2(0) <> "" Then
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                    End If
' EG20 V6.4.0.1�y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z�ǉ��I��
'                                    lstKan(0).AddItem Space(17) & "  N  " & sFileInfo(1)
                                    lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment1(1)
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                    End If
                                    'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                    'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                    If sComment2(1) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment2(1)
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                    End If
'                                    lstKan(0).AddItem Space(17) & "    O" & sFileInfo(2)
                                    lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                                    If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment1(2)
                                        lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                                    End If
                                    'If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                    'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                    If sComment2(2) <> "" Then
'                                        lstKan(0).AddItem Space(22) & sComment2(2)
                                        lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                                    End If
                                End If
                            Else                            '�����t�H���_�Ƀt�@�C�����Ȃ�
' EG20 V6.4.0.1�y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z�ǉ��J�n
                                ' �u���[�N�v�� �u���s�v ���u���v�̏ꍇ
                                lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                                If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                End If
                                'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                If sComment2(0) <> "" Then
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                End If
' EG20 V6.4.0.1�y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z�ǉ��I��
'                                lstKan(0).AddItem Space(17) & "  N  " & sFileInfo(1)
                                lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                End If
                                'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                If sComment2(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                End If
'                                lstKan(0).AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                                'lstKan(iTab_index).AddItem Space(17) & "    O" & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                                'EG20 V30.1.0.1 ADD START
'                                lstKan(iTab_index).AddItem sFileName & "    O" & " ---  ----" & Space(16) & "----/--/-- --:--"
'                                lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                                'EG20 V30.1.0.1 ADD END
                            End If
                        Else
' EG20 V6.4.0.1�y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z�ǉ��J�n
                            ' �u���[�N�v�� �u���s�v ���u���v�̏ꍇ
                            lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
                                lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                            End If
                            'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                            'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                            If sComment2(0) <> "" Then
                                lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                            End If
' EG20 V6.4.0.1�y���_���C���Ή��F���[�N�����s�A���[�N�����̏ꍇ�̕\���s���z�ǉ��I��
'                            lstKan(0).AddItem Space(17) & "  N  " & sFileInfo(1)
                            lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                            End If
                            'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                            'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                            If sComment2(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                            End If
                        End If
                    End If
                Else                                    '����s��t�H���_�Ƀt�@�C�����Ȃ�
                    If chkFolder(2).Value = CHECKBOX_ON Then   '�����t�H���_�\��
                        If uVerData(2).sFileName <> "" Then
                            If sFileInfo(0) = sFileInfo(2) Then
'                                lstKan(0).AddItem sFileName & "W   O" & sFileInfo(0)
                                lstKan(iTab_index).AddItem sFileName & "W   O" & sFileInfo(0)
                                If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(0)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                End If
                                'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                If sComment2(0) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(0)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                End If
'                                lstKan(0).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                                'lstKan(iTab_index).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                                'EG20 V30.1.0.1 ADD START
'                                lstKan(iTab_index).AddItem sFileName & "  N  " & " ---  ----" & Space(16) & "----/--/-- --:--"
'                                lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                                'EG20 V30.1.0.1 ADD END
                            Else
'                                lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
                                lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                                If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(0)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                                End If
                                'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                                'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                If sComment2(0) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(0)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                                End If
'                                lstKan(0).AddItem Space(17) & "    O" & sFileInfo(2)
                                lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                                If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(2)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                                End If
                                'If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                If sComment2(2) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(2)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                                End If
'                                lstKan(0).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                                'lstKan(iTab_index).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                                'EG20 V30.1.0.1 ADD START
'                                lstKan(iTab_index).AddItem sFileName & "  N  " & " ---  ----" & Space(16) & "----/--/-- --:--"
'                                lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                                'EG20 V30.1.0.1 ADD END
                            End If
                        Else                            '�����t�H���_�Ƀt�@�C�����Ȃ�
'                            lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
                            lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                            End If
                            'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                            'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                            If sComment2(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                            End If
'                            lstKan(0).AddItem Space(17) & "  N O" & " -------- --------  -------- ----"
                            'lstKan(iTab_index).AddItem Space(17) & "  N O" & " -------- --------  -------- ----"    'EG20 DEL
                            'EG20 V30.1.0.1 ADD START
'                            lstKan(iTab_index).AddItem sFileName & "  N O" & " ---  ----" & Space(16) & "----/--/-- --:--"
'                            lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                            'EG20 V30.1.0.1 ADD END
                        End If
                    Else                                '�����t�H���_��A�N�e�B�u�\��
'                        lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
                        lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                        If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment1(0)
                            lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                        End If
                        'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                        'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                        If sComment2(0) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment2(0)
                            lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                        End If
'                        lstKan(0).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                        'lstKan(iTab_index).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"    'EG20 DEL
                        'EG20 V30.1.0.1 ADD START
'                        lstKan(iTab_index).AddItem sFileName & "  N  " & " ---  ----" & Space(16) & "----/--/-- --:--"
'                        lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                        'EG20 V30.1.0.1 ADD END
                    End If
                End If
            Else                                        '����s��t�H���_��A�N�e�B�u�\��
                If chkFolder(2).Value = CHECKBOX_ON Then       '�����t�H���_�\��
                    If uVerData(2).sFileName <> "" Then
                        If sFileInfo(0) = sFileInfo(2) Then
'                            lstKan(0).AddItem sFileName & "W   O" & sFileInfo(0)
                            lstKan(iTab_index).AddItem sFileName & "W   O" & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                            End If
                            'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                            'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                            If sComment2(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                            End If
                        Else
'                            lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
                            lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                            If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                            End If
                            'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                            'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                            If sComment2(0) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(0)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                            End If
'                            lstKan(0).AddItem Space(17) & "    O" & sFileInfo(2)
                            lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                            If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(2)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                            End If
                            'If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                            'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                            If sComment2(2) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(2)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                            End If
                        End If
                    Else                                    '�����t�H���_�Ƀt�@�C�����Ȃ�
'                        lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
                        lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                        If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment1(0)
                            lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                        End If
                        'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                        'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                        If sComment2(0) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment2(0)
                            lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                        End If
'                        lstKan(0).AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                        'lstKan(iTab_index).AddItem Space(17) & "    O" & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                       'EG20 V30.1.0.1 ADD START
'                       lstKan(iTab_index).AddItem sFileName & "    O" & " ---  ----" & Space(16) & "----/--/-- --:--"
'                       lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                       'EG20 V30.1.0.1 ADD END
                    End If
                Else                                    '�����t�H���_��A�N�e�B�u�\��
'                    lstKan(0).AddItem sFileName & "W    " & sFileInfo(0)
                    lstKan(iTab_index).AddItem sFileName & "W    " & sFileInfo(0)
                    If Not IsNull(sComment1(0)) Or sComment1(0) <> "" Then
'                        lstKan(0).AddItem Space(22) & sComment1(0)
                        lstKan(iTab_index).AddItem Space(22) & sComment1(0)
                    End If
                    'If Not IsNull(sComment2(0)) Or sComment2(0) <> "" Then
                    'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                    If sComment2(0) <> "" Then
'                        lstKan(0).AddItem Space(22) & sComment2(0)
                        lstKan(iTab_index).AddItem Space(22) & sComment2(0)
                    End If
                End If
            End If
        Else                                '����[�N��t�H���_�Ƀt�@�C�����Ȃ�
            If chkFolder(1).Value = CHECKBOX_ON Then               '����s��t�H���_�\��
                If uVerData(1).sFileName <> "" Then         '����s��t�H���_�Ƀt�@�C���͂���
                    If chkFolder(2).Value = CHECKBOX_ON Then       '�����t�H���_�\��
                        If uVerData(2).sFileName <> "" Then '�����t�H���_�Ƀt�@�C���͂���
                            '����s��t�H���_�Ƣ����t�H���_���r����
                            If sFileInfo(1) = sFileInfo(2) Then
'                                lstKan(0).AddItem sFileName & "  N O" & sFileInfo(1)
                                lstKan(iTab_index).AddItem sFileName & "  N O" & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                End If
                                'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                If sComment2(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                End If
'                                lstKan(0).AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                                'lstKan(iTab_index).AddItem Space(17) & "W    " & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                                'EG20 V30.1.0.1 ADD START
'                                lstKan(iTab_index).AddItem sFileName & "W    " & " ---  ----" & Space(16) & "----/--/-- --:--"
'                                lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                                'EG20 V30.1.0.1 ADD END
                            Else                            '����s��t�H���_�Ƣ����t�H���_�̃o�[�W�������Ⴄ
'                                lstKan(0).AddItem sFileName & "  N  " & sFileInfo(1)
                                lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                                If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                                End If
                                'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                                'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                If sComment2(1) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(1)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                                End If
'                                lstKan(0).AddItem Space(17) & "    O" & sFileInfo(2)
                                lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                                If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment1(2)
                                    lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                                End If
                                'If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                                'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                                If sComment2(2) <> "" Then
'                                    lstKan(0).AddItem Space(22) & sComment2(2)
                                    lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                                End If
'                                lstKan(0).AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                                'lstKan(iTab_index).AddItem Space(17) & "W    " & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                                'EG20 V30.1.0.1 ADD START
'                                lstKan(iTab_index).AddItem sFileName & "W    " & " ---  ----" & Space(16) & "----/--/-- --:--"
'                                lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                                'EG20 V30.1.0.1 ADD END
                            End If
                        Else                                '�����t�H���_�Ƀt�@�C���͂Ȃ�
'                            lstKan(0).AddItem sFileName & "  N  " & sFileInfo(1)
                            lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                            End If
                            'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                            'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                            If sComment2(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                            End If
'                            lstKan(0).AddItem Space(17) & "W   O" & " -------- --------  -------- ----"
                            'lstKan(iTab_index).AddItem Space(17) & "W   O" & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                            'EG20 V30.1.0.1 ADD START
'                            lstKan(iTab_index).AddItem sFileName & "W   O" & " ---  ----" & Space(16) & "----/--/-- --:--"
'                            lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                            'EG20 V30.1.0.1 ADD END
                        End If
                    Else                                    '�����t�H���_��A�N�e�B�u�\��
'                        lstKan(0).AddItem sFileName & "  N  " & sFileInfo(1)
                        lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                        If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment1(1)
                            lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                        End If
                        'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                        'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                        If sComment2(1) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment2(1)
                            lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                        End If
'                        lstKan(0).AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                        'lstKan(iTab_index).AddItem Space(17) & "W    " & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
'                        lstKan(iTab_index).AddItem sFileName & "W    " & " ---  ----" & Space(16) & "----/--/-- --:--"
'                        lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                        'EG20 V30.1.0.1 ADD END
                    End If
                Else                                        '����s��t�H���_�Ƀt�@�C�����Ȃ�
                    If chkFolder(2).Value = CHECKBOX_ON Then       '�����t�H���_�\��
                        If uVerData(2).sFileName <> "" Then
'                            lstKan(0).AddItem sFileName & "    O" & sFileInfo(2)
                            lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                            If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(2)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                            End If
                            'If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                            'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                            If sComment2(2) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(2)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                            End If
'                            lstKan(0).AddItem Space(17) & "W N  " & " -------- --------  -------- ----"
                            'lstKan(iTab_index).AddItem Space(17) & "W N  " & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                            'EG20 V30.1.0.1 ADD START
'                            lstKan(iTab_index).AddItem sFileName & "W N  " & " ---  ----" & Space(16) & "----/--/-- --:--"
'                            lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                            'EG20 V30.1.0.1 ADD END
                        Else                                '�����t�H���_�Ƀt�@�C�����Ȃ�
'                            lstKan(0).AddItem sFileName & "W N O" & " -------- --------  -------- ----"
                            'lstKan(iTab_index).AddItem sFileName & "W N O" & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                            'EG20 V30.1.0.1 ADD START
'                            lstKan(iTab_index).AddItem sFileName & "W N O" & " ---  ----" & Space(16) & "----/--/-- --:--"
'                            lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                            'EG20 V30.1.0.1 ADD END
                        End If
                    Else                                    '�����t�H���_��A�N�e�B�u�\��
'                        lstKan(0).AddItem sFileName & "W N  " & " -------- --------  -------- ----"
                        'lstKan(iTab_index).AddItem sFileName & "W N  " & " -------- --------  -------- ----"    'EG20 V30.1.0.1DEL
                        'EG20 V30.1.0.1 ADD START
'                        lstKan(iTab_index).AddItem sFileName & "W N  " & " ---  ----" & Space(16) & "----/--/-- --:--"
'                        lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                        'EG20 V30.1.0.1 ADD END
                    End If
                End If
            Else                                        '����s��t�H���_��A�N�e�B�u�\��
                If chkFolder(2).Value = CHECKBOX_ON Then       '�����t�H���_�\��
                    If uVerData(2).sFileName <> "" Then
'                        lstKan(0).AddItem sFileName & "    O" & sFileInfo(2)
                        lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                        If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment1(2)
                            lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                        End If
                        'If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                        'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                        If sComment2(2) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment2(2)
                            lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                        End If
'                        lstKan(0).AddItem Space(17) & "W    " & " -------- --------  -------- ----"
                        'lstKan(iTab_index).AddItem Space(17) & "W    " & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
'                        lstKan(iTab_index).AddItem sFileName & "W    " & " ---  ----" & Space(16) & "----/--/-- --:--"
'                        lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                        'EG20 V30.1.0.1 ADD END
                    Else                                '�����t�H���_�Ƀt�@�C�����Ȃ�
'                        lstKan(0).AddItem sFileName & "W   O" & " -------- --------  -------- ----"
                        'lstKan(iTab_index).AddItem sFileName & "W   O" & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
'                        lstKan(iTab_index).AddItem sFileName & "W   O" & " ---  ----" & Space(16) & "----/--/-- --:--"
'                        lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                        'EG20 V30.1.0.1 ADD END
                    End If
                Else                                    '�����t�H���_��A�N�e�B�u�\��
'                    lstKan(0).AddItem sFileName & "W    " & " -------- --------  -------- ----"
                    'lstKan(iTab_index).AddItem sFileName & "W    " & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                    'EG20 V30.1.0.1 ADD START
'                    lstKan(iTab_index).AddItem sFileName & "W    " & " ---  ----" & Space(16) & "----/--/-- --:--"
'                    lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                    'EG20 V30.1.0.1 ADD END
                End If
            End If
        End If
    Else                                                '����[�N��t�H���_��A�N�e�B�u�\��
        If chkFolder(1).Value = CHECKBOX_ON Then               '����s��t�H���_�\��
            If uVerData(1).sFileName <> "" Then         '����s��t�H���_�Ƀt�@�C���͂���
                If chkFolder(2).Value = CHECKBOX_ON Then       '�����t�H���_�\��
                    If uVerData(2).sFileName <> "" Then '�����t�H���_�Ƀt�@�C���͂���
                        '����s��t�H���_�Ƣ����t�H���_���r����
                        If sFileInfo(1) = sFileInfo(2) Then
'                            lstKan(0).AddItem sFileName & "  N O" & sFileInfo(1)
                            lstKan(iTab_index).AddItem sFileName & "  N O" & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                            End If
                            'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                            'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                            If sComment2(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                            End If
                        Else
'                            lstKan(0).AddItem sFileName & "  N  " & sFileInfo(1)
                            lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                            If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                            End If
                            'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                            'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                            If sComment2(1) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(1)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                            End If
'                            lstKan(0).AddItem Space(17) & "    O" & sFileInfo(2)
                            lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                            If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment1(2)
                                lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                            End If
                            'If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                            'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                            If sComment2(2) <> "" Then
'                                lstKan(0).AddItem Space(22) & sComment2(2)
                                lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                            End If
                        End If
                    Else                                '�����t�H���_�Ƀt�@�C���͂Ȃ�
'                        lstKan(0).AddItem sFileName & "  N  " & sFileInfo(1)
                        lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                        If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment1(1)
                            lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                        End If
                        'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                        'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                        If sComment2(1) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment2(1)
                            lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                        End If
'                        lstKan(0).AddItem Space(17) & "    O" & " -------- --------  -------- ----"
                        'lstKan(iTab_index).AddItem Space(17) & "    O" & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
'                        lstKan(iTab_index).AddItem sFileName & "    O" & " ---  ----" & Space(16) & "----/--/-- --:--"
'                        lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                        'EG20 V30.1.0.1 ADD END
                    End If
                Else                                    '�����t�H���_��A�N�e�B�u�\��
'                    lstKan(0).AddItem sFileName & "  N  " & sFileInfo(1)
                    lstKan(iTab_index).AddItem sFileName & "  N  " & sFileInfo(1)
                    If Not IsNull(sComment1(1)) Or sComment1(1) <> "" Then
'                        lstKan(0).AddItem Space(22) & sComment1(1)
                        lstKan(iTab_index).AddItem Space(22) & sComment1(1)
                    End If
                    'If Not IsNull(sComment2(1)) Or sComment2(1) <> "" Then
                    'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                    If sComment2(1) <> "" Then
'                        lstKan(0).AddItem Space(22) & sComment2(1)
                        lstKan(iTab_index).AddItem Space(22) & sComment2(1)
                    End If
                End If
            Else                                        '����s��t�H���_�Ƀt�@�C�����Ȃ�
                If chkFolder(2).Value = CHECKBOX_ON Then       '�����t�H���_�\��
                    If uVerData(2).sFileName <> "" Then
'                        lstKan(0).AddItem sFileName & "    O" & sFileInfo(2)
                        lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                        If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment1(2)
                            lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                        End If
                        'If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                        'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                        If sComment2(2) <> "" Then
'                            lstKan(0).AddItem Space(22) & sComment2(2)
                            lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                        End If
'                        lstKan(0).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"
                        'lstKan(iTab_index).AddItem Space(17) & "  N  " & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
'                        lstKan(iTab_index).AddItem sFileName & "  N  " & " ---  ----" & Space(16) & "----/--/-- --:--"
'                        lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                        'EG20 V30.1.0.1 ADD END
                    Else                                '�����t�H���_�Ƀt�@�C�����Ȃ�
'                        lstKan(0).AddItem sFileName & "  N O" & " -------- --------  -------- ----"
                        'lstKan(iTab_index).AddItem sFileName & "  N O" & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
'                        lstKan(iTab_index).AddItem sFileName & "  N O" & " ---  ----" & Space(16) & "----/--/-- --:--"
'                        lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                        'EG20 V30.1.0.1 ADD END
                    End If
                Else                                    '�����t�H���_��A�N�e�B�u�\��
'                    lstKan(0).AddItem sFileName & "  N  " & " -------- --------  -------- ----"
                    'lstKan(iTab_index).AddItem sFileName & "  N  " & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                    'EG20 V30.1.0.1 ADD START
'                    lstKan(iTab_index).AddItem sFileName & "  N  " & " ---  ----" & Space(16) & "----/--/-- --:--"
'                    lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                    'EG20 V30.1.0.1 ADD END
                End If
            End If
        Else                                    '����s��t�H���_��A�N�e�B�u�\��
            If uVerData(2).sFileName <> "" Then '�����t�H���_�Ƀt�@�C���͂���
'                lstKan(0).AddItem sFileName & "    O" & sFileInfo(2)
                lstKan(iTab_index).AddItem sFileName & "    O" & sFileInfo(2)
                If Not IsNull(sComment1(2)) Or sComment1(2) <> "" Then
'                    lstKan(0).AddItem Space(22) & sComment1(2)
                    lstKan(iTab_index).AddItem Space(22) & sComment1(2)
                End If
                'If Not IsNull(sComment2(2)) Or sComment2(2) <> "" Then
                'IsNull��Null�𔻒f���邪Null�Ƃ����l�����邱�Ƃ͂Ȃ��BNot IsNull��or�����<>""�Ɣ��肪�ł��Ȃ��Ȃ�B
                If sComment2(2) <> "" Then
'                    lstKan(0).AddItem Space(22) & sComment2(2)
                    lstKan(iTab_index).AddItem Space(22) & sComment2(2)
                End If
            Else                                '�����t�H���_�Ƀt�@�C�����Ȃ�
'                lstKan(0).AddItem sFileName & "    O" & " -------- --------  -------- ----"
                'lstKan(iTab_index).AddItem sFileName & "    O" & " -------- --------  -------- ----"    'EG20 V30.1.0.1 DEL
                'EG20 V30.1.0.1 ADD START
'                lstKan(iTab_index).AddItem sFileName & "    O" & " ---  ----" & Space(16) & "----/--/-- --:--"
'                lstKan(iTab_index).AddItem Space(17) & Space(5) & " ----"
                'EG20 V30.1.0.1 ADD END
            End If
        End If
    End If
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
   On Error Resume Next
    
    '�ėp���[����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmKansenGateVerKanri.Caption, False
        pfFormActive (frmKansenGateVerKanri.hwnd)
    End If
End Sub

'EG20 V30.1.0.1 DEL START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSetFolderName
'//  �@�\����  : �f�[�^�W�J
'//  �@�\�T�v  : �t�H���_���Ȃǂ̃f�[�^���O���[�o���G���A�ɓW�J����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//                ���[�N�����s�R�s�[�ł̐������`�F�b�NINI�Ǎ���
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-03  CODED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y��ʃ`�F�b�N�@�\�ǉ��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Sub sSetFolderName()
'
'        TitleBox(0) = "����f�[�^"
'        TitleBox(1) = "�v���O����"
'        TitleBox(2) = "���CPU-Pro1"
'        TitleBox(3) = "���CPU-Pro2"
'        TitleBox(4) = "���CPU-Pro3"
'        TitleBox(5) = "�����i�n�r�j"
'        TitleBox(6) = "�\���P"
'        TitleBox(7) = "�\���Q"
'        TitleBox(8) = "�\���R"
'
'        LogBox(0) = "����"
'        LogBox(1) = "�v���O���C��"
'        LogBox(2) = "�T�u1"
'        LogBox(3) = "�T�u2"
'        LogBox(4) = "�T�u3"
'        LogBox(5) = "OS"
'        LogBox(5) = "�\��1"
'        LogBox(5) = "�\��2"
'        LogBox(5) = "�\��3"
'
'        '�t�H���_���ɐݒ���s��
'        FolderName(0, 0) = EG20_NHAN1WRK
'        FolderName(1, 0) = EG20_NHAN1NOW
'        FolderName(2, 0) = EG20_NHAN1OLD
'        FolderName(0, 1) = EG20_NPRO1WRK
'        FolderName(1, 1) = EG20_NPRO1NOW
'        FolderName(2, 1) = EG20_NPRO1OLD
'        FolderName(0, 2) = EG20_NSCP1WRK
'        FolderName(1, 2) = EG20_NSCP1NOW
'        FolderName(2, 2) = EG20_NSCP1OLD
'        FolderName(0, 3) = EG20_NSCP2WRK
'        FolderName(1, 3) = EG20_NSCP2NOW
'        FolderName(2, 3) = EG20_NSCP2OLD
'        FolderName(0, 4) = EG20_NSCP3WRK
'        FolderName(1, 4) = EG20_NSCP3NOW
'        FolderName(2, 4) = EG20_NSCP3OLD
'        FolderName(0, 5) = EG20_NOSWRK
'        FolderName(1, 5) = EG20_NOSNOW
'        FolderName(2, 5) = EG20_NOSOLD
'        FolderName(0, 6) = EG20_NYOBI1WRK
'        FolderName(1, 6) = EG20_NYOBI1NOW
'        FolderName(2, 6) = EG20_NYOBI1OLD
'        FolderName(0, 7) = EG20_NYOBI2WRK
'        FolderName(1, 7) = EG20_NYOBI2NOW
'        FolderName(2, 7) = EG20_NYOBI2OLD
'' EG20 V5.11.0.1�ǉ��J�n
'        FolderName(0, 8) = EG20_NYOBI3WRK
'        FolderName(1, 8) = EG20_NYOBI3NOW
'        FolderName(2, 8) = EG20_NYOBI3OLD
'' EG20 V5.11.0.1�ǉ��I��
'' EG20 V5.11.0.1�폜�J�n
''        FolderName(0, 8) = EG20_NYOBI2WRK
''        FolderName(1, 8) = EG20_NYOBI2NOW
''        FolderName(2, 8) = EG20_NYOBI2OLD
'' EG20 V5.11.0.1�폜�I��
'
'' EG20 V3.0.0.2�ǉ��J�n
'        DispTitleBox(0) = "����f�[�^  �o�[�W�����i���[�N�j�F"
'        DispTitleBox(1) = "�v���O����  �o�[�W�����i���[�N�j�F"
'        DispTitleBox(2) = "���CPU-Pro1 �o�[�W�����i���[�N�j�F"
'        DispTitleBox(3) = "���CPU-Pro2 �o�[�W�����i���[�N�j�F"
'        DispTitleBox(4) = "���CPU-Pro3 �o�[�W�����i���[�N�j�F"
'        DispTitleBox(5) = "�����i�n�r�j�o�[�W�����i���[�N�j�F"
'        DispTitleBox(6) = "�\���P      �o�[�W�����i���[�N�j�F"
'        DispTitleBox(7) = "�\���Q      �o�[�W�����i���[�N�j�F"
'        DispTitleBox(8) = "�\���R      �o�[�W�����i���[�N�j�F"
'' EG20 V3.0.0.2�ǉ��I��
'
'
''V1.20.0.1 ADD START
''-------EG-R����-------
'    ' �L�[��:����CPU-PRO��\
'    EHANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
'
'    ' �L�[��:���C��CPU-PRO��\
'    EMAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_MAIN_PRO, PATH_GATEVER_FILE)
'
'    ' �L�[���F�T�uCPU-PRO��\
'    ESUB_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_SUB_PRO, PATH_GATEVER_FILE)
'
'    ' �L�[��:���C��CPU-OS��\
'    EMAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_E, GATE_MAIN_OS, PATH_GATEVER_FILE)
'
'''-------NEG����-------
''    ' �L�[��:����CPU-PRO��\
''    NHANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
''
''    ' �L�[��:���C��CPU-PRO��\
''    NMAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_MAIN_PRO, PATH_GATEVER_FILE)
''
''    ' �L�[���F�T�uCPU-PRO��\
''    NSUB_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_SUB_PRO, PATH_GATEVER_FILE)
''
''    ' �L�[��:���C��CPU-OS��\
''    NMAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_N, GATE_MAIN_OS, PATH_GATEVER_FILE)
'''V1.20.0.1 ADD END
'
'' EG20 V5.11.0.1�y�^���\�����P�Ή��z�ǉ��J�n
'    gintUnkaiKind(0) = BootInfoGateType.TYPE_NHAN
'    gintUnkaiKind(1) = BootInfoGateType.TYPE_NPRO
'    gintUnkaiKind(2) = BootInfoGateType.TYPE_NSCP1
'    gintUnkaiKind(3) = BootInfoGateType.TYPE_NSCP2
'    gintUnkaiKind(4) = BootInfoGateType.TYPE_NSCP3
'    gintUnkaiKind(5) = BootInfoGateType.TYPE_NOS
'    gintUnkaiKind(6) = BootInfoGateType.TYPE_NYOBI1
'    gintUnkaiKind(7) = BootInfoGateType.TYPE_NYOBI2
'    gintUnkaiKind(8) = BootInfoGateType.TYPE_NYOBI3
'' EG20 V5.11.0.1�y�^���\�����P�Ή��z�ǉ��I��
'
'' EG20 V6.9.0.1�y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zADD START
'    gintProgramJudgeKind(0) = ProgramJudgeKind.JUDGE_NHAN       ' ����f�[�^
'    gintProgramJudgeKind(1) = ProgramJudgeKind.JUDGE_NPRO       ' �v���O����
'    gintProgramJudgeKind(2) = ProgramJudgeKind.JUDGE_NSCP1      ' �T�uCPU-Pro1
'    gintProgramJudgeKind(3) = ProgramJudgeKind.JUDGE_NSCP2      ' �T�uCPU-Pro2
'    gintProgramJudgeKind(4) = ProgramJudgeKind.JUDGE_NSCP3      ' �T�uCPU-Pro3
'    gintProgramJudgeKind(5) = ProgramJudgeKind.JUDGE_NOS        ' �����iOS�j
'    gintProgramJudgeKind(6) = ProgramJudgeKind.JUDGE_NOCHECK    ' �\��1
'    gintProgramJudgeKind(7) = ProgramJudgeKind.JUDGE_NOCHECK    ' �\��2
'    gintProgramJudgeKind(8) = ProgramJudgeKind.JUDGE_NOCHECK    ' �\��3
'' EG20 V6.9.0.1�y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zADD END

'End Sub
'EG20 V30.1.0.1 DEL END
'EG20 V30.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  �֐�����  : sSetFolderName
'//  �@�\����  : �f�[�^�W�J
'//  �@�\�T�v  : �t�H���_���Ȃǂ̃f�[�^���O���[�o���G���A�ɓW�J����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-02-18  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sSetFolderName()

        TitleBox(0) = "����b�o�t"
        TitleBox(1) = "���C���b�o�t"
        TitleBox(2) = "�T�u�b�o�t"
        TitleBox(3) = "�n�r"
        TitleBox(4) = "�\���P"
        TitleBox(5) = "�\���Q"
    
        LogBox(0) = "����"
        LogBox(1) = "�v���O���C��"
        LogBox(2) = "�T�u"
        LogBox(3) = "�n�r"
        LogBox(4) = "�\���P"
        LogBox(5) = "�\���Q"
        
        '�t�H���_���ɐݒ���s��
        FolderName(0, 0) = EG30_JHANWRK
        FolderName(1, 0) = EG30_JHANNOW
        FolderName(2, 0) = EG30_JHANOLD
        FolderName(0, 1) = EG30_JPROWRK
        FolderName(1, 1) = EG30_JPRONOW
        FolderName(2, 1) = EG30_JPROOLD
        FolderName(0, 2) = EG30_JSCPUWRK
        FolderName(1, 2) = EG30_JSCPUNOW
        FolderName(2, 2) = EG30_JSCPUOLD
        FolderName(0, 3) = EG30_JOSWRK
        FolderName(1, 3) = EG30_JOSNOW
        FolderName(2, 3) = EG30_JOSOLD
        FolderName(0, 4) = EG30_JYOBIWK1
        FolderName(1, 4) = EG30_JYOBINW1
        FolderName(2, 4) = EG30_JYOBIOD1
        FolderName(0, 5) = EG30_JYOBIWRK
        FolderName(1, 5) = EG30_JYOBINOW
        FolderName(2, 5) = EG30_JYOBIOLD

        DispTitleBox(0) = "����b�o�t  �o�[�W�����i���[�N�j�F"
        DispTitleBox(1) = "���C���b�o�t  �o�[�W�����i���[�N�j�F"
        DispTitleBox(2) = "�T�u�b�o�t �o�[�W�����i���[�N�j�F"
        DispTitleBox(3) = "�n�r �o�[�W�����i���[�N�j�F"
        DispTitleBox(4) = "�\���P �o�[�W�����i���[�N�j�F"
        DispTitleBox(5) = "�\���Q �o�[�W�����i���[�N�j�F"

'-------�V��������-------
    ' �L�[��:����CPU-PRO��\
    EG30_HANTEI_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG30, GATE_HANTEI_PRO, PATH_GATEVER_FILE)
    
    ' �L�[��:���C��CPU-PRO��\
    EG30_MAIN_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG30, GATE_MAIN_PRO, PATH_GATEVER_FILE)
        
    ' �L�[���F�T�uCPU-PRO��\
    EG30_SUB_CPU_CHK_FILE = sSetChkFile(GATE_TYPE_EG30, GATE_SUB_PRO1, PATH_GATEVER_FILE)
    
    ' �L�[��:���C��CPU-OS��\
    EG30_MAIN_OS_CHK_FILE = sSetChkFile(GATE_TYPE_EG30, GATE_MAIN_OS, PATH_GATEVER_FILE)

    gintUnkaiKind(0) = BootInfoGateType.TYPE_JHAN
    gintUnkaiKind(1) = BootInfoGateType.TYPE_JPRO
    gintUnkaiKind(2) = BootInfoGateType.TYPE_JSCPU
    gintUnkaiKind(3) = BootInfoGateType.TYPE_JOS

    gintProgramJudgeKind(0) = ProgramJudgeKind.JUDGE_JHAN       'a:����CPU�p�v���O�����i�����j
    gintProgramJudgeKind(1) = ProgramJudgeKind.JUDGE_JPRO       'b:���C��CPU�p�v���O�����i�����j
    gintProgramJudgeKind(2) = ProgramJudgeKind.JUDGE_JSCPU     'c:�T�uCPU�v���O�����i�����j
    gintProgramJudgeKind(3) = ProgramJudgeKind.JUDGE_JOS        ' d:OS�v���O�����i�����j
    gintProgramJudgeKind(4) = ProgramJudgeKind.JUDGE_YOBI1      'e:�\���P�i�����j �`�F�b�N����
    gintProgramJudgeKind(5) = ProgramJudgeKind.JUDGE_YOBI       'f:�\���i�����j �`�F�b�N����
    gintProgramJudgeKind(6) = ProgramJudgeKind.JUDGE_NOCHECK
    gintProgramJudgeKind(7) = ProgramJudgeKind.JUDGE_NOCHECK
    gintProgramJudgeKind(8) = ProgramJudgeKind.JUDGE_NOCHECK

End Sub
'EG20 V30.1.0.1 ADD END
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fMakeListbox
'//  �@�\����  : �o�[�W������񃊃X�g�{�b�N�X�쐬
'//  �@�\�T�v  : �e�t�H���_����o�[�W�����擾���s���A���X�g�{�b�N�X�쐬
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fMakeListbox() As Boolean
    
    Dim bRet As Boolean                        '�߂�l
    
    Dim sCorner As String                      '�R�[�i�[�ԍ�
    Dim sGatePath As String                    '�R�[�i�[�ԍ��t�t�@�C���p�X
    Dim sFilePath As String                    '�t�@�C���t�@�C���p�X
    Dim i As Integer                           '���[�v�J�E���^
    Dim sWorkVer As String                      ' ���[�N�o�[�W����
    Dim sNowVer As String                       ' ���s�o�[�W����
    Dim sOldVer As String                       ' ���o�[�W����

    On Error Resume Next

    sWorkVer = TITLEDISP_VERNOTHING
    sNowVer = TITLEDISP_VERNOTHING
    sOldVer = TITLEDISP_VERNOTHING
    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab

    sCorner = Format(iTab_index + 1, "00")

    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sGatePath = PATH_N_GATE & sCorner

    '***********************************************
    '* �����㎩���t�H���_����S�Ẵo�[�W���������擾���� *
    '***********************************************

    ReDim uVersion(0)

    '����[�N��t�H���_����t�@�C�����X�g���擾����
    sFilePath = sGatePath & FolderName(0, FolderSyubetu)
    bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
    If bRet = True Then
        '�t�@�C�����X�g����o�[�W���������擾����
        sWorkVer = sVersionInfo(sFilePath, MN_FLDWRK)
    End If

    '����s��t�H���_����t�@�C�����X�g���擾����
    sFilePath = sGatePath & FolderName(1, FolderSyubetu)
    bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
    If bRet = True Then
        '�t�@�C�����X�g����o�[�W���������擾����
        sNowVer = sVersionInfo(sFilePath, MN_FLDNOW)
    End If

    '�����t�H���_����t�@�C�����X�g���擾����
    sFilePath = sGatePath & FolderName(2, FolderSyubetu)
    bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
    If bRet = True Then
        '�t�@�C�����X�g����o�[�W���������擾����
        sOldVer = sVersionInfo(sFilePath, MN_FLDOLD)
    End If

    '�o�[�W���������t�@�C�������Ƀ\�[�g����
    sListboxSort

    '�o�[�W�����������X�g�{�b�N�X�ɃZ�b�g����
    Call sVerListDisp(sWorkVer, sNowVer, sOldVer)

End Function

' EG20 V3.0.0.2 �폜�J�n
'Private Function fMakeListbox() As Boolean
'
'    Dim bRet As Boolean                        '�߂�l
'
'    Dim sCorner As String                      '�R�[�i�[�ԍ�
'    Dim sGatePath As String                    '�R�[�i�[�ԍ��t�t�@�C���p�X
'    Dim sFilePath As String                    '�t�@�C���t�@�C���p�X
'    Dim i As Integer                           '���[�v�J�E���^
'
'    On Error Resume Next
'
''    ' �I�𒆂̃R�[�i�[�ԍ��擾
''    iTab_index = SSTab1.Tab
''
''    sCorner = Format(iTab_index + 1, "00")
''
''    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
''    sGatePath = PATH_N_GATE & sCorner
'
'    '***********************************************
'    '* �����㎩���t�H���_����S�Ẵo�[�W���������擾���� *
'    '***********************************************
'    For i = 0 To 5
'
'        iTab_index = i
'
'        sCorner = Format(iTab_index + 1, "00")
'
'        ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
'        sGatePath = PATH_N_GATE & sCorner
'
'        ReDim uVersion(0)
'
'        '����[�N��t�H���_����t�@�C�����X�g���擾����
'        sFilePath = sGatePath & FolderName(0, FolderSyubetu)
'
''       bRet = fReadFileList(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)
'        bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
'        If bRet = True Then
'            '�t�@�C�����X�g����o�[�W���������擾����
''            sVersionInfo FolderName(0, FolderSyubetu), MN_FLDWRK
'            sVersionInfo sFilePath, MN_FLDWRK
'        End If
'
'        '����s��t�H���_����t�@�C�����X�g���擾����
'        sFilePath = sGatePath & FolderName(1, FolderSyubetu)
'
''       bRet = fReadFileList(FolderName(1, FolderSyubetu) & "\" & MN_FILELIST)
'        bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
'        If bRet = True Then
'            '�t�@�C�����X�g����o�[�W���������擾����
''           sVersionInfo FolderName(1, FolderSyubetu), MN_FLDNOW
'            sVersionInfo sFilePath, MN_FLDNOW
'        End If
'
'        '�����t�H���_����t�@�C�����X�g���擾����
'        sFilePath = sGatePath & FolderName(2, FolderSyubetu)
'
''       bRet = fReadFileList(FolderName(2, FolderSyubetu) & "\" & MN_FILELIST)
'        bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
'        If bRet = True Then
'            '�t�@�C�����X�g����o�[�W���������擾����
''           sVersionInfo FolderName(2, FolderSyubetu), MN_FLDOLD
'            sVersionInfo sFilePath, MN_FLDOLD
'        End If
'
'        '�o�[�W���������t�@�C�������Ƀ\�[�g����
'        sListboxSort
'
'        '�o�[�W�����������X�g�{�b�N�X�ɃZ�b�g����
'        sVerListDisp
'
'    Next i
'End Function
' EG20 V3.0.0.2 �폜�I��
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sVerListDisp
'//  �@�\����  : �o�[�W������񃊃X�g�{�b�N�X�ݒ�
'//  �@�\�T�v  : �擾�����o�[�W���������A���X�g�{�b�N�X�ɐݒ�
'//
'//              �^        ����             �Ӗ�
'//  ����      : String    szWorkVersion    ���[�N�o�[�W����
'//  ����      : String    szNowVersion     ���s�o�[�W����
'//  ����      : String    szOldVersion     ���o�[�W����
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'/  REVISIONS    : (EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'/                  EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-18 REVISED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Sub sVerListDisp()                 ' EG20 V3.0.0.2�폜
' EG20 V3.0.0.2�ǉ��J�n
Private Sub sVerListDisp(szWorkVersion As String, _
                            szNowVersion As String, _
                            szOldVersion As String)
' EG20 V3.0.0.2�ǉ��I��

    Dim i As Integer                        '�J�E���^
    'Dim uVerData(2) As MN_VERSION_JIKAI     '�o�[�W�������i�e�t�H���_�j   'EG20 V30.1.0.1 DEL
    Dim uVerData(2) As MN_VERSION_KAN_JIKAI  '�o�[�W�������i�e�t�H���_�j   'EG20 V30.1.0.1 ADD
    Dim lDataNum As Long                    '�o�[�W�������
    Dim szWorkBuffer As String              ' ���[�N�o�b�t�@        ' EG20 V3.0.0.2�ǉ�
    Dim szTitleBuffer As String             ' ���[�N�o�b�t�@        ' EG20 V3.0.0.2�ǉ�

    On Error Resume Next

'    '���X�g�{�b�N�X������������
'    lstKan(0).Clear
'    lstKan(1).Clear
'    lstKan(2).Clear
'    lstKan(3).Clear
'    lstKan(4).Clear
'    lstKan(5).Clear

    lDataNum = UBound(uVersion)             '�o�[�W������񐔎擾
    For i = 1 To lDataNum

        uVerData(0).sFileName = ""          '�t�@�C�������N���A����
        uVerData(1).sFileName = ""          '�t�@�C�������N���A����
        uVerData(2).sFileName = ""          '�t�@�C�������N���A����

        Select Case uVersion(i).iFolder     '�t�H���_����ΏۂƂ���
        Case MN_FLDWRK                      '�u���[�N�v�t�H���_�̏ꍇ
            uVerData(0) = uVersion(i)       '�u���[�N�v�t�H���_���Ɋi�[����
            If i + 1 <= lDataNum Then       '���̃f�[�^������?
                If uVersion(i).sFileName = uVersion(i + 1).sFileName Then
                                                        '�t�@�C����������?
                    Select Case uVersion(i + 1).iFolder '�t�H���_����ΏۂƂ���
                    Case MN_FLDNOW                      '�u���s�v�t�H���_�̏ꍇ
                        uVerData(1) = uVersion(i + 1)   '�u���s�v�t�H���_���Ɋi�[����
                        If i + 2 <= lDataNum Then       '���̃f�[�^������?
                            If uVersion(i + 1).sFileName = uVersion(i + 2).sFileName Then
                                                        '�t�@�C����������?
                                uVerData(2) = uVersion(i + 2)
                                                        '�u���v�t�H���_���Ɋi�[����
                                i = i + 2               '�J�E���^�����X�ɂ���
                            Else
                                i = i + 1               '�J�E���^�����ɂ���
                            End If
                        Else
                            i = i + 1                   '�J�E���^�����ɂ���
                        End If
                    Case MN_FLDOLD                      '�u���v�t�H���_�̏ꍇ
                        uVerData(2) = uVersion(i + 1)   '�u���v�t�H���_���Ɋi�[����
                        i = i + 1                       '�J�E���^�����ɂ���
                    End Select
                End If
            End If
        Case MN_FLDNOW                      '�u���s�v�t�H���_�̏ꍇ
            uVerData(1) = uVersion(i)       '�u���s�v�t�H���_���Ɋi�[����
            If i + 1 <= lDataNum Then       '���̃f�[�^������
                If uVersion(i).sFileName = uVersion(i + 1).sFileName Then
                                                    '�t�@�C����������?
                    uVerData(2) = uVersion(i + 1)   '�u���v�t�H���_���Ɋi�[����
                    i = i + 1                       '�J�E���^�����ɂ���
                End If
            End If
        Case MN_FLDOLD                      '�u���v�t�H���_�̏ꍇ
            uVerData(2) = uVersion(i)       '�u���v�t�H���_���Ɋi�[����
        End Select
        '�t�@�C�������܂Ƃ߂ă��X�g�{�b�N�X�ɐݒ�
        sVersionDisp uVerData()
    Next

' EG20 V3.0.0.2�ǉ��J�n
    ' ���[�N�s�ҏW
    szWorkBuffer = DispTitleBox(FolderSyubetu) & szWorkVersion & vbCrLf
    szTitleBuffer = szWorkBuffer
    ' ���s�s�ҏW
    szWorkBuffer = TITLEDISP_FIXEDVERNOW & szNowVersion & vbCrLf
    szTitleBuffer = szTitleBuffer & szWorkBuffer
    ' ���s�ҏW
    szWorkBuffer = TITLEDISP_FIXEDVEROLD & szOldVersion
    szTitleBuffer = szTitleBuffer & szWorkBuffer

    lblZenVer(iTab_index).Caption = szTitleBuffer

    DispTitleVersion(MN_FOLD_WRK) = szWorkVersion
    DispTitleVersion(MN_FOLD_NOW) = szNowVersion
    DispTitleVersion(MN_FOLD_OLD) = szOldVersion
' EG20 V3.0.0.2�ǉ��I��

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSetChkFile
'//  �@�\����  : ���[�N�����s�R�s�[�Ŏg�p���鐳�����`�F�b�NINI�Ǎ���
'//  �@�\�T�v  : INI�t�@�C���ɂ̓��e���G���A�ɓW�J����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String    �Z�N�V������
'//              String    �L�[��
'//              String    �t�@�C����
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : String    �������`�F�b�NINI�̓��e�i�ُ펞�̓u�����N�j
'//
'//     ORIGINAL  :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yamazaki
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sSetChkFile(sSec As String, sKey As String, sFilePath As String) As String

    Dim iRet As Integer             '�֐��̖߂�l
    Dim sIni_Data As String * 128   'INI�t�@�C�����1�s���擾
    
    Dim sErrFile As String             '�G���[���O�pINI�t�@�C����
    Dim sErrExe As String              '�G���[���O�pINI�g���q
    Dim lngErrCode As Long             '�G���[�R�[�h

    
    '�G���[���[�`����錾
    On Error Resume Next

    'ini�t�@�C���擾
    sIni_Data = ""
    iRet = GetPrivateProfileString(sSec, sKey, DEFAILT, sIni_Data, Len(sIni_Data), sFilePath)
    
    '�ُ폈��
    If iRet = 0 Then
        
        '���O�o�́uINI�t�@�C���Ǎ��ُ�v
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_FREAD
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_ERR_INI_READ, lngErrCode)
        '���O�o�́@���t�@�C����
        Call psFileNameGet(sFilePath, sErrFile, sErrExe)             '�t�@�C���p�X����t�@�C�������擾
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "��File:" & sErrFile & "." & sErrExe, lngErrCode)
        '���O�o�́@���L�[��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, "��Key:" & sKey, lngErrCode)
        
    End If
    
    sSetChkFile = Left$(sIni_Data, iRet)
    
End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fReadFileList
'//  �@�\����  : �t�@�C�����X�g�̎擾
'//  �@�\�T�v  : �t�@�C�����X�g���A�t�@�C�������擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fReadFileList(sFileList As String) As Boolean
    Dim iFileNumber As Integer      '�t�@�C���ԍ�
    Dim sFileName As String         '�t�@�C����
    Dim iListCnt As Integer         '�t�@�C���i�[��

    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�

    iListCnt = 0
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)

    iFileNumber = FreeFile   '���g�p�̃t�@�C���ԍ����擾����

    Open sFileList For Input Access Read As #iFileNumber    '�t�@�C�����X�g�̃I�[�v��
    Do While Not EOF(iFileNumber)                           '�t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
        Line Input #iFileNumber, sFileName                  '�f�[�^��ǂݍ��݂܂��B
        If sFileName <> "" And Left$(sFileName, 1) <> "/" Then                '�t�@�C���������݂���
            iListCnt = iListCnt + 1                         '�t�@�C�����̃J�E���^���A�b�v����
            ReDim Preserve FileList(iListCnt)               '�t�@�C�����i�[�G���A���g������
            ReDim Preserve FileListType(iListCnt)           '�t�@�C�����i�[�G���A���g������
            'EG20 V30.1.0.1 DEL START
'            FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
'            FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 16)
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.1.0.1 ADD START
            '�t�@�C����ʂ͑啶���ɕϊ������A�t�@�C����������啶���ɕϊ�����悤�ɂ���B�i���܂ł͎�ʂ�����������������Ȃ������j
            FileListType(iListCnt - 1) = Trim$(Left$(sFileName, 18))
            FileList(iListCnt - 1) = UCase(Mid$(FileListType(iListCnt - 1), 3, 16))
            'EG20 V30.1.0.1 ADD�@END
                                                            '�t�@�C�������t�@�C�����i�[�G���A�ɃZ�b�g
        End If
    Loop
    Close #iFileNumber      '�t�@�C������܂��B

    fReadFileList = True    '�߂�l�𐳏�Ƃ���

    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    'V1.21.0.1 ADD  START
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    'V1.21.0.1 ADD  END
    fReadFileList = False   '�߂�l���G���[�Ƃ���
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sVersionInfo
'//  �@�\����  : �o�[�W�������̎擾
'//  �@�\�T�v  : �t�@�C�����X�g�ꗗ����o�[�W���������擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sPath
'//  �@�@�@    : Integer�@ iFolder
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-18  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Sub sVersionInfo(sPath As String, iFolder As Integer)                  ' EG20 V3.0.0.2�폜
Private Function sVersionInfo(sPath As String, iFolder As Integer) As String    ' EG20 V3.0.0.2�ǉ�
    Dim i As Integer                    '�J�E���^
    Dim j As Integer                    '�J�E���^
    Dim sMyName As String               '�t�@�C����
    Dim iFileNumber As Integer          '�t�@�C���ԍ�
    Dim lLen As Long                    '�t�@�C���T�C�Y
    'Dim uFooter As MN_FOOT              '�t�b�^���i�[�G���A      'EG20 V30.1.0.1 DEL
    Dim uFooter As MN_KAN_FOOT          '�t�b�^���i�[�G���A       'EG20 V30.1.0.1 ADD
    Dim uFooterDummy    As MN_KAN_FOOT  '�������p
    
    Dim lPos As Long                    '�o�[�W�������i�[�ʒu
    Dim sDateTime As String
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD

    Dim szResultVersion As String        ' �o�̓o�[�W����               ' EG20 V3.0.0.2�ǉ�

    szResultVersion = TITLEDISP_VERNOTHING                              ' EG20 V3.0.0.2�ǉ�
   On Error Resume Next

    For i = 0 To UBound(FileList) - 1   '�t�@�C�����X�g��

        sMyName = sPath & "\" & FileList(i)     '�t�@�C���t���p�X���̍쐬

        'If Dir(sMyName) <> "" Then              '�t�@�C�������݂���?    'V1.20.0.1 DEL
        If objFso.FileExists(sMyName) = True Then  '�t�@�C�������݂���?    'V1.20.0.1 ADD
            lLen = FileLen(sMyName)             '�t�@�C���T�C�Y�̎擾

            iFileNumber = FreeFile              '���g�p�̃t�@�C���ԍ����擾����

            Open sMyName For Binary Access Read As #iFileNumber
                                                '�t�@�C���̃I�[�v��
            uFooter = uFooterDummy  '�O��̕\���p�f�[�^���c���Ă���ꍇ������̂ŏ�����
            Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter
                                                '�t�b�^���̎擾
            ReDim Preserve uVersion(UBound(uVersion) + 1)
                                                '�o�[�W�������i�[�G���A�̊g��
            lPos = UBound(uVersion)             '�o�[�W�������i�[�ʒu�̎擾
            'uVersion(lPos).sFileName = UCase(FileListType(i))       '�t�@�C������啶���ɂ��ăZ�b�g    'EG20 V30.1.0.1 DEL
            uVersion(lPos).sFileName = UCase(FileList(i))       '�t�@�C������啶���ɂ��ăZ�b�g    'EG20 V30.1.0.1 ADD
            uVersion(lPos).iFolder = iFolder                    '�t�H���_���Z�b�g
            'uVersion(lPos).sMachineName = uFooter.sKisyu        '�@�햼�Z�b�g   'EG20 V30.1.0.1 DEL
            uVersion(lPos).sSyubetsu = LCase(Right$("0" & Hex(uFooter.bySyubetsu), 2)) & Chr(uFooter.byMakerName)  '��ʂ��Z�b�g   'EG20 V30.1.0.1 ADD
            'uVersion(lPos).sFooterFile = uFooter.sFileName      '�t�@�C�����Z�b�g      'EG20 V30.1.0.1 DEL
            'uVersion(lPos).sDataVersion = uFooter.sFileVersion     '�f�[�^���{�o�[�W����
            'JTR�\�[�X���Q�l�Ƀf�[�^���{�o�[�W������ҏW
            'NULL�����`�F�b�N
            For j = 0 To UBound(uFooter.byFileVersion)
                '�����ANULL�����i0x00)�������Ă���ꍇ�́A�X�y�[�X(0x20)�ɕύX
                If uFooter.byFileVersion(j) = &H0 Then
                    uFooter.byFileVersion(j) = &H20
                End If
            Next j
            '�f�[�^���{�o�[�W�������Z�b�g
            uVersion(lPos).sDataVersion = ""    '������
            For j = 0 To UBound(uFooter.byFileVersion)
                'ASCII�R�[�h���當����ɕϊ����Đݒ�
                uVersion(lPos).sDataVersion = uVersion(lPos).sDataVersion & Chr(uFooter.byFileVersion(j))
            Next j
            'EG20 V30.1.0.1 ADD START
            '�u�f�[�^���{�o�[�W�����v�i18�o�C�g�j�̌�Ƀt�@�C���ʃo�[�W������ǉ�
            For j = 0 To UBound(uFooter.byZentaiVersion)
                'ASCII�R�[�h���當����ɕϊ����Đݒ�
                uVersion(lPos).sDataVersion = uVersion(lPos).sDataVersion & Chr(uFooter.byZentaiVersion(j))
            Next j
            'EG20 V30.1.0.1 ADD END
            
            sDateTime = ""
            For j = 0 To 3
                sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
                'EG20 V30.1.0.1 ADD START
                If j = 1 Or j = 2 Then
                    sDateTime = sDateTime & "/"
                End If
                'EG20 V30.1.0.1 ADD END
            Next
            sDateTime = sDateTime & " "
            For j = 4 To 5
                sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
                'EG20 V30.1.0.1 ADD START
                If j = 4 Then
                    sDateTime = sDateTime & ":"
                End If
               'EG20 V30.1.0.1 ADD END
            Next
            uVersion(lPos).sFileDate = sDateTime
            'uVersion(lPos).sVersion = uFooter.sVersion          '�o�[�W�������Z�b�g   'EG20 V30.1.0.1 DEL
            'EG20 V30.1.0.1 ADD START
            '�o�[�W���������Z�b�g
            uVersion(lPos).sVersion = ""    '������
            For j = 0 To UBound(uFooter.byZentaiVersion)
                'ASCII�R�[�h���當����ɕϊ����Đݒ�
                uVersion(lPos).sVersion = uVersion(lPos).sVersion & Chr(uFooter.byZentaiVersion(j))
            Next
            'EG20 V30.1.0.1 ADD END
            uVersion(lPos).sComment = uFooter.sHyoji            '�\��������Z�b�g

' EG20 V3.0.0.2�ǉ��J�n
            ' �t�@�C�����X�g�̐擪�ŁA���ŏ��Ɍ��������t�@�C���̃o�[�W������ݒ�
            If szResultVersion = TITLEDISP_VERNOTHING Then
                'szResultVersion = uFooter.sVersion     'EG20 V30.1.0.1 DEL
                'EG20 V30.1.0.1 ADD START
                szResultVersion = ""    '������
                For j = 0 To UBound(uFooter.byFileVersion)
                    'ASCII�R�[�h���當����ɕϊ����Đݒ�
                    szResultVersion = szResultVersion & Chr(uFooter.byZentaiVersion(j))
                Next j
                'EG20 V30.1.0.1 ADD END
            End If
' EG20 V3.0.0.2�ǉ��I��

            Close #iFileNumber                  '�t�@�C������܂�
        End If
    Next
    
    Set objFso = Nothing    'V1.20.0.1 ADD

    sVersionInfo = szResultVersion              ' EG20 V3.0.0.2�ǉ�

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sListboxSort
'//  �@�\����  : �o�[�W�������̃\�[�g
'//  �@�\�T�v  : �o�[�W���������t�@�C�������Ƀ\�[�g����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-18 REVISED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sListboxSort()
    Dim i As Integer                '�J�E���^
    Dim j As Integer                '�J�E���^
    'Dim uBuff As MN_VERSION_JIKAI   '�o�[�W�������i�[�o�b�t�@    'EG20 V30.1.0.1 DEL
    Dim uBuff As MN_VERSION_KAN_JIKAI   '�o�[�W�������i�[�o�b�t�@     'EG20 V30.1.0.1 ADD

    On Error Resume Next
   
    For i = 1 To UBound(uVersion) - 1
        For j = i + 1 To UBound(uVersion)
            '�t�@�C�����̔�r���s��
            If uVersion(j).sFileName < uVersion(i).sFileName Then
                '�t�@�C��������������Έڂ��ւ���
                uBuff = uVersion(i)
                uVersion(i) = uVersion(j)
                uVersion(j) = uBuff
            ElseIf uVersion(j).sFileName = uVersion(i).sFileName Then
                '�t�H���_�̔�r���s��
                If uVersion(j).iFolder = MN_FLDWRK And uVersion(i).iFolder = MN_FLDNOW Then
                    uBuff = uVersion(i)
                    uVersion(i) = uVersion(j)
                    uVersion(j) = uBuff
                ElseIf uVersion(j).iFolder = MN_FLDNOW And uVersion(i).iFolder = MN_FLDOLD Then
                    uBuff = uVersion(i)
                    uVersion(i) = uVersion(j)
                    uVersion(j) = uBuff
                ElseIf uVersion(j).iFolder = MN_FLDWRK And uVersion(i).iFolder = MN_FLDOLD Then
                    uBuff = uVersion(i)
                    uVersion(i) = uVersion(j)
                    uVersion(j) = uBuff
                End If
            End If
        Next
    Next
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Public Sub psVersionDisp()
'
'    Dim strFilePath     As String   '�o�[�W�����t�@�C���p�X
'    Dim bRet            As Boolean  '�߂�l
'    Dim intFileNo       As Integer  '�t�@�C���ԍ�
'    Dim strWork         As String   '��ƃG���A
'    Dim strVerData      As String   '�S�̃o�[�W����
'    Dim intCnt          As Integer  '�J�E���^�[
'    Dim lngErrCode      As Long     '�G���[�R�[�h
'
''*******************************
''VB�G���[����
'On Error GoTo Error_psVersionDisp
''*******************************
'
'    '�}�̏o�͖t�����s��
'    cmdOutput.Enabled = False
'
'    '���X�g������
'    LstFile.Clear
'
'    '�S�̃o�[�W����������
'    lblZenVer.Caption = "�S�̃o�[�W�����i���[�N�j:--.--.--.--" & vbCrLf & _
'                        "�@�@�@�@�@�@�@�i���s�j�@:--.--.--.--" & vbCrLf & _
'                        "�@�@�@�@�@�@�@�i���j    :--.--.--.--"
'
'    '��ƃG���A������
'    strWork = ""
'
'    '�S�̃o�[�W����������
'    strVerData = ""
'
'    'LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���p�X�쐬
'    strFilePath = PATH_LDU_APP & PATH_LDU_WORK & LDU_VER_FILE
'
'    bRet = True
'    '///////////////////////////////////////////////////////////////////////////////////////////
'    '/ ����DA:LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���쐬
'    '///////////////////////////////////////////////////////////////////////////////////////////
'    bRet = dllCreateIDU_LDUVerFile(mlngChkFolderType, strFilePath, PATH_LDU_APP)
'
'    'LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���쐬����
'    If bRet Then
'       '�uLD���[�e�B���e�B�o�[�W�����Ǘ���ʁF�o�[�W�������t�@�C���쐬����v���O�o��
'       Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, CREATE_FILE_OK, 0)
'    'LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���쐬���s
'    Else
'       '�uLD���[�e�B���e�B�o�[�W�����Ǘ���ʁF�o�[�W�������t�@�C���쐬�ُ�v���O�o��
'       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
'       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
'       Exit Sub
'    End If
'
'    'LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���̗L���m�F
'    If Len(Trim(Dir(strFilePath))) = 0 Then
'        Exit Sub
'    End If
'
'    'LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���̃t�@�C���ԍ����擾����B
'    intFileNo = FreeFile
'
'    'LD���[�e�B���e�B��ʕ\���p�o�[�W�����t�@�C���I�[�v��
'    Open strFilePath For Input As #intFileNo
'
'
'        '���[�N
'        Line Input #intFileNo, strWork
'
'        If (Trim(strWork) = "") Then
'            strVerData = "�S�̃o�[�W�����i���[�N�j�F--.--.--.--" & vbCrLf
'        Else
'            '�S�̃o�[�W����������쐬
'            strVerData = strVerData & strWork & vbCrLf
'        End If
'
'        '���s
'        Line Input #intFileNo, strWork
'        If (Trim(strWork) = "") Then
'            strVerData = strVerData & "�@�@�@�@�@�@�@�i���s�j�@�F--.--.--.--" & vbCrLf
'        Else
'            '�S�̃o�[�W����������쐬
'            strVerData = strVerData & strWork & vbCrLf
'        End If
'
'        '��
'        Line Input #intFileNo, strWork
'        If (Trim(strWork) = "") Then
'            strVerData = strVerData & "�@�@�@�@�@�@�@�i���j    �F--.--.--.--" & vbCrLf
'        Else
'            '�S�̃o�[�W����������쐬
'            strVerData = strVerData & strWork & vbCrLf
'        End If
'
'        '�S�̃o�[�W�����o��
'        lblZenVer.Caption = strVerData
'
'        strWork = ""
'
'        '���X�g�\�����ǂݍ��݁i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
'        Do While Not EOF(1)
'
'            Line Input #intFileNo, strWork
'
'            '���s�R�[�h�݂͓̂ǂ݂Ƃ΂�
'            If Trim(strWork) <> "" Then
'
'                '���X�g�ɏo��
'                LstFile.AddItem (strWork)
'
'            End If
'        Loop
'
'    '�t�@�C���N���[�Y
'    Close #intFileNo
'
'    '�}�̏o�͖t������
'    cmdOutput.Enabled = True
'
'    Exit Sub
'
''*******************************
''VB�G���[����
'Error_psVersionDisp:
'   '�uLD���[�e�B���e�B�o�[�W�����Ǘ���ʁF�o�[�W�������t�@�C���쐬�ُ�v���O�o��
'   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
'   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
''    �t�@�C���N���[�Y
'    Close #intFileNo
''*******************************
'End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfInstallSeitouseiChck
'//  �@�\����  : �O�����̓v���O��������f�[�^�������`�F�b�N����
'//  �@�\�T�v  : �O�����̓v���O��������f�[�^�������`�F�b�N�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.6.0.1) 2009-06-16   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��s��C��
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(2.6.0.1) 2010-11-16  REVISED BY [TCC] S.Terao
'//                 �t�@�C�����`�F�b�N�s��C��
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-03 REVISED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��y��ʃ`�F�b�N�@�\�ǉ��z
'//     REVISIONS :(EG20 V6.11.0.1) 2013-03-27 REVISED BY  [TCC] H.Kondoh
'//                 �}�̓����@�\�ύX�Ή�
'//                   ��ʂO�̏ꍇ���ُ�Ƃ���悤�ɕύX
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X)----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfInstallSeitouseiChck(sInputPass As String) As Boolean
    Dim lngFileListCnt As Long               '�t�@�C�����X�g��
    Dim strWork     As String                '��ƃG���A
    Dim iFileNumber As Integer               '���g�p�t�@�C���ԍ�
    Dim myLen As Long                        '������̒���
    Dim SysCodeTxt As String                 '�o�C�g�ϊ���(�S�p�����p)
    Dim lngSumRet As Long
    Dim i As Integer
    Dim sGetFileListName As String           '�t�@�C�����X�g���L�ڃt�@�C����
    Dim bRet As Boolean
    Dim lngCnt As Long
    Dim iRet   As Integer                    '�o�[�W�����`�F�b�NDLL�߂�l
    Dim iGouki As Integer                    '���@�ԍ�
    Dim sVersionInfoPath As String           '�o�[�W�������t�@�C��(���@��)
    Dim sSrcFileName As String               '�t�@�C�����X�g��
    Dim lngErrCode   As Long
    Dim intCheckKind As Integer              ' �`�F�b�N���     ' EG20 V6.9.0.1ADD
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD
    
    On Error Resume Next
    
    pfInstallSeitouseiChck = True
    
    '********************************
    '*�v�����������`�F�b�N
    '********************************
    '�O���}�̃t�H���_���t�@�C�������쐬
    sSrcFileName = sInputPass & MN_FILELIST
    '�O���}�̂̌���������
    'If Dir(sSrcFileName) <> "" Then     'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    'V1.20.0.1 ADD END
    Else
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
      
      '�t�@�C�������݂��Ȃ�
      MsgBox "�}�̓��ɁA�t�@�C�����X�g�����݂��܂���B", _
             vbOKOnly + vbExclamation, _
             "�����[�N �R�s�["
     '�u�����ް�ޮ݁F�t�@�C�����X�g�����v���O�o��
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)

      pfInstallSeitouseiChck = False
      Set objFso = Nothing    'V1.20.0.1 ADD
      Exit Function
    End If

   '����[�N��t�H���_����t�@�C�����X�g���擾����
    bRet = fReadFileList(sInputPass & MN_FILELIST)

    '�T���l�`�F�b�N
    For lngCnt = 0 To UBound(FileList) - 1
        If pfFileSumChk(sInputPass & FileList(lngCnt), lngSumRet) <> True Then

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
           
            '�T���l�ُ�
            If lngSumRet = SUM_CHK.SumErr Then
                'EG20 V30.1.0.1 DEL START
'               MsgBox "�T���l���ُ�ł��B" _
'                      & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
'                      vbOKOnly + vbExclamation, _
'                      "�������D�@ �o�[�W�����Ǘ�"
                'EG20 V30.1.0.1 DEL END
                'EG20 V30.1.0.1 ADD START
               MsgBox "�T���l���ُ�ł��B" _
                      & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
                      vbOKOnly + vbExclamation, _
                      "�V�����������D�@ �o�[�W�����Ǘ�"
                'EG20 V30.1.0.1 ADD END
            '�T���l�ُ�ȊO�ُ�
            ElseIf lngSumRet = SUM_CHK.SumErr_Else Then
                   '�u���[�N�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
                'EG20 V30.1.0.1 DEL START
'               MsgBox "�R�s�[�G���[���������܂����B" _
'                     & Chr(vbKeyReturn) & "�G���[�R�[�h��" _
'                     & str$(Err.Number), _
'                     vbOKOnly + vbExclamation, _
'                     "�������D�@ �o�[�W�����Ǘ�"
                'EG20 V30.1.0.1 DEL END
                'EG20 V30.1.0.1 ADD START
               MsgBox "�R�s�[�G���[���������܂����B" _
                     & Chr(vbKeyReturn) & "�G���[�R�[�h��" _
                     & str$(Err.Number), _
                     vbOKOnly + vbExclamation, _
                     "�V�����������D�@ �o�[�W�����Ǘ�"
                'EG20 V30.1.0.1 ADD END
            End If
            pfInstallSeitouseiChck = False
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_SUMCHK_ERROR, 0)
            Exit Function
        End If
    Next

    '�t�@�C�����ő�`�F�b�N
    If UBound(FileList) > FILECNT_MAX Then
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
       'EG20 V30.1.0.1 DEL START
'       MsgBox "�t�@�C����������𒴂��Ă��܂��B" _
'              & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
'              vbOKOnly + vbExclamation, _
'              "�������D�@ �o�[�W�����Ǘ�"
        'EG20 V30.1.0.1 DEL END
        'EG20 V30.1.0.1 ADD START
        MsgBox "�t�@�C����������𒴂��Ă��܂��B" _
              & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
              vbOKOnly + vbExclamation, _
              "�V�����������D�@ �o�[�W�����Ǘ�"
        'EG20 V30.1.0.1 ADD END
      pfInstallSeitouseiChck = False

      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)

      Exit Function
    End If
'V2.6.0.1 DEL START
'    '�t�@�C�����T�C�Y�`�F�b�N
'    lngFileListCnt = UBound(FileList)
'
'    On Error GoTo FileGetError
'
'    iFileNumber = FreeFile          '���g�p�̃t�@�C���ԍ����擾����
'
'    bRet = True
'
'    '�t�@�C�����X�g���I�[�v���B
'    Open sInputPass & MN_FILELIST For Input As #iFileNumber
'    For i = 0 To lngFileListCnt
'       If i = lngFileListCnt Then
'          Exit For
'       End If
'       '�t�@�C�������擾����B
'       Input #iFileNumber, strWork
'       '�t�@�C������`�Ȃ�
'       If strWork = "" Then
'          '���[�v����
'          MsgBox "�t�@�C�������ُ�ł��B" _
'                  & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
'          bRet = False
'          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'          Exit For
'       '�t�H�[�}�b�g�ُ�
'       ElseIf " " <> Mid(strWork, 2, 1) And Left$(strWork, 1) <> "/" Then
'          '���[�v����
'          MsgBox "�t�@�C�������ُ�ł��B" _
'                  & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
'           bRet = False
'           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'           Exit For
'       '�t�H�[�}�b�g�ُ�
'       ElseIf (InStr(strWork, ".") - 1) = -1 And Left$(strWork, 1) <> "/" Then
'           MsgBox "�t�@�C�������ُ�ł��B" _
'                  & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
'           bRet = False
'           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'           Exit For
'       '�u/*--�v���̃R�����g���͏���
'       ElseIf Left$(strWork, 1) = "/" Then
'               '�������Ȃ��B
'       Else
'          '�t�@�C�����݂̂𒊏o
'          sGetFileListName = Mid(strWork, 3, 16)
'          '�擾�t�@�C�����̃T�C�Y���擾
'          myLen = LenB(StrConv(Trim(sGetFileListName), vbFromUnicode))                                              '���p���Z�̃o�C�g�����擾
'          If FILE_NAME_MAX_SIZE < myLen Then
'            '13�o�C�g�ȏ�̏ꍇ
'            MsgBox "�t�@�C�������ُ�ł��B" _
'                   & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                   vbOKOnly + vbExclamation, _
'                   sJverName & "�������D�@ �o�[�W�����Ǘ�"
'            bRet = False
'            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'            Exit For
'           End If
'        End If
'     Next
'    '�t�@�C�����X�g���N���[�Y�B
'    Close #iFileNumber
'V2.6.0.1 DEL END
'V2.6.0.1 ADD START
    For i = 0 To UBound(FileList) - 1
       '�擾�t�@�C�����̃T�C�Y���擾
       myLen = LenB(StrConv(Trim(FileList(i)), vbFromUnicode))                                              '���p���Z�̃o�C�g�����擾
       If FILE_NAME_MAX_SIZE < myLen Then
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
          '�v���O���X�o�[����������
          Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
          
          '13�o�C�g�ȏ�̏ꍇ
          'EG20 V30.1.0.1 DEL START
'          MsgBox "�t�@�C�������ُ�ł��B" _
'                 & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  "�������D�@ �o�[�W�����Ǘ�"
          'EG20 V30.1.0.1 DEL END
          'EG20 V30.1.0.1 ADD START
          MsgBox "�t�@�C�������ُ�ł��B" _
                 & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
                  vbOKOnly + vbExclamation, _
                  "�V�����������D�@ �o�[�W�����Ǘ�"
         'EG20 V30.1.0.1 ADD END
                
           bRet = False
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
           Exit For
       End If
    Next
'V2.6.0.1 ADD END

' EG20 V6.9.0.1�y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zADD START
    If bRet = False Then
        pfInstallSeitouseiChck = bRet
        Exit Function
    End If

    For i = 0 To UBound(FileList) - 1
        ' �t�@�C�����X�g���̎�ʂ𒊏o
        'intCheckKind = CInt(Left$(FileListType(i), 1))     'EG20 V30.1.0.1 DEL
        intCheckKind = Asc(Left$(FileListType(i), 1))       'EG20 V30.1.0.1 ADD
'EG20 V6.11.0.1 DEL Start
'        If ((gintProgramJudgeKind(FolderSyubetu) = intCheckKind) Or _
'            (intCheckKind = ProgramJudgeKind.JUDGE_NOCHECK)) Then
'            ' �f�[�^��ʑI�𕔂̑I����e�ƃt�@�C�����X�g���̎�ʂ̔�r���ʂ��u��v�v�A��������
'            ' �t�@�C�����X�g���̎�ʂ��u�`�F�b�N�Ȃ��v
'            ' ���`�F�b�N���ʐ���
'EG20 V6.11.0.1 DEL End
'EG20 V6.11.0.1 ADD Start
        If (gintProgramJudgeKind(FolderSyubetu) = intCheckKind) Then
            ' �f�[�^��ʑI�𕔂̑I����e�ƃt�@�C�����X�g���̎�ʂ̔�r���ʂ��u��v�v
            ' ���`�F�b�N���ʐ���
'EG20 V6.11.0.1 ADD End
            bRet = True
        Else
            ' ��L�ȊO
            ' ���`�F�b�N���ʈُ�
            bRet = False
            ' �v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            ' ���b�Z�[�W�\��
            'EG20 V30.1.0.1 DEL START
'            MsgBox "�I�������f�[�^��ʂƃC���X�g�[�����ނ�" & Chr(vbKeyReturn) _
'                     & "��v���܂���", _
'                   vbOKOnly + vbExclamation, _
'                   "�������D�@ �o�[�W�����Ǘ�"
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.1.0.1 ADD START
            MsgBox "�I�������f�[�^��ʂƃC���X�g�[�����ނ�" & Chr(vbKeyReturn) _
                     & "��v���܂���", _
                   vbOKOnly + vbExclamation, _
                   "�V�����������D�@ �o�[�W�����Ǘ�"
            'EG20 V30.1.0.1 ADD END
            
            ' �G���[���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_PRGKIND_ERROR, 0)
            Exit For
        End If
    Next
' EG20 V6.9.0.1�y�ʎY�Ή��F��ʃ`�F�b�N�@�\�ǉ��zADD END

    pfInstallSeitouseiChck = bRet
Exit Function

FileGetError:
    If iFileNumber > 0 Then
       Close #iFileNumber
    End If
    iFileNumber = 0
    pfInstallSeitouseiChck = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fSelectFile
'//  �@�\����  : �o�[�W�����`�F�b�N�t�@�C����
'//  �@�\�T�v  : �Ώۃo�[�W�����`�F�b�N�t�@�C�������擾����
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-18 REVISED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fSelectFile() As String
 
 'If gStrCurrentForm = sFormName_EJVer Then      'EG20 V30.1.0.1 DEL
    '�o�[�W�����`�F�b�N�t�@�C������ݒ肷��B
    Select Case FolderSyubetu
       Case 0 '����CPU-Pro
            fSelectFile = EG30_HANTEI_CPU_CHK_FILE
       
       Case 1 '���C��CPU-Pro
            fSelectFile = EG30_MAIN_CPU_CHK_FILE
       
       Case 2 '�T�uCPU-Pro
            fSelectFile = EG30_SUB_CPU_CHK_FILE
       
       Case 3 '���C��CPU-OS
            fSelectFile = EG30_MAIN_OS_CHK_FILE
     
     End Select
'EG20 V30.1.0.1 DEL START
'  Else
'    '�o�[�W�����`�F�b�N�t�@�C������ݒ肷��B
'    Select Case FolderSyubetu
'       Case 0 '����CPU-Pro
'             fSelectFile = NHANTEI_CPU_CHK_FILE
'
'       Case 1 '���C��CPU-Pro
'            fSelectFile = NMAIN_CPU_CHK_FILE
'
'       Case 2 '�T�uCPU-Pro
'            fSelectFile = NSUB_CPU_CHK_FILE
'
'       Case 3 '���C��CPU-OS
'            fSelectFile = NMAIN_OS_CHK_FILE
'
'    End Select
'   End If
'EG20 V30.1.0.1 DEL END

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fNewVersion
'//  �@�\����  : �ŐV�o�[�W��������
'//  �@�\�T�v  : �ŐV(���[�N)�o�[�W�������A���s(���s)�o�[�W�����ɓo�^
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sPath
'//  �@�@�@    : Integer�@ iFolder
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�v�����������`�F�b�N�����ǉ�
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�P�Ή��s��C��
'//                 �t�F�[�Y�R�Ή��@�@�퐳�����`�F�b�N�����ǉ�
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.8.0.1) 2012-04-17  CODED BY  [TCC] H.Sugimoto
'//                �y�c��:�ێ�^���̐ؑ֌��ʒʒm�Ή��z
'//     REVISIONS :(EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                �y�^���\�����P�Ή��z
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fNewVersion() As Boolean
    Dim bRet As Boolean                      '�߂�l
    Dim lngCnt                  As Long      '�J�E���^�[
    Dim sSrcFileName            As String    '���[�N�t�H���_���t�@�C�����X�g
    Dim sFileName As String
    Dim lngErrCode As Long                   '�G���[�R�[�h
    'V1.4.0.1 ADD START
    Dim lngFileListCnt As Long               '�t�@�C�����X�g��
    Dim strWork     As String                '��ƃG���A
    Dim iFileNumber As Integer               '���g�p�t�@�C���ԍ�
    Dim myLen As Long                        '������̒���
    Dim SysCodeTxt As String                 '�o�C�g�ϊ���(�S�p�����p)
    Dim lngSumRet As Long
    Dim i As Integer
    Dim sGetFileListName As String           '�t�@�C�����X�g���L�ڃt�@�C����
    'V1.4.0.1 ADD END
    Dim iKansiAplChk As Integer              '�A�v���N���`�F�b�N�߂�l�@'V1.6.0.1 ADD
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD
    
    Dim sCorner As String                    '�R�[�i�[�ԍ�
    Dim sGatePath As String                  '�R�[�i�[�ԍ��t�t�@�C���p�X
    Dim sFilePath As String                  '�t�@�C���t�@�C���p�X
    
    On Error Resume Next
    
    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sGatePath = PATH_N_GATE & sCorner

    sFilePath = sGatePath & FolderName(0, FolderSyubetu)

    '����[�N��t�H���_�̃t�@�C�����X�g����������
    '���[�N�t�H���_���t�@�C�������쐬
'    sSrcFileName = FolderName(0, FolderSyubetu) & "\" & MN_FILELIST
    sSrcFileName = sFilePath & "\" & MN_FILELIST
    '�t�@�C���̌���������
    'If Dir(sSrcFileName) <> "" Then     'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    If objFso.FileExists(sSrcFileName) = True Then
      Set objFso = Nothing
    'V1.20.0.1 ADD END
    Else
      
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
      '�v���O���X�o�[����������
      Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
      '�t�@�C�������݂��Ȃ�
      MsgBox "�u���[�N�v�t�H���_���� " & TitleBox(FolderSyubetu) & "�ɁA" _
             & Chr(vbKeyReturn) & "�t�@�C�����X�g�����݂��܂���B", _
             vbOKOnly + vbExclamation, _
             TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
     '�u�����ް�ޮ݁F�t�@�C�����X�g�����v���O�o��
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)

      fNewVersion = False
      Set objFso = Nothing    'V1.20.0.1 ADD
      Exit Function
    End If
  
    '����[�N��t�H���_����t�@�C�����X�g���擾����
    'bRet = fReadFileList(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)�@'V1.8.0.1 DEL
    
    bRet = pfSeitouseiChck    'V1.4.0.1�@ADD
    '�����v���O��������f�[�^�������`�F�b�N���s��(�Ώۃt�@�C���FHAN_KUKA.KUK)
'    bRet = fDataFileCheck(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST) 'V1.4.0.1�@DEL
'V1.8.0.1 ADD START
    '����[�N��t�H���_����t�@�C�����X�g���A�o�^�t�@�C�������J�E���g����
    If bRet = True Then
'       bRet = fReadFileList(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)
       bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
    End If
'V1.8.0.1 ADD END

  If bRet = True Then
    '�����t�H���_���̃t�@�C����S�č폜����
     If sOldFolderRemove <> True Then
'        Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2�ǉ�    EG20 V3.6.0.1�폜
        Call pubfuncErrorOccur(MN_FOLD_OLD)          ' EG20 V3.6.0.1�ǉ�
         fNewVersion = False
         Exit Function
     End If

    '�����t�H���_���̃t�@�C������s��t�H���_�̓��e�ɒu������
    If sCopyNOWtoOLD <> True Then
'        Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2�ǉ�    EG20 V3.6.0.1�폜
        Call pubfuncErrorOccur(MN_FOLD_OLD)          ' EG20 V3.6.0.1�ǉ�
        fNewVersion = False
        Exit Function
    End If

    '����s��t�H���_���̃t�@�C���𢃏�[�N��t�H���_�̓��e�ɒu������
    If sCopyWRKtoNOW <> True Then
        Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2�ǉ�
        fNewVersion = False
        Exit Function
    End If
    
' EG20 V3.0.0.2 �ǉ��J�n
    ' ���D�@���ʃG���A�X�V����
    Call pubfuncCommonAreaUpdate
' EG20 V3.0.0.2 �ǉ��I��
 
    '�����o�[�W�������X�V�v�����[�����Ǘ��v���Z�X�֑��M����B
    'V1.6.0.1�@ADD�@START
    '�Ď��ՋN��/���N���`�F�b�N���s���B�`�F�b�N��Ԃɂ�菈��������s���B
     iKansiAplChk = CheckAppStart(PROC_KANRI)
     If iKansiAplChk <> 0 Then
        '�Ď��ՋN�����F�Ǘ��v���Z�X�Ɏ����o�[�W�������X�V�v�����[���𑗐M����B
    'V1.6.0.1 ADD END
      'If gStrCurrentForm = sFormName_EJVer Then     'EG20 V30.1.0.1 DEL
         'psVersionUpdateReqest (ML_REQUEST_EGATE)      'EG20 V30.1.0.1 DEL
         psVersionUpdateReqest (ML_REQUEST_EG30GATE)       'EG20 V30.1.0.1 ADD
      'EG20 V30.1.0.1 DEL START
'      Else
'         psVersionUpdateReqest (ML_REQUEST_NGATE)
'      End If
      'EG20 V30.1.0.1 DEL END
    'V1.6.0.1 ADD START
    Else
        '�Ď��Ֆ��N�����F���D�@�o�[�W�����X�V�������ʂɐ����ݒ肷��B
        gintGateVerInfUpdRes = MailSts.stsNormal
    End If
    'V1.6.0.1 ADD END

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '���D�@�o�[�W�����X�V��������
    If gintGateVerInfUpdRes = MailSts.stsNormal Then
' EG20 V5.8.0.1�폜�J�n
'        ' �^����ԍX�V                                              ' EG20 V5.5.0.1�ǉ�
'        Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1�ǉ�
' EG20 V5.8.0.1�폜�I��
' EG20 V5.8.0.1�ǉ��J�n
        ' �^����ԍX�V
        Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_KIRIKAE)
' EG20 V5.8.0.1�ǉ��I��
'        Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iTab_index + 1)   ' EG20 V5.6.0.1�ǉ�           ' EG20 V5.11.0.1�폜
        Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iTab_index + 1, gintUnkaiKind(FolderSyubetu))    ' EG20 V5.11.0.1�ǉ�
        '����
        MsgBox "�u���[�N�v�t�H���_�̓��e��,�u���s�v�t�H���_�ɓo�^���āA" _
                & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & " �̍ŐV�̃o�[�W�����Ƃ��܂����B", _
                vbOKOnly + vbExclamation, _
                TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
        fNewVersion = True
    Else
        '�ُ�
        'If gStrCurrentForm = sFormName_EJVer Then      ' EG20 V30.1.0.1 DEL
            'EG20 V30.1.0.1 DEL START
'           MsgBox "���D�@�̃o�[�W�����쐬�ňُ킪�������܂����B", _
'                  vbOKOnly + vbExclamation, _
'                  "�������D�@ �o�[�W�����Ǘ�"
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.1.0.1 ADD START
           MsgBox "���D�@�̃o�[�W�����쐬�ňُ킪�������܂����B", _
                  vbOKOnly + vbExclamation, _
                  "�V�����������D�@ �o�[�W�����Ǘ�"
            'EG20 V30.1.0.1 ADD END
        ' EG20 V30.1.0.1 DEL START
'        Else
'         MsgBox "���D�@�̃o�[�W�����쐬�ňُ킪�������܂����B", _
'                 vbOKOnly + vbExclamation, _
'                 "�������D�@ �o�[�W�����Ǘ�"
'        End If
        ' EG20 V30.1.0.1 DEL END
        
        fNewVersion = False
    End If
  
    fNewVersion = True
  Else
    fNewVersion = False
  End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfSeitouseiChck
'//  �@�\����  : �v���O��������f�[�^�������`�F�b�N����
'//  �@�\�T�v  : �v���O��������f�[�^�������`�F�b�N�������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL :(1.4.0.1) 2009-03-17   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��@�v�����������`�F�b�N����
'//     REVISIONS :(1.6.0.1) 2009-06-16  REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή��s��C��
'//                 �t�F�[�Y�R�Ή��@�@�퐳�����`�F�b�N�ǉ�
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfSeitouseiChck() As Boolean
    Dim lngFileListCnt As Long               '�t�@�C�����X�g��
    Dim strWork     As String                '��ƃG���A
    Dim iFileNumber As Integer               '���g�p�t�@�C���ԍ�
    Dim myLen As Long                        '������̒���
    Dim SysCodeTxt As String                 '�o�C�g�ϊ���(�S�p�����p)
    Dim lngSumRet As Long
    Dim i As Integer
    Dim sGetFileListName As String           '�t�@�C�����X�g���L�ڃt�@�C����
    Dim bRet As Boolean
    Dim lngCnt As Long
    Dim iRet   As Integer                    '�o�[�W�����`�F�b�NDLL�߂�l
    Dim iGouki As Integer                    '���@�ԍ�
    Dim sVersionInfoPath As String           '�o�[�W�������t�@�C��(���@��)
    Dim iCnt             As Integer          '���@�J�E���^�[�@V1.6.0.1�@ADD
    
    Dim sCorner As String                    '�R�[�i�[�ԍ�
    Dim sGatePath As String                  '�R�[�i�[�ԍ��t�t�@�C���p�X
    Dim sFilePath As String                  '�t�@�C���t�@�C���p�X
    
    On Error Resume Next
    
    pfSeitouseiChck = True
   
    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sGatePath = PATH_N_GATE & sCorner

    '********************************
    '*�v�����������`�F�b�N
    '********************************
    '�����v���O��������f�[�^�������`�F�b�N���s��(�Ώۃt�@�C���FHAN_KUKA.KUK)
'    bRet = fDataFileCheck(FolderName(0, FolderSyubetu) & "\" & MN_FILELIST)
    
    sFilePath = sGatePath & FolderName(0, FolderSyubetu)
    
    bRet = fDataFileCheck(sFilePath & "\" & MN_FILELIST)
    If bRet = False Then
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
       '�v���O���X�o�[����������
       Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
       If sNGSts <> "" And sNGKoumoku <> "" Then
          'EG20 V30.1.0.1 DEL START
'          MsgBox "�^���f�[�^�������`�F�b�N�ُ�(" & sNGSts & "�F" & sNGKoumoku & "�j", _
'                 vbOKOnly + vbExclamation, _
'                 "�������D�@ �o�[�W�����Ǘ�"
          'EG20 V30.1.0.1 DEL END
          'EG20 V30.1.0.1 ADD START
          MsgBox "�^���f�[�^�������`�F�b�N�ُ�(" & sNGSts & "�F" & sNGKoumoku & "�j", _
                 vbOKOnly + vbExclamation, _
                 "�V�����������D�@ �o�[�W�����Ǘ�"
          'EG20 V30.1.0.1 ADD END
       Else
          MsgBox "�ُ�I�����܂����B", _
                 vbOKOnly + vbExclamation, _
                 TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
       End If
'       Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2�ǉ�
       Call pubfuncErrorOccur(MN_FOLD_WRK)          ' EG20 V3.6.0.1�ǉ�
       pfSeitouseiChck = False
       Exit Function
    End If

' EG20 V3.0.0.2�ǉ��J�n
    ' ���D�@���ʔ��菈��
    bRet = pubfuncCommonGateCheck(MN_FOLD_WRK)
    If bRet = False Then
       pfSeitouseiChck = False
       Exit Function
    End If

' EG20 V3.0.0.2�ǉ��I��

'V1.6.0.1 DEL START
'    '�T���l�`�F�b�N
'    For lngCnt = 0 To UBound(FileList) - 1
'        If pfFileSumChk(FolderName(0, FolderSyubetu) & "\" & FileList(lngCnt), lngSumRet) <> True Then
'            '�T���l�ُ�
'            If lngSumRet = SUM_CHK.SumErr Then
'               MsgBox "�T���l���ُ�ł��B" _
'                      & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
'                      vbOKOnly + vbExclamation, _
'                      sJverName & "�������D�@ �o�[�W�����Ǘ�"
'            '�T���l�ُ�ȊO�ُ�
'            ElseIf lngSumRet = SUM_CHK.SumErr_Else Then
'               MsgBox "�ُ�I�����܂����B", _
'                     vbOKOnly + vbExclamation, _
'                     TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
'            End If
'            pfSeitouseiChck = False
'            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_SUMCHK_ERROR, 0)
'            Exit Function
'        End If
'    Next
'
'    '�t�@�C�����ő�`�F�b�N
'    If UBound(FileList) > FILECNT_MAX Then
'       MsgBox "�t�@�C����������𒴂��Ă��܂��B" _
'              & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
'              vbOKOnly + vbExclamation, _
'              sJverName & "�������D�@ �o�[�W�����Ǘ�"
'      pfSeitouseiChck = False
'
'      Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)
'
'      Exit Function
'    End If
'
'    '�t�@�C�����T�C�Y�`�F�b�N
'    lngFileListCnt = UBound(FileList)
'
'    On Error GoTo FileGetError
'
'    iFileNumber = FreeFile          '���g�p�̃t�@�C���ԍ����擾����
'    '�t�@�C�����X�g���I�[�v���B
'    Open FolderName(0, FolderSyubetu) & "\" & MN_FILELIST For Input As #iFileNumber
'    For i = 0 To lngFileListCnt
'       If i = lngFileListCnt Then
'          Exit For
'       End If
'       '�t�@�C�������擾����B
'       Input #iFileNumber, strWork
'       '�t�@�C������`�Ȃ�
'       If strWork = "" Then
'          '���[�v����
'          MsgBox "�t�@�C�������ُ�ł��B" _
'                  & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
'          bRet = False
'          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'          Exit For
'       '�t�H�[�}�b�g�ُ�
'       ElseIf " " <> Mid(strWork, 2, 1) Then
'          '���[�v����
'          MsgBox "�t�@�C�������ُ�ł��B" _
'                  & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
'           bRet = False
'           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'           Exit For
'       ElseIf (InStr(strWork, ".") - 1) = -1 Then
'           MsgBox "�t�@�C�������ُ�ł��B" _
'                  & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                  vbOKOnly + vbExclamation, _
'                  sJverName & "�������D�@ �o�[�W�����Ǘ�"
'           bRet = False
'           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'           Exit For
'       Else
'          '�t�@�C�����݂̂𒊏o
'          sGetFileListName = Mid(strWork, 3, 16)
'          '�擾�t�@�C�����̃T�C�Y���擾
'          myLen = LenB(StrConv(sGetFileListName, vbFromUnicode))                                              '���p���Z�̃o�C�g�����擾
'          If FILE_NAME_MAX_SIZE < myLen Then
'            '13�o�C�g�ȏ�̏ꍇ
'            MsgBox "�t�@�C�������ُ�ł��B" _
'                   & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                   vbOKOnly + vbExclamation, _
'                   sJverName & "�������D�@ �o�[�W�����Ǘ�"
'            bRet = False
'            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'            Exit For
'           End If
'        End If
'     Next
'    '�t�@�C�����X�g���N���[�Y�B
'    Close #iFileNumber
'V1.6.0.1 DEL END
'V1.11.0.1 DEL START
'    If gStrCurrentForm = sFormName_EJVer Then
''V1.6.0.1 ADD�@START
'   For iCnt = 1 To MAX_GATE_NO
'      'EG-R�����̂݁F�����o�[�W�����`�F�b�NDLL����
'      iGouki = pfGetGoukiNo(iCnt)
'      If iGouki <> 0 Then
''V1.6.0.1 ADD�@END
'       'iGouki = pfGetGoukiNo 'V1.6.0.1 DEL
'       sVersionInfoPath = Replace(GATE_VERSION_INFO_FILE, "##", Format(iGouki, "0#"))
'
'       'iRet = dllVerChk(E_EPRO1WRK & "\\" & GATE_VERSION_KANRI_FILE, PATH_GATE & sVersionInfoPath, PATH_HOSHU_LOG & GATE_VERSION_NGLIST_FILE)�@�@�@�@�@�@�@�@�@'V1.6.0.1�@DEL
'       iRet = dllVerChk(FolderName(0, FolderSyubetu) & "\" & GATE_VERSION_KANRI_FILE, PATH_GATE & sVersionInfoPath, PATH_HOSHU_LOG & GATE_VERSION_NGLIST_FILE)  'V1.6.0.1�@ADD
'       If iRet = 1 Then
'          bRet = True
'       Else
'          bRet = False
'          MsgBox "�ُ�I�����܂����B", _
'                 vbOKOnly + vbExclamation, _
'                 TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
'          'V1.6.0.1 ADD START
'           pfSeitouseiChck = False
'           Exit Function
'          'V1.6.0.1 ADD END
'       End If
'       End If 'V1.6.0.1 ADD
'      Next 'V1.6.0.1 ADD
'    End If
''V1.6.0.1 ADD START
'V1.11.0.1 DEL END
    '�@�퐳�����`�F�b�N(�Ώۃt�@�C���FXX_GATEY.VEF�@XX:���[�U�[���@Y�F�f�[�^���)
'    bRet = fKishuCheck(FolderName(0, FolderSyubetu) & "\")
    sFilePath = sGatePath & FolderName(0, FolderSyubetu)
    
    bRet = fKishuCheck(sFilePath & "\")
    
    If bRet = False Then
       Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2�ǉ�
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
       '�v���O���X�o�[����������
       Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
       MsgBox "�ُ�I�����܂����B", _
                  vbOKOnly + vbExclamation, _
                  TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
       pfSeitouseiChck = False
       Exit Function
    End If
'V1.6.0.1 ADD END

    pfSeitouseiChck = bRet
Exit Function

FileGetError:
    If iFileNumber > 0 Then
       Close #iFileNumber
    End If
    iFileNumber = 0
    Call pubfuncErrorOccur(MN_FOLD_NOW)         ' EG20 V3.0.0.2�ǉ�
    pfSeitouseiChck = False
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sOldFolderRemove
'//  �@�\����  : ���t�H���_���t�@�C���폜����
'//  �@�\�T�v  : ���t�H���_���̃t�@�C�����폜����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sOldFolderRemove() As Boolean
    Dim MyName As String                '�t�@�C����
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                     '�t�@�C���I�u�W�F�N�g
    'V1.20.0.1 ADD END
    
    Dim sCorner As String                      '�R�[�i�[�ԍ�
    Dim sGatePath As String                    '�R�[�i�[�ԍ��t�t�@�C���p�X
    Dim sFilePath As String                    '�t�@�C���t�@�C���p�X
    
   '�߂�l������
    sOldFolderRemove = True
    
    On Error GoTo ErrorHandler          '�G���[�n���h���̓o�^
    
    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sGatePath = PATH_N_GATE & sCorner
 
    '�u���s�v�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
'    gstrMyPath = FolderName(2, FolderSyubetu) & "\"
    gstrMyPath = sGatePath & FolderName(2, FolderSyubetu) & "\"
    'V1.20.0.1 DEL START
'    MyName = Dir(gstrMyPath & "*.*", vbNormal)  ' �ŏ��̃f�B���N�g������Ԃ��܂��B
'    Do While MyName <> ""                   ' ���[�v���J�n���܂��B
'        ' ���݂̃f�B���N�g���Ɛe�f�B���N�g���͖������܂��B
'        If MyName <> "." And MyName <> ".." Then
'            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
'            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
'                '�t�@�C�����폜����
'                Kill gstrMyPath & MyName
'            End If
'        End If
'        MyName = Dir        ' ���̃f�B���N�g������Ԃ��܂��B
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(gstrMyPath).files  '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then      '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
                '�t�@�C�����폜����
                Kill gstrMyPath & MyName
            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    
    Exit Function           '�������I������

ErrorHandler:   ' �G���[�������[�`���B
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    '�u���[�N�����s�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
     MsgBox "�ُ�I�����܂����B", _
           vbOKOnly + vbExclamation, _
           TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
    '�u�����ް�ޮ݁F���t�H���_̧�ٍ폜�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OLDFILE_DELETE_ERROR, lngErrCode)

    sOldFolderRemove = False
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sNowFolderRemove
'//  �@�\����  : ���s�t�H���_���̃t�@�C���폜����
'//  �@�\�T�v  : ���s�t�H���_���̃t�@�C�����폜����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sNowFolderRemove() As Boolean
    Dim MyName As String                '�t�@�C����
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                     '�t�@�C���I�u�W�F�N�g
    'V1.20.0.1 ADD END

    Dim sCorner As String                 '�R�[�i�[�ԍ�
    Dim sGatePath As String               '�R�[�i�[�ԍ��t�t�@�C���p�X
    Dim sFilePath As String
    On Error GoTo ErrorHandler          '�G���[�n���h���̓o�^

    '�����l�ݒ�
    sNowFolderRemove = True
    
    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sGatePath = PATH_N_GATE & sCorner

    sFilePath = sGatePath & FolderName(1, FolderSyubetu)
    
    '�u���s�v�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
'    gstrMyPath = FolderName(1, FolderSyubetu) & "\"
    gstrMyPath = sFilePath & "\"
    'V1.20.0.1 DEL START
'    MyName = Dir(gstrMyPath & "*.*", vbNormal)  ' �ŏ��̃f�B���N�g������Ԃ��܂��B
'    Do While MyName <> ""                   ' ���[�v���J�n���܂��B
'        ' ���݂̃f�B���N�g���Ɛe�f�B���N�g���͖������܂��B
'        If MyName <> "." And MyName <> ".." Then
'            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
'            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
'
'                Kill gstrMyPath & MyName        '�t�@�C�����폜����
'
'            End If
'        End If
'        MyName = Dir        ' ���̃f�B���N�g������Ԃ��܂��B
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(gstrMyPath).files  '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then      '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then

                Kill gstrMyPath & MyName        '�t�@�C�����폜����

            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END

    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    '�u�������s�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
    MsgBox "�ُ�I�����܂����B", _
           vbOKOnly + vbExclamation, _
           TitleBox(FolderSyubetu) & "  �������s �R�s�["

    '�u�����ް�ޮ݁F���s�t�H���_̧�ٍ폜�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOWFILE_DELETE_ERROR, lngErrCode)

    sNowFolderRemove = False
    
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sWrkFolderRemove
'//  �@�\����  : ���[�N�t�H���_���t�@�C���폜����
'//  �@�\�T�v  : ���[�N�t�H���_���̃t�@�C�����폜����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-17  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                �y�^���\�����P�Ή��z
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sWrkFolderRemove() As Boolean
    Dim MyName As String                '�t�@�C����
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim lngPgmHanteiStsWork As Long     '�v���O���������ԁi���[�N�j   ' EG20 V3.6.0.1�ǉ�
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                     '�t�@�C���I�u�W�F�N�g
    'V1.20.0.1 ADD END
    
    Dim sCorner As String               '�R�[�i�[�ԍ�
    Dim sGatePath As String             '�R�[�i�[�ԍ��t�t�@�C���p�X
    
    On Error GoTo ErrorHandler          '�G���[�n���h���̓o�^

    '�����l�ݒ�
    sWrkFolderRemove = True
   
    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sGatePath = PATH_N_GATE & sCorner
  
    '���[�N�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
'    gstrMyPath = FolderName(0, FolderSyubetu) & "\"
    gstrMyPath = sGatePath & FolderName(0, FolderSyubetu) & "\"
    
    'V1.20.0.1 DEL START
'    MyName = Dir(gstrMyPath & "*.*", vbNormal)  ' �ŏ��̃f�B���N�g������Ԃ��܂��B
'    Do While MyName <> ""                   ' ���[�v���J�n���܂��B
'        ' ���݂̃f�B���N�g���Ɛe�f�B���N�g���͖������܂��B
'        If MyName <> "." And MyName <> ".." Then
'            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
'            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
'                '�t�@�C�����폜����
'                Kill gstrMyPath & MyName
'            End If
'        End If
'        MyName = Dir        ' ���̃f�B���N�g������Ԃ��܂��B
'    Loop
    'V1.20.0.1 DEL END
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(gstrMyPath).files  '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then      '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
                '�t�@�C�����폜����
                Kill gstrMyPath & MyName
            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END

' EG20 V3.6.0.1�ǉ��J�n
    '�Ď��ݒ�G���A�u�v���O��������ُ��ԁi���[�N�j�v�̏�Ԃ��擾����
    lngPgmHanteiStsWork = pfGetKansiSet(IdKansiSet.PG_HANTEI_ERR_STS_WORK)

    '�u�v���O��������ُ��ԁi���[�N�j�v�i����j
    Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_ERR_STS_WORK, ErrCode.Normal)
    
    '�ω����������ꍇ�A�u��ԕω��ʒm�v�𑗐M����
    If lngPgmHanteiStsWork <> ErrCode.Normal Then
        Call sSendMailStsChgInf(MailSts.stsNormal, lngPgmHanteiStsWork)
    End If
' EG20 V3.6.0.1�ǉ��I��
    
' EG20 V5.11.0.1�폜�J�n
'' EG20 V5.8.0.1�폜�J�n
''    ' �^����ԍX�V                                              ' EG20 V5.5.0.1�ǉ�
''    Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI)         ' EG20 V5.5.0.1�ǉ�
'' EG20 V5.8.0.1�폜�I��
'' EG20 V5.8.0.1�ǉ��J�n
'    ' �^����ԍX�V
'    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_NASHI)
'' EG20 V5.8.0.1�ǉ��I��
'    Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI, iTab_index + 1)   ' EG20 V5.6.0.1�ǉ�
' EG20 V5.11.0.1�폜�I��
' EG20 V5.11.0.1�ǉ��J�n
    ' �^����ԍX�V
    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_CLEAR)
    Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_NASHI, iTab_index + 1, gintUnkaiKind(FolderSyubetu))    ' EG20 V5.11.0.1�ǉ�
' EG20 V5.11.0.1�ǉ��I��
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    '�u���[�N�N���A����I���v�|�b�v�A�b�v��ʕ\��
    MsgBox "�u���[�N�v�t�H���_���� " & TitleBox(FolderSyubetu) & "���A" _
               & Chr(vbKeyReturn) & "�S�č폜���܂����B", _
               vbOKOnly + vbExclamation, _
               TitleBox(FolderSyubetu) & "  ���[�N �N���A"

    Exit Function '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    Call pubfuncErrorOccur(MN_FOLD_WRK)             ' EG20 V3.6.0.1�ǉ�
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '�u���[�N�N���A�ُ�I���v�|�b�v�A�b�v��ʕ\��
     MsgBox "�ُ�I�����܂����B", _
           vbOKOnly + vbCritical, _
           "���[�N �N���A"
           
   '�u�����ް�ޮ݁Fܰ�̫���̧�ٍ폜�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_WRKFILE_DELETE_ERROR, lngErrCode)
           
    sWrkFolderRemove = False
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sCopyNOWtoOLD
'//  �@�\����  : ���s�o�[�W�����ۑ�����
'//  �@�\�T�v  : ���s�t�H���_���̃t�@�C�����A���t�H���_�ɃR�s�[����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sCopyNOWtoOLD() As Boolean
    Dim MyName As String                '�t�@�C����
    Dim sSrcFileName As String          '�R�s�[���t�@�C���̃t���p�X��
    Dim sDstFileName As String          '�R�s�[��t�@�C���̃t���p�X��
    Dim iResponse As Integer            'MsgBox�{�^���R�[�h
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                     '�t�@�C���I�u�W�F�N�g
    'V1.20.0.1 ADD END
    
    Dim sCorner As String                      '�R�[�i�[�ԍ�
    Dim sGatePath As String                    '�R�[�i�[�ԍ��t�t�@�C���p�X
    
    On Error GoTo ErrorHandler              '�G���[�n���h���ݒ�
  
    '�߂�l������
    sCopyNOWtoOLD = True
   
       ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sGatePath = PATH_N_GATE & sCorner

    '���s�t�H���_���̃f�B���N�g���̖��O��\�����܂��B
'    gstrMyPath = FolderName(1, FolderSyubetu) & "\"
    gstrMyPath = sGatePath & FolderName(1, FolderSyubetu) & "\"
    'V1.20.0.1 DEL START
'    MyName = Dir(gstrMyPath & "*.*", vbNormal)  ' �ŏ��̃f�B���N�g������Ԃ��܂��B
'    Do While MyName <> ""                   ' ���[�v���J�n���܂��B
'        ' ���݂̃f�B���N�g���Ɛe�f�B���N�g���͖������܂��B
'        If MyName <> "." And MyName <> ".." Then
'            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
'            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then
'
'                '���s�t�H���_���t�@�C�������쐬����
'                sSrcFileName = gstrMyPath & MyName
'
'                '���t�H���_���t�@�C�������쐬����
'                sDstFileName = FolderName(2, FolderSyubetu) & "\" & MyName
'
'                '���[�N�t�H���_���̃t�@�C�������s�t�H���_�ɃR�s�[����
'                FileCopy sSrcFileName, sDstFileName
'
'            End If
'        End If
'        MyName = Dir        ' ���̃f�B���N�g������Ԃ��܂��B
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(gstrMyPath).files  '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then      '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��
            If (GetAttr(gstrMyPath & MyName) And vbDirectory) <> vbDirectory Then

                '���s�t�H���_���t�@�C�������쐬����
                sSrcFileName = gstrMyPath & MyName

                '���t�H���_���t�@�C�������쐬����
'                sDstFileName = FolderName(2, FolderSyubetu) & "\" & MyName
                sDstFileName = sGatePath & FolderName(2, FolderSyubetu) & "\" & MyName

                '���[�N�t�H���_���̃t�@�C�������s�t�H���_�ɃR�s�[����
                FileCopy sSrcFileName, sDstFileName

            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    
    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
           
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
           ' �u���[�N�����s�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
            MsgBox "�ُ�I�����܂����B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
    
    sCopyNOWtoOLD = False
    
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sCopyWRKtoNOW
'//  �@�\����  : �ŐV�o�[�W�����R�s�[
'//  �@�\�T�v  : ���[�N�t�H���_���̃t�@�C�����A���s�t�H���_�ɃR�s�[
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��iPASSINF�R�s�[�Ή��j
'//     REVISIONS :(EG20 V3.5.0.1) 2012-02-07  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sCopyWRKtoNOW() As Boolean
    
    Dim i As Integer                '�J�E���^
    Dim sSrcFileName As String      '�R�s�[���t�@�C����
    Dim sDstFileName As String      '�R�s�[��t�@�C����
    Dim bError As Boolean           '�t���O
    Dim bRet As Boolean             '�߂�l
    
    Dim sCorner As String                '�R�[�i�[�ԍ�
    Dim sGatePath As String              '�R�[�i�[�ԍ��t�t�@�C���p�X
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD

    On Error GoTo ErrorHandler      '�G���[�n���h���̓o�^
  
    '�߂�l������
    sCopyWRKtoNOW = True
    
    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sGatePath = PATH_N_GATE & sCorner

    '****************************
    '* �t�@�C�����X�g���R�s�[���� *
    '****************************
      
'    sSrcFileName = FolderName(0, FolderSyubetu) & "\" & MN_FILELIST
    sSrcFileName = sGatePath & FolderName(0, FolderSyubetu) & "\" & MN_FILELIST
                                    '���[�N�t�H���_���t�@�C�������쐬����
'    sDstFileName = FolderName(1, FolderSyubetu) & "\" & MN_FILELIST
    sDstFileName = sGatePath & FolderName(1, FolderSyubetu) & "\" & MN_FILELIST
                                    '���s�t�H���_���t�@�C�������쐬����
    'If Dir(sSrcFileName) <> "" Then     '�t�@�C���̌���������   'V1.20.0.1 DEL
    If objFso.FileExists(sSrcFileName) = True Then     '�t�@�C���̌���������   'V1.20.0.1 ADD
        '�t�@�C�����X�g���u���[�N�v�t�H���_����u���s�v�t�H���_�ɃR�s�[����
        FileCopy sSrcFileName, sDstFileName
    Else                                '�t�@�C�������݂��Ȃ�
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
     '�u���[�N�t�H���_�t�@�C�����X�g�Ȃ��v�|�b�v�A�b�v��ʕ\��
     MsgBox "�u���[�N�v�t�H���_���� " & TitleBox(FolderSyubetu) & "�ɁA" _
             & Chr(vbKeyReturn) & "�t�@�C�����X�g�����݂��܂���B", _
             vbOKOnly + vbExclamation, _
             TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
     sCopyWRKtoNOW = False
     Set objFso = Nothing    'V1.20.0.1 ADD
     Exit Function                   '�������I������
    End If

    bError = False                  '�G���[�t���O���u�U�v�ɂ���
    For i = 0 To UBound(FileList) - 1
                                    '�t�@�C�����X�g�ꗗ�����J��Ԃ�
'        sSrcFileName = FolderName(0, FolderSyubetu) & "\" & FileList(i)
        sSrcFileName = sGatePath & FolderName(0, FolderSyubetu) & "\" & FileList(i)
                                    '���[�N�t�H���_���t�@�C�������쐬����
'        sDstFileName = FolderName(1, FolderSyubetu) & "\" & FileList(i)
        sDstFileName = sGatePath & FolderName(1, FolderSyubetu) & "\" & FileList(i)
                                    '���s�t�H���_���t�@�C�������쐬����

        '���[�N�t�H���_���̃t�@�C�������s�t�H���_�ɃR�s�[����
        'If Dir(sSrcFileName) <> "" Then     '�t�@�C���̌���������   'V1.20.0.1 DEL
        If objFso.FileExists(sSrcFileName) = True Then   '�t�@�C���̌���������   'V1.20.0.1 ADD
            '�t�@�C�����u���[�N�v�t�H���_����u���s�v�t�H���_�ɃR�s�[����
            FileCopy sSrcFileName, sDstFileName
        End If
    Next
    
    Set objFso = Nothing    'V1.20.0.1 ADD
    
' EG20 V3.0.0.2�ǉ��J�n
    If pfuncCopyPASSINF(iTab_index, MN_FOLD_WRK) = False Then
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
' EG20 V3.5.0.1�ǉ��J�n
        MsgBox "�ُ�I�����܂����B", _
                vbOKOnly + vbExclamation, _
                TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
' EG20 V3.5.0.1�ǉ��I��
        sCopyWRKtoNOW = False
    End If
' EG20 V3.0.0.2�ǉ��I��
    
    Exit Function                           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    '�G���[����������΁A�_�C�A���O���o���悤�ɂ���B�i�G���[�R�[�h�ɂ�����炸�j
    'Select Case Err.Number
    '    Case 53 '�u���[�N�����s�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
            MsgBox "�ُ�I�����܂����B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  ���[�N�����s �R�s�["
            
            sCopyWRKtoNOW = False
            Set objFso = Nothing    'V1.20.0.1 ADD
            Exit Function
    '    Case Else
                ' ���̃G���[�����������ɋL�q���܂��B
    'End Select
    sCopyWRKtoNOW = False
    Set objFso = Nothing    'V1.20.0.1 ADD
End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fDataFileCheck
'//  �@�\����  : �����v���O��������f�[�^�������`�F�b�N����
'//  �@�\�T�v  : �ΏۂƂȂ�HAN_KUKA.KUK�L���`�F�b�N���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.11.0.1) 2009-10-23  CODED   BY [TCC] D.Yamashita
'//                 �E�t�F�[�Y�R�c�����ڑΉ��@�ُ펞�N���[�Y�����ǉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fDataFileCheck(sFileList As String) As Boolean
    Dim iFileNumber As Integer      '�t�@�C���ԍ�
    Dim sFileName As String         '�t�@�C����
    Dim iListCnt As Integer         '�t�@�C���i�[��
    Dim sFolderPath As String       'HAN_KUKA.KUK�t�H���_�p�X�p
    Dim sHANKUKAPath As String      'HAN_KUKA.KUK�t���p�X�p
     
    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�

    iListCnt = 0
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)

    iFileNumber = FreeFile   '���g�p�̃t�@�C���ԍ����擾����

    Open sFileList For Input Access Read As #iFileNumber    '�t�@�C�����X�g�̃I�[�v��
    Do While Not EOF(iFileNumber)                           '�t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
        Line Input #iFileNumber, sFileName                  '�f�[�^��ǂݍ��݂܂��B
        If sFileName <> "" And Left$(sFileName, 1) <> "/" Then                '�t�@�C���������݂���
            iListCnt = iListCnt + 1                         '�t�@�C�����̃J�E���^���A�b�v����
            ReDim Preserve FileList(iListCnt)               '�t�@�C�����i�[�G���A���g������
            ReDim Preserve FileListType(iListCnt)           '�t�@�C�����i�[�G���A���g������
            FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
            FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 16)
                                                            '�t�@�C�������t�@�C�����i�[�G���A�ɃZ�b�g
            If HANKUKA_KUK = FileList(iListCnt - 1) Then
               'HAN_KUKA.KUK�t�@�C�����L�����ꍇ�A�f�[�^�������`�F�b�N���s���B
               psFolderPathGet sFileList, sFolderPath
               sHANKUKAPath = sFolderPath & HANKUKA_KUK
               If fHankukaChck(sHANKUKAPath) = False Then
                 '�f�[�^�������`�F�b�N�ُ펞�́A�߂�l��False��ݒ肷��B
                  fDataFileCheck = False
                  Close #iFileNumber        '�t�@�C������܂��B   'V1.11.0.1 ADD
                  Exit Function
               End If
            End If
        End If
  Loop
  
  Close #iFileNumber        '�t�@�C������܂��B

  fDataFileCheck = True     '�߂�l�𐳏�Ƃ���

    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:               ' �G���[�������[�`���B
    fDataFileCheck = False  '�߂�l���G���[�Ƃ���
    Close #iFileNumber      '�t�@�C������܂��B
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fKishuCheck
'//  �@�\����  : �����v���O��������f�[�^�������`�F�b�N����
'//  �@�\�T�v  : �ΏۂƂȂ�f�[�^�̋@�퐳�����`�F�b�N���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yamazaki
'//                ���[�N�����s�R�s�[�ł̋@�퐳�����`�F�b�N�ύX
'//                Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fKishuCheck(sFileList As String) As Boolean
    Dim sKisyu       As String * 8     '�擾�@�햼
    Dim sMyName      As String         '�@�퐳�����`�F�b�N���X�g�t�@�C����
    Dim sFileName    As String         '�t�@�C�����X�g�L�ڃt�@�C����
    Dim sChkFileName As String         '�@�퐳�����`�F�b�N�t�@�C���p�X
    Dim sVerChkFile  As String         '�o�[�W�����`�F�b�N�t�@�C����
    
    Dim lLen         As Long           '�t�@�C���T�C�Y
    Dim lPos         As Long           '�o�[�W�������i�[�ʒu
           
    Dim i            As Integer        '�J�E���^�[
    Dim iCnt         As Integer        '�o�^���R�[�h��
    Dim iListCnt     As Integer        '�t�@�C���i�[��
    Dim iFileNumber  As Integer        '�t�@�C���ԍ�

    Dim bRet         As Boolean        '�@�퐳�����`�F�b�N����

    Dim uHeder       As MN_HEDER       '�w�b�_���i�[�G���A
    Dim uFotter      As MN_FOOT        '�t�b�^���i�[�G���A
    
    Dim sChkData As String             '��r�������o    'V1.20.0.1 ADD
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD
    
    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�
     
    '������
    iCnt = 0
    iListCnt = 0
    iFileNumber = 0
    fKishuCheck = False
        
    ReDim Preserve FileList(iListCnt)
    ReDim Preserve FileListType(iListCnt)
    
    '�o�[�W�����f�[�^(�@�퐳�����`�F�b�N���X�g�t�@�C���p�X)�쐬
    sVerChkFile = fSelectFile
    
    '�t�@�C�����擾�s��=�@�퐳�����`�F�b�N�t�@�C���Ȃ�
    If sVerChkFile = "" Then
       '�������`�F�b�N���s���K�v�Ȃ����߁A�����Ԃ��B
       fKishuCheck = True
       Set objFso = Nothing    'V1.20.0.1 ADD
       Exit Function
    End If
    
    sMyName = sFileList & sVerChkFile
    
    'If Dir(sMyName) <> "" Then              '�t�@�C�������݂���?     'V1.20.0.1 DEL
    If objFso.FileExists(sMyName) = True Then    '�t�@�C�������݂���?  'V1.20.0.1 ADD
       
       iFileNumber = FreeFile               '���g�p�̃t�@�C���ԍ����擾����
       
       Open sMyName For Input Access Read As #iFileNumber     '�o�[�W�����f�[�^�̃I�[�v��
       
       '�f�[�^�ǂݍ���
       Line Input #iFileNumber, sFileName
          
       '�ǂݍ��݃f�[�^���A�w�b�_���������B
       sFileName = Mid(sFileName, Len(uHeder) - 3)
       
       '�t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
       Do While Not EOF(iFileNumber)
          
          '�ǂݍ��݁B
          Line Input #iFileNumber, sFileName
           
           '�擾��񂪁u/�v�ȍ~�̃R�����g�Ȃ�ΏۊO�B
           '�f�[�^���{���ȊO�Ȃ�ΏۊO
           '�f�[�^���{���݂̂̏ꍇ�̂݁A�t�@�C�����擾���s���B
           If sFileName <> "" And Left$(sFileName, 1) <> "/" _
                              And " " = Mid(sFileName, 2, 1) Then   '�t�@�C���������݂���
              iListCnt = iListCnt + 1                         '�t�@�C�����̃J�E���^���A�b�v����
              ReDim Preserve FileList(iListCnt)               '�t�@�C�����i�[�G���A���g������
              ReDim Preserve FileListType(iListCnt)           '�t�@�C�����i�[�G���A���g������
              '�t�@�C�������t�@�C�����i�[�G���A�ɃZ�b�g
              FileListType(iListCnt - 1) = UCase(Trim$(Left$(sFileName, 18)))
              FileList(iListCnt - 1) = Mid$(FileListType(iListCnt - 1), 3, 12)
              '�o�^���R�[�h�����J�E���g
              iCnt = iCnt + 1
            End If
       Loop
       
       Close #iFileNumber                                     '�t�@�C������܂��B
       iFileNumber = 0
    Else
       '�t�@�C�������݂��Ȃ��ꍇ�F�������`�F�b�N���s��Ȃ��B
       fKishuCheck = True
       Set objFso = Nothing    'V1.20.0.1 ADD
       Exit Function
    End If
    
    'V1.20.0.1 ADD  START
    If iCnt = 0 Then
       '�t�@�C�����X�g�R�[�h�����݂��Ȃ��ꍇ�F�������`�F�b�N���s��Ȃ��B
       fKishuCheck = True
       Set objFso = Nothing    'V1.20.0.1 ADD
       Exit Function
    End If
    'V1.20.0.1 ADD  END
    
    '�t�@�C���@�퐳�����`�F�b�N���s���B
    For i = 0 To iCnt - 1
         '�`�F�b�N�Ώۃt�@�C���p�X�쐬
        sChkFileName = sFileList & FileList(i)
    
        'If Dir(sChkFileName) <> "" Then              '�t�@�C�������݂���?  'V1.20.0.1 DEL
        If objFso.FileExists(sChkFileName) = True Then  '�t�@�C�������݂���?   'V1.20.0.1 ADD
            
            lLen = FileLen(sChkFileName)             '�t�@�C���T�C�Y�̎擾

            iFileNumber = FreeFile                   '���g�p�̃t�@�C���ԍ����擾����
            '�t�@�C���̃I�[�v�����s���B
            Open sChkFileName For Binary Access Read As #iFileNumber
            '�t�b�^���̎擾
            Get #iFileNumber, lLen - Len(uFotter) + 1, uFotter
            
            Close #iFileNumber                       '�t�@�C������܂�
            iFileNumber = 0
            
            '�@�햼�Z�b�g
            sKisyu = uFotter.sKisyu
            
            sChkData = "" '�������@'V1.20.0.1 ADD
            
' EG20 V3.0.0.2 �ǉ��J�n
            '�������o
            'sChkData = Left(sKisyu, Len(EG20_JIKAI_KISHU))     'EG20 V30.1.0.1 DEL
            sChkData = Left(sKisyu, Len(EG30_JIKAI_KISHU))      'EG20 V30.1.0.1 ADD
            'If EG20_JIKAI_KISHU = sChkData Then            'EG20 V30.1.0.1 DEL
            If EG30_JIKAI_KISHU = sChkData Then             'EG20 V30.1.0.1 ADD
                bRet = True  '�@�퐳�����F����
            Else
                bRet = False '�@�퐳�����F�ُ�
                fKishuCheck = bRet
                Set objFso = Nothing    'V1.20.0.1 ADD
                Exit Function
            End If
' EG20 V3.0.0.2 �ǉ��I��
            
' EG20 V3.0.0.2 �폜�J�n
'            '�����`�F�b�N
'            If gStrCurrentForm = sFormName_EJVer Then
'               'EG-R������
'               'If EGR_JIKAI_KISHU = Trim(sKisyu) Then  'V1.20.0.1 DEL
'               'V1.20.0.1 ADD START
'               '�������o
'               sChkData = Left(sKisyu, Len(EGR_JIKAI_KISHU))
'               If EGR_JIKAI_KISHU = sChkData Then
'               'V1.20.0.1 ADD END
'                   bRet = True  '�@�퐳�����F����
'               Else
'                   bRet = False '�@�퐳�����F�ُ�
'                   fKishuCheck = bRet
'                   Set objFso = Nothing    'V1.20.0.1 ADD
'                   Exit Function
'               End If
'            Else
'               'NEG������
'               'If NEG_JIKAI_KISHU = Trim(sKisyu) Then    'V1.20.0.1 DEL
'               'V1.20.0.1 ADD START
'               '�������o
'               sChkData = Left(sKisyu, Len(NEG_JIKAI_KISHU))
'               If NEG_JIKAI_KISHU = sChkData Then
'               'V1.20.0.1 ADD END
'                   bRet = True  '�@�퐳�����F����
'               Else
'                   bRet = False '�@�퐳�����F�ُ�
'                   fKishuCheck = bRet
'                   Set objFso = Nothing    'V1.20.0.1 ADD
'                   Exit Function
'               End If
'            End If
' EG20 V3.0.0.2 �폜�I��

        End If
    Next

  fKishuCheck = bRet
  
  Set objFso = Nothing    'V1.20.0.1 ADD
  
 Exit Function

ErrorHandler:
   If iFileNumber <> 0 Then
       Close #iFileNumber                                     '�t�@�C������܂��B
   End If
    
   '�߂�l���ُ�Ƃ���
   fKishuCheck = False
       
   Set objFso = Nothing    'V1.20.0.1 ADD

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fHankukaChck
'//  �@�\����  : HAN_KUKA.KUK�������`�F�b�N����
'//  �@�\�T�v  : �ΏۂƂȂ�HAN_KUKA.KUK�̓��e���`�F�b�N����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sFileList�@[IN]�t�@�C�����X�g�̃t���p�X��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.4.0.1) 2009-04-06   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�Ή�
'//     REVISIONS :(1.16.0.1) 2009-12-21   REVISED BY [TCC] S.Terao
'//                 �s��Ή�
'//     REVISIONS :(V2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 EG-R(KK)�@������Ή��@KUK�������`�F�b�N�ύX
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fHankukaChck(sFilePath As String) As Boolean
    Dim iFileNumber As Integer           '�t�@�C���ԍ�
    Dim i As Integer
    Dim lSts As Long
    Dim sKeyName As String
    Dim lPos As Long                     '�o�[�W�������i�[�ʒu
    Dim lLen As Long                     '�t�@�C���T�C�Y
    'Dim uFooter As MN_FOOT          '�t�b�^���i�[�G���A      'EG20 V30.1.0.1 DEL
    Dim uFooter As MN_KAN_FOOT          '�t�b�^���i�[�G���A   'EG20 V30.1.0.1 ADD
'    Dim uHeder As MN_FOOT           '�w�b�_���i�[�G���A     'V1.4.0.1 DEL
    Dim sDateTime As String
    Dim j As Integer
    Dim lngErrCode As Long          '�G���[�R�[�h
    'V1.4.0.1 ADD START
    Dim uHeder As HAN_KUKA_KUK_HEADER       '�w�b�_���i�[�G���A
    Dim sGetInfo As String * MAX_PATH_SIZE  'INI�t�@�C���擾�p
    Dim sChkFileData As String
    Dim iMojisu As Integer
    
    'V1.16.0.1 ADD Start
    Dim bChkSts As Boolean              '�`�F�b�N���ʃt���O
    Dim sChkData As String              '��r�������o
    'V1.16.0.1 ADD End
    
   '�������F����(�u�����N�j
    sNGSts = ""
    sNGKoumoku = ""
    'V1.4.0.1 ADD END
    Dim oFs As New FileSystemObject 'V2.5.0.1 ADD
    
    fHankukaChck = False
    
'V2.5.0.1 ADD START
 '�t�@�C���L���`�F�b�N���s���B
 If oFs.FileExists(sFilePath) = False Then
    '�t�@�C����������ΐ������`�F�b�N���s��Ȃ��B
    fHankukaChck = True
    Set oFs = Nothing
    Exit Function
 End If
'V2.5.0.1 ADD END

 'V1.4.0.1 DEL START
'   For i = 0 To INI_MAX
'      '�w�b�_�F���Ғl�@�햼�擾
'      sKeyName = Format(HEDER_KISHU_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     HAN_KUKA_DATA.sHederKisyu(i), _
'                                     Len(HAN_KUKA_DATA.sHederKisyu(i)), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'        Exit For
'      End If
'
'      '�w�b�_�F���Ғl�t�@�C�����擾
'      sKeyName = Format(HEDER_FILE_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     HAN_KUKA_DATA.sHederFile(i), _
'                                     Len(HAN_KUKA_DATA.sHederFile(i)), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'        Exit For
'      End If
'
'      '�t�b�^�F���Ғl�@�햼�擾
'      sKeyName = Format(FOTTER_KISHU_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     HAN_KUKA_DATA.sFotterKisyu(i), _
'                                     Len(HAN_KUKA_DATA.sFotterKisyu(i)), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'        Exit For
'      End If
'
'      '�t�b�^�F���Ғl�t�@�C�����擾
'      sKeyName = Format(FOTTER_FILE_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     HAN_KUKA_DATA.sFotterFile(i), _
'                                     Len(HAN_KUKA_DATA.sFotterFile(i)), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'        Exit For
'      End If
'
'    Next i
    'V1.4.0.1 DEL END
    'V1.4.0.1 ADD START
    '������
    For i = 0 To INI_MAX - 1
        HAN_KUKA_DATA.sHederKisyu(i) = ""
        HAN_KUKA_DATA.sHederFile(i) = ""
        HAN_KUKA_DATA.sFotterKisyu(i) = ""
        HAN_KUKA_DATA.sFotterFile(i) = ""
    Next
    For i = 0 To INI_MAX - 1
      '�w�b�_�F���Ғl�@�햼�擾
      sKeyName = Format(HEDER_KISHU_NAME & "0" & i + 1)
      'EG20 V30.1.0.1 DEL START
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     sGetInfo, _
'                                     Len(sGetInfo), _
'                                     GATE_HANTEI_CHK_FILE)
      'EG20 V30.1.0.1 DEL END
      'EG20 V30.1.0.1 ADD START
      lSts = GetPrivateProfileString(EG30_HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      'EG20 V30.1.0.1 ADD END
      If lSts = False Then
       
      Else
        HAN_KUKA_DATA.sHederKisyu(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
      '�w�b�_�F���Ғl�t�@�C�����擾
      sKeyName = Format(HEDER_FILE_NAME & "0" & i + 1)
      'EG20 V30.1.0.1 DEL START
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     sGetInfo, _
'                                     Len(sGetInfo), _
'                                     GATE_HANTEI_CHK_FILE)
      'EG20 V30.1.0.1 DEL END
      'EG20 V30.1.0.1 ADD START
      lSts = GetPrivateProfileString(EG30_HANTEI_CHK, _
                                     sKeyName, _
                                     "", _
                                     sGetInfo, _
                                     Len(sGetInfo), _
                                     GATE_HANTEI_CHK_FILE)
      'EG20 V30.1.0.1 ADD END
      If lSts = False Then
        
      Else
         HAN_KUKA_DATA.sHederFile(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
      End If
      'EG20 V30.1.0.1 DEL START�i�V�����̓t�b�^�����j
'      '�t�b�^�F���Ғl�@�햼�擾
'      �t�b�^�͂Ȃ�
'      sKeyName = Format(FOTTER_KISHU_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     sGetInfo, _
'                                     Len(sGetInfo), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'
'      Else
'        HAN_KUKA_DATA.sFotterKisyu(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
'      End If
'      '�t�b�^�F���Ғl�t�@�C�����擾
'      �t�b�^�͖���
'      sKeyName = Format(FOTTER_FILE_NAME & "0" & i + 1)
'      lSts = GetPrivateProfileString(HANTEI_CHK, _
'                                     sKeyName, _
'                                     "", _
'                                     sGetInfo, _
'                                     Len(sGetInfo), _
'                                     GATE_HANTEI_CHK_FILE)
'      If lSts = False Then
'
'      Else
'        HAN_KUKA_DATA.sFotterFile(i) = Left$(sGetInfo, (InStr(sGetInfo, vbNullChar) - 1))
'      End If
      'EG20 V30.1.0.1 DEL END
    Next i
    'V1.4.0.1 ADD END

    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�
    
    'HAN_KUKA.KUK�t�@�C���T�C�Y�擾
    lLen = FileLen(sFilePath)
    
    '���g�p�̃t�@�C���ԍ����擾����
    iFileNumber = FreeFile
    
    'V1.4.0.1 DEL START
'    'HAN_KUKA.KUK�t�@�C�����I�[�v������B
'    Open sFilePath For Input Access Read As #iFileNumber
'
'    'HAN_KUKA.KUK�t�@�C���̃w�b�_�����擾����B
''    Get #iFileNumber, 1, uHeder
    'V1.4.0.1 DEL END

    'V1.4.0.1 ADD START
    'HAN_KUKA.KUK�t�@�C�����I�[�v������B
    Open sFilePath For Binary Access Read As #iFileNumber
            
    'HAN_KUKA.KUK�t�@�C���̃w�b�_�����擾����B
    Get #iFileNumber, 1, uHeder
    'V1.4.0.1 ADD END

   'HAN_KUKA.KUK�t�@�C���̃t�b�^�����擾����B
    Get #iFileNumber, lLen - Len(uFooter) + 1, uFooter

    'HAN_KUKA.KUK�t�@�C�����N���[�Y����B
    Close #iFileNumber
    
    iFileNumber = 0                          'V1.4.0.1 ADD
'V1.4.0.1 DEL START
    '�@�햼/�t�@�C�����`�F�b�N
'    For i = 0 To 5
'       '�w�b�_���F�@�햼�`�F�b�N
'       If uHeder.sKisyu <> HAN_KUKA_DATA.sHederKisyu(i) Then
'          Exit Function
'       End If
'       '�w�b�_���F�t�@�C�����`�F�b�N
'       If uHeder.sFileName <> HAN_KUKA_DATA.sHederFile(i) Then
'          Exit Function
'       End If
'       '�t�b�^���F�@�햼�`�F�b�N
'       If uFooter.sKisyu <> HAN_KUKA_DATA.sFotterKisyu(i) Then
'          Exit Function
'       End If
'       '�t�b�^���F�t�@�C�����`�F�b�N
'       If uFooter.sFileName <> HAN_KUKA_DATA.sFotterFile(i) Then
'          Exit Function
'       End If
'     Next
'V1.4.0.1 DEL END
   'V1.4.0.1 ADD START
   '�w�b�_���F�@�햼�`�F�b�N
   iMojisu = InStr(uHeder.sKisyuName, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uHeder.sKisyuName, 1)
   Else
     sChkFileData = Mid(uHeder.sKisyuName, 1, iMojisu)
   End If
'V1.16.0.1 DEL START
'    For i = 0 To INI_MAX - 1
'      If sChkFileData <> HAN_KUKA_DATA.sHederKisyu(i) Then
'         If i = INI_MAX - 1 Then
'            '�@�햼���Ғl�S�s��v�F
'            sNGSts = ERROR_HEDER
'            sNGKoumoku = KISHU_NAME_ERROR
'            GoTo ErrorHandler
'         End If
'      Else
'        Exit For
'      End If
'    Next
'V1.16.0.1 DEL END
'V1.16.0.1 ADD START
    bChkSts = False
    For i = 0 To INI_MAX - 1
       If HAN_KUKA_DATA.sHederKisyu(i) <> "" Then
          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sHederKisyu(i)))
          If sChkData = HAN_KUKA_DATA.sHederKisyu(i) Then
             bChkSts = True
           Exit For
          End If
      End If
    Next
    '�`�F�b�N���ʃt���O����
    If bChkSts = False Then
       '�@�햼���Ғl�S�s��v�F
        sNGSts = ERROR_HEDER
        sNGKoumoku = KISHU_NAME_ERROR
         GoTo ErrorHandler
    End If
'V1.16.0.1 ADD END

   '�w�b�_���F�t�@�C�����`�F�b�N
   iMojisu = InStr(uHeder.sProgrumName, " ") - 1
   If iMojisu < 0 Then
     sChkFileData = Mid(uHeder.sProgrumName, 1)
   Else
     sChkFileData = Mid(uHeder.sProgrumName, 1, iMojisu)
   End If

'V1.16.0.1 DEL START
'   For i = 0 To INI_MAX - 1
'       If sChkFileData <> HAN_KUKA_DATA.sHederFile(i) Then
'         If i = INI_MAX - 1 Then
'            '�t�@�C�������Ғl�S�s��v�F
'            sNGSts = ERROR_HEDER
'            sNGKoumoku = FILE_NAME_ERRORE
'            GoTo ErrorHandler
'         End If
'      Else
'         Exit For
'      End If
'    Next
'V1.16.0.1 DEL END
'V1.16.0.1 ADD START
    bChkSts = False
    For i = 0 To INI_MAX - 1
       If HAN_KUKA_DATA.sHederFile(i) <> "" Then
          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sHederFile(i)))
          If sChkData = HAN_KUKA_DATA.sHederFile(i) Then
             bChkSts = True
           Exit For
          End If
      End If
    Next
    '�`�F�b�N���ʃt���O����
    If bChkSts = False Then
       '�@�햼���Ғl�S�s��v�F
        sNGSts = ERROR_HEDER
        sNGKoumoku = FILE_NAME_ERRORE
         GoTo ErrorHandler
    End If
'V1.16.0.1 ADD END
    
   '�쐬���t�`�F�b�N
   '�w�b�_���F�쐬���t�����l���ǂ���
    sDateTime = ""
    For j = 0 To 3
        sDateTime = sDateTime & Right$("0" & Hex(uHeder.byWriteTime(j)), 2)
    Next
    For j = 4 To 5
        sDateTime = sDateTime & Right$("0" & Hex(uHeder.byWriteTime(j)), 2)
    Next
    If IsNumeric(sDateTime) = False Then
       sNGSts = ERROR_HEDER
       sNGKoumoku = CREATE_DATA_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    
    '�o�[�W�������l�`�F�b�N
    If IsNumeric(uHeder.sVersion) = False Then
       sNGSts = ERROR_HEDER
       sNGKoumoku = VERSION_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    
   'EG20 V30.1.0.1 DEL START �V�����ł̓t�b�^���ɋ@�햼�A�f�[�^���͑��݂��Ȃ��B
'   '�t�b�^���F�@�햼�`�F�b�N
'   iMojisu = InStr(uFooter.sKisyu, " ") - 1
'   If iMojisu < 0 Then
'     sChkFileData = Mid(uFooter.sKisyu, 1)
'   Else
'     sChkFileData = Mid(uFooter.sKisyu, 1, iMojisu)
'   End If
''V1.16.0.1 DEL START
''    For i = 0 To INI_MAX - 1
''      If sChkFileData <> HAN_KUKA_DATA.sFotterKisyu(i) Then
''         If i = INI_MAX - 1 Then
''             '�@�햼���Ғl�S�s��v�F
''             sNGSts = ERROR_FOTTER
''             sNGKoumoku = KISHU_NAME_ERROR
''             GoTo ErrorHandler
''          End If
''       Else
''         Exit For
''       End If
''    Next
''V1.16.0.1 DEL END
''V1.16.0.1 ADD START
'   bChkSts = False
'    For i = 0 To INI_MAX - 1
'       If HAN_KUKA_DATA.sFotterKisyu(i) <> "" Then
'          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sFotterKisyu(i)))
'          If sChkData = HAN_KUKA_DATA.sFotterKisyu(i) Then
'             bChkSts = True
'           Exit For
'          End If
'      End If
'    Next
'    '�`�F�b�N���ʃt���O����
'    If bChkSts = False Then
'       '�@�햼���Ғl�S�s��v�F
'        sNGSts = ERROR_FOTTER
'        sNGKoumoku = KISHU_NAME_ERROR
'         GoTo ErrorHandler
'    End If
''V1.16.0.1 ADD END
'
'   '�t�b�^���F�t�@�C�����`�F�b�N
'   iMojisu = InStr(uFooter.sFileName, " ") - 1
'   If iMojisu < 0 Then
'     sChkFileData = Mid(uFooter.sFileName, 1)
'   Else
'     sChkFileData = Mid(uFooter.sFileName, 1, iMojisu)
'   End If
''V1.16.0.1 DEL START
''    For i = 0 To INI_MAX - 1
''       If sChkFileData <> HAN_KUKA_DATA.sFotterFile(i) Then
''          If i = INI_MAX - 1 Then
''             '�@�햼���Ғl�S�s��v�F
''             sNGSts = ERROR_FOTTER
''             sNGKoumoku = FILE_NAME_ERRORE
''             GoTo ErrorHandler
''          End If
''       Else
''         Exit For
''       End If
''    Next
''   'V1.4.0.1 ADD END
''V1.16.0.1 DEL END
''V1.16.0.1 ADD START
'   bChkSts = False
'    For i = 0 To INI_MAX - 1
'       If HAN_KUKA_DATA.sFotterFile(i) <> "" Then
'          sChkData = Left(sChkFileData, Len(HAN_KUKA_DATA.sFotterFile(i)))
'          If sChkData = HAN_KUKA_DATA.sFotterFile(i) Then
'             bChkSts = True
'           Exit For
'          End If
'      End If
'    Next
'    '�`�F�b�N���ʃt���O����
'    If bChkSts = False Then
'       '�@�햼���Ғl�S�s��v�F
'        sNGSts = ERROR_FOTTER
'        sNGKoumoku = FILE_NAME_ERRORE
'         GoTo ErrorHandler
'    End If
''V1.16.0.1 ADD END
    'EG20 V30.1.0.1 DEL END

'V1.4.0.1 DEL START
'   '�쐬���t�`�F�b�N
'   '�w�b�_���F�쐬���t�����l���ǂ���
'    sDateTime = ""
'    For j = 0 To 3
'        sDateTime = sDateTime & Right$("0" & Hex(uHeder.byWriteTime(j)), 2)
'    Next
'    sDateTime = sDateTime & " "
'    For j = 4 To 5
'        sDateTime = sDateTime & Right$("0" & Hex(uHeder.byWriteTime(j)), 2)
'    Next
'    If (Chr(sDateTime) >= "A" And Chr(sDateTime) <= "Z") And _
'        (Chr(sDateTime) >= "a" And Chr(sDateTime) <= "z") Then
'         Exit Function
'    End If
'V1.4.0.1 DEL END
      
    '�t�b�^���F�쐬���t�����l���ǂ���
     sDateTime = ""
     For j = 0 To 3
         sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
     Next
    'sDateTime = sDateTime & " " 'V1.4.0.1 DEL
     For j = 4 To 5
        sDateTime = sDateTime & Right$("0" & Hex(uFooter.byWriteTime(j)), 2)
    Next
    'V1.4.0.1 DEL START
'    If (Chr(sDateTime) >= "A" And Chr(sDateTime) <= "Z") And _
'       (Chr(sDateTime) >= "a" And Chr(sDateTime) <= "z") Then
'        Exit Function
'    End If
    'V1.4.0.1 DEL END
    
    'V1.4.0.1 ADD START
    If IsNumeric(sDateTime) = False Then
       sNGSts = ERROR_FOTTER
       sNGKoumoku = CREATE_DATA_ERROR
       GoTo ErrorHandler
       Exit Function
    End If
    'V1.4.0.1 ADD END
    'V1.4.0.1 DEL START
'      '�o�[�W�����l�`�F�b�N
'    '�w�b�_���F�o�[�W�����l�����l���ǂ���
'    If (Chr(uHeder.sVersion) >= "A" And Chr(uHeder.sVersion) <= "Z") And _
'        (Chr(uHeder.sVersion) >= "a" And Chr(uHeder.sVersion) <= "z") Then
'        Exit Function
'    End If
'
'    '�t�b�^���F�o�[�W�����l�����l���ǂ���
'    If (Chr(uFooter.sVersion) >= "A" And Chr(uFooter.sVersion) <= "Z") And _
'       (Chr(uFooter.sVersion) >= "a" And Chr(uFooter.sVersion) <= "z") Then
'        Exit Function
'    End If
    'V1.4.0.1 DEL END
    
    'EG20 V30.1.0.1 DEL START �V�����̃t�b�^���ɂ̓o�[�W�����͑��݂��Ȃ�
'    'V1.4.0.1 ADD START
'    '�o�[�W�����l�`�F�b�N
'    '�t�b�^���F�o�[�W�����l�����l���ǂ���
'    If IsNumeric(uFooter.sVersion) = False Then
'       sNGSts = ERROR_FOTTER
'       sNGKoumoku = VERSION_ERROR
'       GoTo ErrorHandler
'       Exit Function
'    End If
'    'V1.4.0.1 ADD END
    'EG20 V30.1.0.1 DEL END
    
    '�u�����ް�ޮ݁F�����`�F�b�N����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_FILESTS_OK, 0)
    
    '���ׂ�OK�̏ꍇ�ATRUE�ł�����B
    fHankukaChck = True

Exit Function 'V1.4.0.1 ADD
'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    'V1.4.0.1 ADD START
    If iFileNumber > 0 Then
       'HAN_KUKA.KUK�t�@�C�����N���[�Y����B
       Close #iFileNumber
    End If
    iFileNumber = 0
    'V1.4.0.1 ADD END
    
    '�u�����ް�ޮ݁F�����`�F�b�N�ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   ' Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_FILESTS_OK, 0) 'V1.4.0.1 DEL
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILESTS_ERROR, lngErrCode)  'V1.4.0.1 ADD
    fHankukaChck = False   '�߂�l���G���[�Ƃ���
    'HAN_KUKA.KUK�t�@�C�����N���[�Y����B
    'Close #iFileNumber                        'V1.4.0.1 DEL
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fOldVersion
'//  �@�\����  : ���o�[�W��������
'//  �@�\�T�v  : �ꐢ��O�̃o�[�W���������s(���s)�o�[�W�����ɕԂ��B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-29   REVISED BY [TCC] S.Terao
'//                �t�F�[�Y�R�Ή��@�Ǘ��ւ̃��[�����M�������u���[�N�����s�R�s�[�v���ɂ��킹��
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-18  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V5.6.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                �y�^���\�����P�Ή��z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fOldVersion() As Boolean
    Dim bRet As Boolean                     '�߂�l
    Dim lngCnt                  As Long     '�J�E���^�[
    Dim sSrcFileName            As String   '���t�H���_���t�@�C�����X�g
    Dim lngSumRet               As Long
    Dim lngErrCode              As Long     '�G���[�R�[�h
    Dim iKansiAplChk As Integer              '�A�v���N���`�F�b�N�߂�l�@'V1.6.0.1 ADD

    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD
    
    Dim sCorner As String                      '�R�[�i�[�ԍ�
    Dim sGatePath As String                    '�R�[�i�[�ԍ��t�t�@�C���p�X
    Dim sFilePath As String                    '�t�@�C���t�@�C���p�X
    
    On Error Resume Next
 
    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sGatePath = PATH_N_GATE & sCorner

   '���t�H���_���̃t�@�C�����X�g����������B
'    sSrcFileName = FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
    sSrcFileName = sGatePath & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
                                    '�u���v�t�H���_���t�@�C�������쐬����
    'If Dir(sSrcFileName) <> "" Then     '�t�@�C���̌���������  'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    If objFso.FileExists(sSrcFileName) = True Then
        Set objFso = Nothing
    'V1.20.0.1 ADD END
    Else                                '�t�@�C�������݂��Ȃ�
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        MsgBox "�u���v�t�H���_���� " & TitleBox(FolderSyubetu) & "�ɁA" _
                   & Chr(vbKeyReturn) & "�t�@�C�����X�g�����݂��܂���B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  �������s �R�s�["
        '�u�����ް�ޮ݁F�t�@�C�����X�g�����v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_NOTFOUND_FILELIST, lngErrCode)
 
        fOldVersion = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function                   '�������I������
    End If
    
    '�����t�H���_����t�@�C�����X�g���擾����
    sFilePath = sGatePath & FolderName(MN_FOLD_OLD, FolderSyubetu)

'    bRet = fReadFileList(FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST)
    bRet = fReadFileList(sFilePath & "\" & MN_FILELIST)
  
' EG20 V3.6.0.1 �y����TR-No.260�z�ǉ��J�n
    bRet = fDataFileCheck(sFilePath & "\" & MN_FILELIST)
    If bRet = False Then
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
       
       If sNGSts <> "" And sNGKoumoku <> "" Then
          'EG20 V30.1.0.1 DEL START
'          MsgBox "�^���f�[�^�������`�F�b�N�ُ�(" & sNGSts & "�F" & sNGKoumoku & "�j", _
'                 vbOKOnly + vbExclamation, _
'                 "�������D�@ �o�[�W�����Ǘ�"
          'EG20 V30.1.0.1 DEL END
          'EG20 V30.1.0.1 ADD START
          MsgBox "�^���f�[�^�������`�F�b�N�ُ�(" & sNGSts & "�F" & sNGKoumoku & "�j", _
                 vbOKOnly + vbExclamation, _
                 "�V�����������D�@ �o�[�W�����Ǘ�"
          'EG20 V30.1.0.1 ADD END
       Else
          MsgBox "�ُ�I�����܂����B", _
                 vbOKOnly + vbExclamation, _
                 TitleBox(FolderSyubetu) & "  �������s �R�s�["
       End If
       Call pubfuncErrorOccur(MN_FOLD_OLD)
       fOldVersion = False
       Exit Function
    End If
' EG20 V3.6.0.1 �y����TR-No.260�z�ǉ��I��
  
' EG20 V3.0.0.2 �ǉ��J�n
    If pubfuncCommonGateCheck(MN_FOLD_OLD) = False Then
        fOldVersion = False
       Exit Function
    End If
' EG20 V3.0.0.2 �ǉ��I��
  
    '����s��t�H���_���̃t�@�C����S�č폜����
    If sNowFolderRemove <> True Then
        Call pubfuncErrorOccur(MN_FOLD_NOW)     ' EG20 V3.0.0.2�ǉ�
        fOldVersion = False
        Exit Function
    End If
    
    '�����t�H���_���̃t�@�C������s��t�H���_�̓��e�ɒu������
    If sCopyOLDtoNOW <> True Then
        Call pubfuncErrorOccur(MN_FOLD_NOW)     ' EG20 V3.0.0.2�ǉ�
        fOldVersion = False
        Exit Function
    End If

    Call pubfuncCommonAreaUpdate                ' EG20 V3.0.0.2 �ǉ�

'V1.6.0.1 DEL START
'   '�����o�[�W�������X�V�v�����[�����Ǘ��v���Z�X�֑��M����B
'     If gStrCurrentForm = sFormName_EJVer Then
'        psVersionUpdateReqest (ML_REQUEST_EGATE)
'     Else
'        psVersionUpdateReqest (ML_REQUEST_NGATE)
'     End If
'V1.6.0.1 DEL END
'V1.6.0.1 ADD START
    '�����o�[�W�������X�V�v�����[�����Ǘ��v���Z�X�֑��M����B
    '�Ď��ՋN��/���N���`�F�b�N���s���B�`�F�b�N��Ԃɂ�菈��������s���B
     iKansiAplChk = CheckAppStart(PROC_KANRI)
     If iKansiAplChk <> 0 Then
        '�Ď��ՋN�����F�Ǘ��v���Z�X�Ɏ����o�[�W�������X�V�v�����[���𑗐M����B
      'If gStrCurrentForm = sFormName_EJVer Then        'EG20 V30.1.0.1 DEL
         'psVersionUpdateReqest (ML_REQUEST_EGATE)      'EG20 V30.1.0.1 DEL
         psVersionUpdateReqest (ML_REQUEST_EG30GATE)       'EG20 V30.1.0.1 ADD
      ' EG20 V30.1.0.1 DEL START
'      Else
'         psVersionUpdateReqest (ML_REQUEST_NGATE)
'      End If
      ' EG20 V30.1.0.1 DEL END
    Else
        '�Ď��Ֆ��N�����F���D�@�o�[�W�����X�V�������ʂɐ����ݒ肷��B
        gintGateVerInfUpdRes = MailSts.stsNormal
    End If
'V1.6.0.1 ADD END
     
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
     
     '���D�@�o�[�W�����X�V�����ُ�
    If gintGateVerInfUpdRes = MailSts.stsNormal Then
' EG20 V5.8.0.1�ǉ��J�n
        ' �^����ԍX�V
        Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_KIRIKAE)
' EG20 V5.8.0.1�ǉ��I��
'        Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iTab_index + 1)   ' EG20 V5.6.0.1�ǉ�           ' EG20 V5.11.0.1�폜
        Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_KIRIKAE, iTab_index + 1, gintUnkaiKind(FolderSyubetu))    ' EG20 V5.11.0.1�ǉ�
        '����
        MsgBox "�u���v�t�H���_�̓��e���A�u���s�v�t�H���_�ɖ߂��āA" _
                    & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & "�̈ꐢ��O�̃o�[�W�������A" _
                    & Chr(vbKeyReturn) & "���s�o�[�W�����Ƃ��܂����B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  �������s �R�s�["
        fOldVersion = True
    Else
        '�ُ�
        'If gStrCurrentForm = sFormName_EJVer Then      ' EG20 V30.1.0.1 DEL
          'EG20 V30.1.0.1 DEL START
'          MsgBox "���D�@�̃o�[�W�����쐬�ňُ킪�������܂����B", _
'                  vbOKOnly + vbExclamation, _
'                  "�������D�@ �o�[�W�����Ǘ�"
          'EG20 V30.1.0.1 DEL END
          'EG20 V30.1.0.1 ADD START
          MsgBox "���D�@�̃o�[�W�����쐬�ňُ킪�������܂����B", _
                  vbOKOnly + vbExclamation, _
                  "�V�����������D�@ �o�[�W�����Ǘ�"
          'EG20 V30.1.0.1 DEL END
        ' EG20 V30.1.0.1 DEL START
'        Else
'           MsgBox "���D�@�̃o�[�W�����쐬�ňُ킪�������܂����B", _
'                   vbOKOnly + vbExclamation, _
'                   "�������D�@ �o�[�W�����Ǘ�"
'        End If
        fOldVersion = False
    End If

    fOldVersion = True
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sCopyOLDtoNOW
'//  �@�\����  : ���o�[�W�����ɖ߂�����
'//  �@�\�T�v  : ���t�H���_���̃t�@�C�����A���s�t�H���_�ɃR�s�[����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-12  REVISED BY [TCC] S.Yoshimori
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(EG20 V3.5.0.1) 2012-02-07  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sCopyOLDtoNOW() As Boolean
    Dim i As Integer                '�J�E���^
    Dim sSrcFileName As String      '�R�s�[���t�@�C����
    Dim sDstFileName As String      '�R�s�[��t�@�C����
    Dim bError As Boolean           '�G���[�t���O
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.20.0.1 ADD

    Dim sCorner As String                      '�R�[�i�[�ԍ�
    Dim sGatePath As String                    '�R�[�i�[�ԍ��t�t�@�C���p�X
    
    On Error GoTo ErrorHandler
    
    '�����l�ݒ�
    sCopyOLDtoNOW = True

    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sGatePath = PATH_N_GATE & sCorner

    '****************************
    '* �t�@�C�����X�g���R�s�[���� *
    '****************************
'    sSrcFileName = FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
    sSrcFileName = sGatePath & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & MN_FILELIST
                                    '�u���v�t�H���_���t�@�C�������쐬����
'    sDstFileName = FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & MN_FILELIST
    sDstFileName = sGatePath & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & MN_FILELIST
                                    '�u���s�v�t�H���_���t�@�C�������쐬����
    'If Dir(sSrcFileName) <> "" Then     '�t�@�C���̌���������  'V1.20.0.1 DEL
    If objFso.FileExists(sSrcFileName) = True Then '�t�@�C���̌���������   'V1.20.0.1 ADD
        '�t�@�C�����X�g���u���v�t�H���_����u���s�v�t�H���_�ɃR�s�[����
        FileCopy sSrcFileName, sDstFileName
    Else
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
       
       '�u���t�H���_�t�@�C�����X�g�Ȃ��v�|�b�v�A�b�v��ʕ\��
        MsgBox "�u���v�t�H���_���� " & TitleBox(FolderSyubetu) & "�ɁA" _
                   & Chr(vbKeyReturn) & "�t�@�C�����X�g�����݂��܂���B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  �������s �R�s�["
        sCopyOLDtoNOW = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function                   '�������I������
    End If

    bError = False                  '�G���[�t���O���u�U�v�ɂ���
    For i = 0 To UBound(FileList) - 1
                                    '�t�@�C�����X�g�����J��Ԃ�
        '���t�H���_���t�@�C�������쐬����
'        sSrcFileName = FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & FileList(i)
        sSrcFileName = sGatePath & FolderName(MN_FOLD_OLD, FolderSyubetu) & "\" & FileList(i)

        '���s�t�H���_���t�@�C�������쐬����
'        sDstFileName = FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & FileList(i)
        sDstFileName = sGatePath & FolderName(MN_FOLD_NOW, FolderSyubetu) & "\" & FileList(i)

        '���t�H���_���̃t�@�C�������s�t�H���_�ɃR�s�[����
        'If Dir(sSrcFileName) <> "" Then     '�t�@�C���̌���������  'V1.20.0.1 DEL
        If objFso.FileExists(sSrcFileName) = True Then '�t�@�C���̌���������   'V1.20.0.1 ADD
            '�t�@�C�����u���v�t�H���_����u���s�v�t�H���_�ɃR�s�[����
            FileCopy sSrcFileName, sDstFileName
        Else                                '�t�@�C�������݂��Ȃ�
            bError = True                   '�G���[�t���O���u�^�v�ɂ���
        End If
    Next
    If bError = True Then
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        '�u���t�H���_�t�@�C�����X�g�o�^�Ȃ��v�|�b�v�A�b�v��ʕ\��
        MsgBox "�u���v�t�H���_���� " & TitleBox(FolderSyubetu) & "�ɁA" _
                   & Chr(vbKeyReturn) & "�t�@�C�����X�g�ɓo�^����Ă��āA���݂��Ȃ��t�@�C��������܂����B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  �������s �R�s�["
        sCopyOLDtoNOW = False
        Set objFso = Nothing    'V1.20.0.1 ADD
        Exit Function
    End If

    Set objFso = Nothing    'V1.20.0.1 ADD
    
' EG20 V3.0.0.2�ǉ��J�n
    If pfuncCopyPASSINF(iTab_index, MN_FOLD_OLD) = False Then
' EG20 V3.5.0.1�ǉ��J�n
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        MsgBox "�ُ�I�����܂����B", _
               vbOKOnly + vbExclamation, _
               TitleBox(FolderSyubetu) & "  �������s �R�s�["
' EG20 V3.5.0.1�ǉ��I��
        sCopyOLDtoNOW = False
    End If
' EG20 V3.0.0.2�ǉ��I��
    
    Exit Function       '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    '�u�������s�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
    MsgBox "�ُ�I�����܂����B", _
           vbOKOnly + vbExclamation, _
           TitleBox(FolderSyubetu) & "  �������s �R�s�["
        
    sCopyOLDtoNOW = False
    Set objFso = Nothing    'V1.20.0.1 ADD
End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fMakeOutPutFile
'//  �@�\����  : �}�̏o�͏������s���B
'//  �@�\�T�v  : �}�̏o�̓t�@�C���쐬�Əo�͂��s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 �t�@�C���N���[�Y�����ǉ�
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-17  REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-16  REVISED BY [TCC] M.Matsumoto
'//                 �y��-273�Ή��z
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fMakeOutPutFile()
   Dim sOutFileName As String '�}�̏o�̓t�@�C����[��ʕ�]
   Dim iFileNumber As Integer '�t�@�C���ԍ�
   Dim i As Integer           '�J�E���^�[
   Dim bFlag As Boolean       '�t���O
   Dim iResponse As Integer   'MsgBox�߂�l
   Dim lngErrCode As Long     '�G���[�R�[�h
   Dim fso         As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
   Dim strWriteDir As String               '�o�͐�t�H���_
   Dim strStationName As String
' EG20 V2.0.1.1 ADD START�y�c����60�z
   Dim iTab_index  As Integer
   Dim strSyubetu As String     ' ��ʖ�
' EG20 V2.0.1.1 ADD END�y�c����60�z
    
   On Error Resume Next 'V1.21.0.1 ADD

' EG20 V2.0.1.1 ADD START�y�c����60�z
    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
' EG20 V2.0.1.1 ADD END  �y�c����60�z

  '�t�H���_�I�𕔂Ɏw��L���`�F�b�N
  bFlag = False                                 '�t���O���u�U�v�ɂ���
  For i = 0 To 2                                '�t�H���_�����J��Ԃ�
     If chkFolder(i).Value = CHECKBOX_ON Then   '�u�H�H�v�t�H���_���w�肳��Ă���
        bFlag = True                            '�t���O���u�^�v�ɂ���
        Exit For                                '���[�v�𔲂���
     End If
  Next
              
  If bFlag = False Then                       '�t�H���_�w�薳��
     'If gStrCurrentForm = sFormName_EJVer Then     'EG20 V30.1.0.1 DEL
       '�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
       'EG20 V30.1.0.1 DEL START
'         MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
'                 vbOKOnly + vbExclamation, _
'                 "�������D�@ �o�[�W�����Ǘ�"
       'EG20 V30.1.0.1 DEL END
       'EG20 V30.1.0.1 ADD START
         MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
                 vbOKOnly + vbExclamation, _
                 "�V�����������D�@ �o�[�W�����Ǘ�"
        'EG20 V30.1.0.1 ADD END
     'EG20 V30.1.0.1 DEL START
'     Else
'       '�u�\���t�H���_�w��Ȃ��v�|�b�v�A�b�v�\��
'         MsgBox "�\��̫��ގw�肪�ЂƂ��I������Ă��܂���B", _
'                vbOKOnly + vbExclamation, _
'                "�������D�@ �o�[�W�����Ǘ�"
'     End If
         '�����𔲂���
     Exit Function
   End If
  
  
    'EG20 V2.1.0.1 ADD START �y��-273�Ή��z
    If lstKan(iTab_index).ListCount = 0 Then
        '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
        Exit Function
    End If
    'EG20 V2.1.0.1 ADD END
  
  '�t�H���_�I���|�b�v�A�b�v��ʕ\��
'  strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", "")                         'V1.12.0.1 DEL
  strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)    'V1.12.0.1 ADD

  '�w��t�H���_�Ȃ�
  If Len(strWriteDir) = 0 Then
       Exit Function
  End If

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
  '�v���O���X�o�[��\������
  Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

  '�R�s�[��t�H���_�̗L���m�F
  If fso.FolderExists(strWriteDir) = False Then
     '�R�s�[��t�H���_�쐬
     fso.CreateFolder (strWriteDir)
  End If
   
  '�w���擾
   strStationName = gsGetStationEkiName
  
   
   strSyubetu = ""
   '�������t�H�[���ɂ��A�}�̏o�͂���t�@�C�����쐬
'   If gStrCurrentForm = sFormName_EJVer Then
       '���\�[�X�I�𕔕���
       Select Case FolderSyubetu
        Case 0      '����CPU-Pro
'        ' EG20 V2.0.1.1 DEL START�y�c����60�z
'          sOutFileName = PATH_WORK & VER_TXT_EJHANTEIPRO
'          'strWriteDir = strWriteDir & VER_TXT_EJHANTEIPRO      'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJHANTEIPRO 'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   �y�c����60�z
        ' EG20 V2.0.1.1 ADD START�y�c����60�z
          'EG20 V30.1.0.1 DEL START
'          sOutFileName = PATH_WORK & strStationName & VER_TXT_EJHANTEIPRO
'          strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJHANTEIPRO 'V1.8.0.1 ADD
          'EG20 V30.1.0.1 DEL END
          'EG20 V30.1.0.1 ADD START
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_KJHANTEIPRO
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_KJHANTEIPRO
          'EG20 V30.1.0.1 ADD END
          strSyubetu = "����f�[�^"
        ' EG20 V2.0.1.1 ADD END  �y�c����60�z
        Case 1      '���C��CPU-Pro
'        ' EG20 V2.0.1.1 DEL START�y�c����60�z
'          sOutFileName = PATH_WORK & VER_TXT_EJMAINPRO
'          'strWriteDir = strWriteDir & VER_TXT_EJMAINPRO        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJMAINPRO   'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   �y�c����60�z
        ' EG20 V2.0.1.1 ADD START�y�c����60�z
          'EG20 V30.1.0.1 DEL START
'          sOutFileName = PATH_WORK & strStationName & VER_TXT_EJMAINPRO
'          strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJMAINPRO   'V1.8.0.1 ADD
          'EG20 V30.1.0.1 DEL END
          'EG20 V30.1.0.1 ADD START
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_KJMAINPRO
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_KJMAINPRO
          'EG20 V30.1.0.1 ADD END
          strSyubetu = "�v���O����"
        ' EG20 V2.0.1.1 ADD END  �y�c����60�z
        Case 2      '�T�uCPU-Pro1
'        ' EG20 V2.0.1.1 DEL START�y�c����60�z
'          sOutFileName = PATH_WORK & VER_TXT_EJSUBPRO
'          'strWriteDir = strWriteDir & VER_TXT_EJSUBPRO        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJSUBPRO   'V1.8.0.1 ADD
         ' EG20 V2.0.1.1 DEL END   �y�c����60�z
        ' EG20 V2.0.1.1 ADD START�y�c����60�z
          'EG20 V30.1.0.1 DEL START
'          sOutFileName = PATH_WORK & strStationName & VER_TXT_EJSUBPRO1
'          strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJSUBPRO1   'V1.8.0.1 ADD
'          strSyubetu = "�T�uCPU-Pro1"
          'EG20 V30.1.0.1 DEL END
          'EG20 V30.1.0.1 ADD START
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_KJSUBPRO
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_KJSUBPRO
          strSyubetu = "�T�uCPU-Pro"
          'EG20 V30.1.0.1 DEL END
        ' EG20 V2.0.1.1 ADD END  �y�c����60�z
        Case 3      '�T�uCPU-Pro2
'        ' EG20 V2.0.1.1 DEL START�y�c����60�z
'          sOutFileName = PATH_WORK & VER_TXT_EJMAINOS
'          'strWriteDir = strWriteDir & VER_TXT_EJMAINOS        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJMAINOS   'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   �y�c����60�z
        ' EG20 V2.0.1.1 ADD START�y�c����60�z
          'EG20 V30.1.0.1 DEL START
'          sOutFileName = PATH_WORK & strStationName & VER_TXT_EJSUBPRO2
'          strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJSUBPRO2   'V1.8.0.1 ADD
'          strSyubetu = "�T�uCPU-Pro2"
          'EG20 V30.1.0.1 DEL END
          'EG20 V30.1.0.1 ADD START
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_KJMAINOS
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_KJMAINOS
          strSyubetu = "�����i�n�r�j"
          'EG20 V30.1.0.1 ADD END
        ' EG20 V2.0.1.1 ADD END  �y�c����60�z
        Case 4      '�T�uCPU-Pro3
'        ' EG20 V2.0.1.1 DEL START�y�c����60�z
'          sOutFileName = PATH_WORK & VER_TXT_EJYOBI1
'          'strWriteDir = strWriteDir & VER_TXT_EJYOBI1         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJYOBI1    'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   �y�c����60�z
        ' EG20 V2.0.1.1 ADD START�y�c����60�z
          'EG20 V30.1.0.1 DEL START
'          sOutFileName = PATH_WORK & strStationName & VER_TXT_EJSUBPRO3
'          strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJSUBPRO3    'V1.8.0.1 ADD
'          strSyubetu = "�T�uCPU-Pro3"
          'EG20 V30.1.0.1 DEL END
          'EG20 V30.1.0.1 ADD START
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_KJYOBI1
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_KJYOBI1
          strSyubetu = "�\���P"
          'EG20 V30.1.0.1 ADD END
        ' EG20 V2.0.1.1 ADD END  �y�c����60�z
        Case 5      '���C��CPU-OS
'        ' EG20 V2.0.1.1 DEL START�y�c����60�z
'          sOutFileName = PATH_WORK & VER_TXT_EJMAINOS
'          'strWriteDir = strWriteDir & VER_TXT_EJMAINOS        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_EJMAINOS   'V1.8.0.1 ADD
        ' EG20 V2.0.1.1 DEL END   �y�c����60�z
        ' EG20 V2.0.1.1 ADD START�y�c����60�z
          'EG20 V30.1.0.1 DEL START
'          sOutFileName = PATH_WORK & strStationName & VER_TXT_EJMAINOS
'          strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJMAINOS   'V1.8.0.1 ADD
'          strSyubetu = "�����i�n�r�j"
          'EG20 V30.1.0.1 DEL END
          'EG20 V30.1.0.1 ADD START
          sOutFileName = PATH_WORK & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_KJYOBI
          strWriteDir = strWriteDir & "\" & strStationName & "_" & gstrCornerName(iTab_index) & VER_TXT_KJYOBI
          strSyubetu = "�\��"
          'EG20 V30.1.0.1 ADD END
        ' EG20 V2.0.1.1 ADD END  �y�c����60�z
        'EG20 V30.1.0.1 DEL START
'        Case 6      '�\��1
''        ' EG20 V2.0.1.1 DEL START�y�c����60�z
''          sOutFileName = PATH_WORK & VER_TXT_EJYOBI1
''          'strWriteDir = strWriteDir & VER_TXT_EJYOBI1         'V1.8.0.1 DEL
''          strWriteDir = strWriteDir & "\" & VER_TXT_EJYOBI1    'V1.8.0.1 ADD
'        ' EG20 V2.0.1.1 DEL END   �y�c����60�z
'        ' EG20 V2.0.1.1 ADD START�y�c����60�z
'          sOutFileName = PATH_WORK & strStationName & VER_TXT_EJYOBI1
'          strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJYOBI1    'V1.8.0.1 ADD
'          strSyubetu = "�\���P"
'        ' EG20 V2.0.1.1 ADD END  �y�c����60�z
'        Case 7      '�\��2
''        ' EG20 V2.0.1.1 DEL START�y�c����60�z
''          sOutFileName = PATH_WORK & VER_TXT_EJYOBI2
''          'strWriteDir = strWriteDir & VER_TXT_EJYOBI2         'V1.8.0.1 DEL
''          strWriteDir = strWriteDir & "\" & VER_TXT_EJYOBI2    'V1.8.0.1 ADD
'        ' EG20 V2.0.1.1 DEL END   �y�c����60�z
'        ' EG20 V2.0.1.1 ADD START�y�c����60�z
'          sOutFileName = PATH_WORK & strStationName & VER_TXT_EJYOBI2
'          strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJYOBI2    'V1.8.0.1 ADD
'          strSyubetu = "�\���Q"
'        ' EG20 V2.0.1.1 ADD END  �y�c����60�z
'        Case 8      '�\��3
''        ' EG20 V2.0.1.1 DEL START�y�c����60�z
''          sOutFileName = PATH_WORK & VER_TXT_EJYOBI3
''          'strWriteDir = strWriteDir & VER_TXT_EJYOBI2         'V1.8.0.1 DEL
''          strWriteDir = strWriteDir & "\" & VER_TXT_EJYOBI3    'V1.8.0.1 ADD
'        ' EG20 V2.0.1.1 DEL END   �y�c����60�z
'        ' EG20 V2.0.1.1 ADD START�y�c����60�z
'          sOutFileName = PATH_WORK & strStationName & VER_TXT_EJYOBI3
'          strWriteDir = strWriteDir & "\" & strStationName & VER_TXT_EJYOBI3    'V1.8.0.1 ADD
'          strSyubetu = "�\���R"
'        ' EG20 V2.0.1.1 ADD END  �y�c����60�z
         'EG20 V30.1.0.1 DEL END
        End Select
'  Else
'       '���\�[�X�I�𕔕���
'       Select Case FolderSyubetu
'        Case 0      '����CPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_NJHANTEIPRO
'          'strWriteDir = strWriteDir & VER_TXT_NJHANTEIPRO      'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJHANTEIPRO 'V1.8.0.1 ADD
'        Case 1      '���C��CPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_NJMAINPRO
'          'strWriteDir = strWriteDir & VER_TXT_NJMAINPRO        'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJMAINPRO   'V1.8.0.1 ADD
'        Case 2      '�T�uCPU-Pro
'          sOutFileName = PATH_WORK & VER_TXT_NJSUBPRO
'          'strWriteDir = strWriteDir & VER_TXT_NJSUBPRO         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJSUBPRO    'V1.8.0.1 ADD
'        Case 3      '���C��CPU-OS
'          sOutFileName = PATH_WORK & VER_TXT_NJMAINOS
'          'strWriteDir = strWriteDir & VER_TXT_NJMAINOS         'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJMAINOS    'V1.8.0.1 ADD
'        Case 4      '�\��1
'          sOutFileName = PATH_WORK & VER_TXT_NJYOBI1
'          'strWriteDir = strWriteDir & VER_TXT_NJYOBI1          'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJYOBI1     'V1.8.0.1 ADD
'        Case 5      '�\��2
'          sOutFileName = PATH_WORK & VER_TXT_NJYOBI2
'          'strWriteDir = strWriteDir & VER_TXT_NJYOBI2          'V1.8.0.1 DEL
'          strWriteDir = strWriteDir & "\" & VER_TXT_NJYOBI2     'V1.8.0.1 ADD
'        End Select
'  End If

  iFileNumber = FreeFile              '���g�p�̃t�@�C���ԍ����擾����
 
  '�Ώۃt�@�C�����I�[�v������B
  Open sOutFileName For Output Access Write As #iFileNumber
  
  ' �ݒu�w����������
   Print #iFileNumber, "�ݒu�w�F" & strStationName
   Print #iFileNumber, ""
     
  ' �f�[�^��ʁi���[�N�j��������
   Print #iFileNumber, "�f�[�^��ʁF" & strSyubetu
   Print #iFileNumber, ""

  ' �S�̃o�[�W������������
   Print #iFileNumber, "�S�̃o�[�W�����i���[�N�j�F" & DispTitleVersion(MN_FOLD_WRK)
   Print #iFileNumber, "�@�@�@�@�@�@�@�i���s�j�@�F" & DispTitleVersion(MN_FOLD_NOW)
   Print #iFileNumber, "�@�@�@�@�@�@�@�i���j�@�@�F" & DispTitleVersion(MN_FOLD_OLD)
   Print #iFileNumber, ""

'  For i = 0 To lstKan(0).ListCount - 1
  For i = 0 To lstKan(iTab_index).ListCount - 1
  '���X�g�{�b�N�X�ɕ\������Ă��镪�����A�������ށB
'       Print #iFileNumber, lstKan(0).List(i) & Chr(vbKeyReturn)
'       Print #iFileNumber, lstKan(iTab_index).List(i) & Chr(vbKeyReturn)   ' EG20 V3.0.0.2�폜
       Print #iFileNumber, lstKan(iTab_index).List(i)                       ' EG20 V3.0.0.2�ǉ�
  Next
 
  '�Ώۃt�@�C�����N���[�Y����B
  Close #iFileNumber

  '�t�@�C���̗L���m�F
  If fso.FileExists(sOutFileName) = False Then
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
     '�v���O���X�o�[����������
     Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
     '�t�@�C�������ُ�|�b�v�A�b�v��ʕ\��
     MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
     Exit Function
  End If
    
  On Error GoTo COPY_ERROR
  '�t�@�C���R�s�[
  fso.CopyFile sOutFileName, strWriteDir
  '�u�}�̏o�͐���I���v�|�b�v�A�b�v��ʕ\��
  'V1.8.0.1 DEL START
  'iResponse = MsgBox("����I�����܂����B", vbOKOnly, _
  '                   "�o�͌���")
  'V1.8.0.1 DEL END
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
  '�v���O���X�o�[����������
  Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
  
  MsgBox "����I�����܂����B", vbInformation, "�o�͌���"   'V1.8.0.1 ADD
                   
  '�u�����ް�ޮ݁F�}�̏o�͏�������v���O�o��
  Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_OUTPUT_OK, 0)
  
  Set fso = Nothing

  Exit Function
    
'*******************************
'VB�G���[����
COPY_ERROR:
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        '�����ُ�̏ꍇ�A�o�͌��ʃ|�b�v�A�b�v(�ُ�)�\��
        MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
        '�u�����ް�ޮ݁F�}�̏o�͏����ُ�v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_OUTPUT_ERROR, lngErrCode)
        Set fso = Nothing
'*******************************
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-16   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�Q�s��C��
'//                 �t�F�[�Y�R�Ή�
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 ���̓t�@�C���i�[�f�B���N�g���ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//                 �t�@�C���I����ʂ�OS�d�l�ɕύX
'//                 Dir�֐���FileSystemObject�ɒu������
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                  EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V5.5.0.1) 2012-03-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V5.11.0.1) 2012-05-10  CODED BY  [TCC] H.Sugimoto
'//                �y�^���\�����P�Ή��z
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20 V6.9.0.1) 2012-12-05  CODED BY  [TCC] H.Sugimoto
'//                 �ʎY�Ή��yTOMAS�p�̈�R�s�[�Ή��z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(EG20 V30.3.0.1) 2014-10-16  CODED BY  [TCC] T.Nakajima
'//                  �k���V�����t�F�[�Y�Q�Ή� �yHKRK_Kansi06_004_02�z
'//     REVISIONS :(EG20 V30.4.0.1) 2015-01-15 CODED BY  [TCC] S.Kuroda
'//                 �k���V�����t�F�[�Y�R�Ή��yHKRK_kansi02_001_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sFDInstall(sFlag As String)
    Dim MyName As String            '�t�@�C���t���p�X��
    Dim sSrcFileName As String      '�R�s�[���t�@�C����
    Dim sDstFileName As String      '�R�s�[��t�@�C����
    Dim iResponse As Integer        'MsgBox�{�^���R�[�h
    Dim sInputPass As String        '�C���X�g�[�����f�B���N�g����(STD)or�t�@�C����(LZH)
    Dim sInputFolder As String      '�C���X�g�[�����t�H���_���BLZH�̎��A�𓀐�t�H���_�B
    Dim lngErrCode As Long          '�G���[�R�[�h
    'V1.6.0.1 ADD START
    Dim bRet As Boolean             '�������`�F�b�N�߂�l
    Dim sChkName As String          '�`�F�b�N�t�@�C��
    'V1.6.0.1 ADD END
    
    'V1.20.0.1 ADD START
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim objFi As File                    '�t�@�C���I�u�W�F�N�g
    'V1.20.0.1 ADD END
    
    Dim sCorner As String            '�R�[�i�[�ԍ�
    Dim sGatePath As String          '�R�[�i�[�ԍ��t�t�@�C���p�X
    Dim sFilePath As String          '�t�@�C���t�@�C���p�X
    Dim lngPgmHanteiStsWork As Long     '�v���O���������ԁi���[�N�j   ' EG20 V3.0.0.2�ǉ�
    Dim szTargetFolder As String     ' �����ύX��t�H���_��             ' EG20 V5.8.0.1�ǉ�
    
    Dim sTomasPath As String         ' TOMAS�p�̈�t�@�C���p�X
    
    On Error GoTo ErrorHandler      '�G���[�n���h���̓o�^

    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sGatePath = PATH_N_GATE & sCorner

' EG20 V5.8.0.1�ǉ��J�n
    szTargetFolder = sGatePath & FolderName(0, FolderSyubetu)
' EG20 V5.8.0.1�ǉ��I��

    If sFlag = "STD" Then
    '�W���i�񈳏k�j�t�@�C���w��̎�:
    '�f�B���N�g���I����ʂ�\�������A���̓t�@�C���i�[�f�B���N�g�����𓾂�B
'       sInputPass = pfDirSelection("a:", "�C���X�g�[���}�̂̃f�B���N�g���I��")     'V1.12.0.1 DEL
        'sInputPass = pfDirSelection("H:", "�C���X�g�[���}�̂̃f�B���N�g���I��")      'V1.12.0.1 ADD    'V1.20.0.1 DEL
        sInputPass = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
        If sInputPass = "" Then
        '�f�B���N�g�����w��Ȃ����͏����I��
            'V1.20.0.1 ADD START
            Set objFso = Nothing
            Set objFi = Nothing
            'V1.20.0.1 ADD END
            Exit Sub
        End If
        sInputFolder = sInputPass
    Else
    '���k�t�@�C���w��̎�:
    '���k�t�@�C���I����ʂ�\�������ALZH�t�@�C���t���p�X���𓾂�i�f�t�H���g�͂e�c��\���B�j�B
'       sInputPass = pfCabFileSelection("a:")     'V1.12.0.1 DEL
        'V1.20.0.1 DEL START
       'sInputPass = pfCabFileSelection("H:")      'V1.12.0.1 ADD
        'If sInputPass = "" Then Exit Sub '�t�@�C�����I������Ȃ���Ζ߂�B
        'V1.20.0.1 DEL END
        'V1.20.0.1 ADD START
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
        CommonDialog1.Filter = "���k�t�@�C���i*.cab�j|*.cab|"
        '�t�@�C���I����ʂ��J��
        CommonDialog1.ShowOpen
        '�I�������t�@�C�������擾
        sInputPass = CommonDialog1.FileName
        If sInputPass = "" Then '�t�@�C�����I��
            Set objFso = Nothing
            Set objFi = Nothing
            Exit Sub    '�t�@�C�����I������Ȃ���Ώ������f
        End If
        
        Call ChDrive("D")  'V2.5.0.1 ADD
        
        'V1.20.0.1 ADD END
       '�𓀗p�ꎞ�t�H���_���쐬����B
       psMakeFolder MELTED_FOLDER_FULLPASS
       '���k�t�@�C�����A�𓀗p�ꎞ�t�H���_�ɉ𓀁E�i�[������B
        Call psCabReqest(CABREQEST.CAB_THAW, sInputPass, MELTED_FOLDER_FULLPASS)
        If glngCabErrCd <> 0 Then
            psDeleteFolder MELTED_FOLDER_FULLPASS
            'V1.20.0.1 ADD START
            Set objFso = Nothing
            Set objFi = Nothing
            'V1.20.0.1 ADD END
            Exit Sub
        End If
        sInputFolder = MELTED_FOLDER_FULLPASS
    End If
    
    '�u���[�N�R�s�[�m�F�v�|�b�v�A�b�v��ʕ\��
    iResponse = MsgBox(sInputPass & " �̑S�Ẵt�@�C�����A" _
                       & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                       & TitleBox(FolderSyubetu) & "�́u���[�N�v�t�H���_�ɃR�s�[���܂��B " _
                       & "��낵���ł����H", _
                       vbYesNo + vbExclamation, _
                       TitleBox(FolderSyubetu) & "  �}�́����[�N �R�s�[")
    If iResponse = vbNo Then
    '[������] �{�^����I��:�������Ȃ��B
    '�A���A���k�t�@�C���w��̎��́A�𓀗p�ꎞ�t�H���_���폜����B
       If sFlag = "LZH" Then
           psDeleteFolder MELTED_FOLDER_FULLPASS
       End If
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
    
    'V1.6.0.1 ADD START
    '�O�����̓v�����������`�F�b�N
    If sFlag = "STD" Then
       '�}�́����[�N �R�s�[��
       bRet = pfInstallSeitouseiChck(sInputPass)
    Else
       '���k�t�@�C�������[�N �R�s�[��
       bRet = pfInstallSeitouseiChck(MELTED_FOLDER_FULLPASS & "\")
    End If
    If bRet = False Then
        Call pubfuncErrorOccur(MN_FOLD_WRK)         ' EG20 V3.0.0.2�ǉ�
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
       If sFlag = "LZH" Then
           psDeleteFolder MELTED_FOLDER_FULLPASS
       End If
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End
       'V1.20.0.1 ADD START
       Set objFso = Nothing
       Set objFi = Nothing
       'V1.20.0.1 ADD END
       Exit Sub
    End If
    
    '�o�[�W�����`�F�b�N�t�@�C���L���`�F�b�N���s���B
    sChkName = fSelectFile
    'V1.20.0.1 DEL START
'    sChkName = Dir(FolderName(0, FolderSyubetu) & "\" & sChkName)
'    If sChkName <> "" Then
'      Kill FolderName(0, FolderSyubetu) & "\" & sChkName
'    End If
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    
    sFilePath = sGatePath & FolderName(0, FolderSyubetu)
    
'    If objFso.FileExists(FolderName(0, FolderSyubetu) & "\" & sChkName) = True Then
    If objFso.FileExists(sFilePath & "\" & sChkName) = True Then
        '�w��t�@�C�������݂���
'        sChkName = objFso.GetFileName(FolderName(0, FolderSyubetu) & "\" & sChkName)
        sChkName = objFso.GetFileName(sFilePath & "\" & sChkName)
'        Kill FolderName(0, FolderSyubetu) & "\" & sChkName
        Kill sFilePath & "\" & sChkName
    Else
        sChkName = ""
    End If
    'V1.20.0.1 ADD END
    'V1.6.0.1 ADD START
    
    '�w��t�H���_���̃t�@�C�����A�S�āu���[�N�v�t�H���_�ɃR�s�[����B
    'V1.20.0.1 DEL START
'    MyName = Dir(sInputFolder & "\*.*", vbNormal)  ' �ŏ��̃f�B���N�g������Ԃ��܂��B
'    Do While MyName <> ""                   ' ���[�v���J�n���܂��B
'        ' ���݂̃f�B���N�g���Ɛe�f�B���N�g���͖������܂��B
'        If MyName <> "." And MyName <> ".." Then
'            '�}�̓��t�@�C�������쐬����
'            sSrcFileName = sInputFolder & "\" & MyName
'            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
'            If (GetAttr(sSrcFileName) And vbDirectory) <> vbDirectory Then
'                '���[�N�t�H���_���t�@�C�������쐬����
'                sDstFileName = FolderName(0, FolderSyubetu) & "\" & MyName
'                '�}�̓��̃t�@�C�������[�N�t�H���_�ɃR�s�[����
'                FileCopy sSrcFileName, sDstFileName
'            End If
'        End If
'        MyName = Dir                    ' ���̃f�B���N�g������Ԃ��܂��B
'    Loop
    'V1.20.0.1 DEL END
    'V1.20.0.1 ADD START
    For Each objFi In objFso.GetFolder(sInputFolder).files   '���[�v���J�n
        If objFso.FileExists(objFi.Path) = True Then  '�t�@�C�����̎擾�`�F�b�N
            '�f�B���N�g�������擾
            MyName = objFi.Name
            '�}�̓��t�@�C�������쐬
            sSrcFileName = sInputFolder & "\" & MyName
            ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
            If (GetAttr(sSrcFileName) And vbDirectory) <> vbDirectory Then
                '���[�N�t�H���_���t�@�C�������쐬����
'                sDstFileName = FolderName(0, FolderSyubetu) & "\" & MyName
                sDstFileName = sGatePath & FolderName(0, FolderSyubetu) & "\" & MyName

                '�}�̓��̃t�@�C�������[�N�t�H���_�ɃR�s�[����
                FileCopy sSrcFileName, sDstFileName
            End If
        End If
    Next
    
    Set objFso = Nothing
    Set objFi = Nothing
    'V1.20.0.1 ADD END
    
    '���k�t�@�C���w��̎��́A�𓀗p�ꎞ�t�H���_���폜����B(�g�p�ς݂̂���)
    If sFlag = "LZH" Then
        psDeleteFolder MELTED_FOLDER_FULLPASS
    End If
    
' EG20 V5.8.0.1�폜�J�n
'    ' �^����ԍX�V                                              ' EG20 V5.5.0.1�ǉ�
'    Call pubFuncUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI)           ' EG20 V5.5.0.1�ǉ�
' EG20 V5.8.0.1�폜�I��
' EG20 V5.8.0.1�ǉ��J�n
    '�ǂݎ��O���̊֐������s
    dllChangeAttributeContents (szTargetFolder)

' EG20 V6.9.0.1 �y�ʎY�Ή��FTOMAS�p�̈�R�s�[�Ή��zADD START
    ' �������ׂ��Ώۂ��R�[�i1�̏ꍇ
    ' TOMAS�̈�iN_GATE00�j��N_GATE01�̓��e�ŃR�s�[
    'If iTab_index = 0 Then     'EG20 V30.1.0.1 DEL
    'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL START
    '���[�N�R�s�[���悤�Ƃ��邽�тɂ��̃R�[�i����00�փR�s�[���邽�߁A�擪�R�[�i�̔�����폜
    'If iTab_index = gintKansenFirstCornerIdx Then  'EG20 V30.1.0.1 ADD
    'EG20 V30.3.0.1 �yHKRK_Kansi06_004_02�z DEL END
        ' �폜��̃t�H���_�iTOMAS�̈�j���w��
        sTomasPath = PATH_N_GATE & "00" & FolderName(0, FolderSyubetu) & "\"
        sInputFolder = sGatePath & FolderName(0, FolderSyubetu) & "\"
        
        ' TOMAS�̈���폜
        If funcRemoveFile(sTomasPath) = False Then
            
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            
            'EG20 V30.1.0.1 DEL START
'            MsgBox "�s�n�l�`�r�p�̈�R�s�[�ُ�I��", _
'                    vbOKOnly + vbExclamation, _
'                    "�������D�@�@�o�[�W�����Ǘ�"
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.1.0.1 ADD START
            MsgBox "�s�n�l�`�r�p�̈�R�s�[�ُ�I��", _
                    vbOKOnly + vbExclamation, _
                    "�V�����������D�@�@�o�[�W�����Ǘ�"
            'EG20 V30.1.0.1 ADD END
            
            '�u�����ް�ޮ݁FTOMAS̫���̧�ٍ폜�ُ�v���O�o��
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_TOMASFILE_DELETE_ERROR, lngErrCode)
        
            GoTo TomasErrorHandler
        End If
        
        ' TOMAS�̈�փR�s�[
        If funcCopyFile(sInputFolder, sTomasPath, lngErrCode) = False Then
            
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
            
            'EG20 V30.1.0.1 DEL START
'            MsgBox "�s�n�l�`�r�p�̈�R�s�[�ُ�I��", _
'                    vbOKOnly + vbExclamation, _
'                    "�������D�@�@�o�[�W�����Ǘ�"
            'EG20 V30.1.0.1 DEL END
            'EG20 V30.1.0.1 ADD START
            MsgBox "�s�n�l�`�r�p�̈�R�s�[�ُ�I��", _
                    vbOKOnly + vbExclamation, _
                    "�V�����������D�@�@�o�[�W�����Ǘ�"
            'EG20 V30.1.0.1 ADD END
            
            '�u�����ް�ޮ݁FTOMAS�̈��߰�����ُ�v���O�o��
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_TOMASFILE_COPY_ERROR, lngErrCode)
        
            GoTo TomasErrorHandler
        End If
    'EG20 V30.3.0.1�yHKRK_Kansi06_004_02�z DEL START
    'End If
    'EG20 V30.3.0.1�yHKRK_Kansi06_004_02�z DEL END
' EG20 V6.9.0.1 �y�ʎY�Ή��FTOMAS�p�̈�R�s�[�Ή��zADD END

    ' �^����ԍX�V
    Call pubFuncUpdateUnkaiStatus(BootInfoHoshuType.TYPE_GATE, BOOTINFO_UNKAI_ARI)
' EG20 V5.8.0.1�ǉ��I��
'    Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI, iTab_index + 1)   ' EG20 V5.6.0.1�ǉ�           ' EG20 V5.11.0.1�폜
    Call pubFuncGateUpdateUnkaiStatus(BOOTINFO_UNKAI_ARI, iTab_index + 1, gintUnkaiKind(FolderSyubetu))    ' EG20 V5.11.0.1�ǉ�
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '�u���[�N�R�s�[����I���v�|�b�v�A�b�v��ʕ\��
    MsgBox "�C���X�g�[���}�̂̑S�Ẵt�@�C�����A" _
            & Chr(vbKeyReturn) & TitleBox(FolderSyubetu) & "�́u���[�N�v�t�H���_��" _
            & Chr(vbKeyReturn) & "�R�s�[���܂����B", _
            vbOKOnly + vbExclamation, _
            TitleBox(FolderSyubetu) & "  �}�́����[�N �R�s�["
    
    '�u�����ް�ޮ݁F�}�́�ܰ���߰��������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_OK, 0)
    
    '���X�g�{�b�N�X������������
    lstKan(0).Clear
    lstKan(1).Clear
    lstKan(2).Clear
    lstKan(3).Clear
    lstKan(4).Clear
    lstKan(5).Clear
  
    '�o�[�W������񃊃X�g�{�b�N�X���쐬����
    fMakeListbox
    
    '�Ď��ݒ�G���A�u�v���O��������ُ��ԁi���[�N�j�v�̏�Ԃ��擾����
    lngPgmHanteiStsWork = pfGetKansiSet(IdKansiSet.PG_HANTEI_ERR_STS_WORK)

    '�u�v���O��������ُ��ԁi���[�N�j�v�i����j
    Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_ERR_STS_WORK, ErrCode.Normal)
    
    '�ω����������ꍇ�A�u��ԕω��ʒm�v�𑗐M����
    If lngPgmHanteiStsWork <> ErrCode.Normal Then
        Call sSendMailStsChgInf(MailSts.stsNormal, lngPgmHanteiStsWork)
    End If
    
    Exit Sub    '�������I������

ErrorHandler:   ' �G���[�����B
    'V1.20.0.1 ADD START
    Set objFso = Nothing
    Set objFi = Nothing
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    'V1.20.0.1 ADD END
    Select Case Err.Number
        Case 53 ' �u�w���ʃt�@�C���Ȃ��v�|�b�v�A�b�v��ʕ\��
            MsgBox "�C���X�g�[���}�̂� " & TitleBox(FolderSyubetu) & "�́A" _
                   & Chr(vbKeyReturn) & "�ЂƂ����݂��܂���B", _
                   vbOKOnly + vbExclamation, _
                   TitleBox(FolderSyubetu) & "  �����[�N �R�s�["
            Exit Sub
        Case 71 '�u�}�̂Ȃ��v�|�b�v�A�b�v��ʕ\��
            iResponse = MsgBox("�}�̂���������Ă��܂���B", _
                    vbRetryCancel + vbExclamation, _
                    TitleBox(FolderSyubetu) & "  �����[�N �R�s�[")
            If iResponse = vbRetry Then    '�u��蒼���v�{�^����I�������ꍇ
                Resume      ' �G���[�����������s���珈���ĊJ
            Else                            '�u�L�����Z���v�{�^����I�������ꍇ
                Exit Sub    '�������I������
            End If
        Case Else  '�u���[�N�R�s�[�ُ�I���v�|�b�v�A�b�v��ʕ\��
           MsgBox "�C���X�g�[���}�̂���̃R�s�[�G���[���������܂����B" _
                   & Chr(vbKeyReturn) & "�G���[�R�[�h��" _
                   & str$(Err.Number), _
                   vbOKOnly + vbExclamation, _
                   "�����[�N �R�s�["
    End Select
    
    Call pubfuncErrorOccur(MN_FOLD_WRK)         ' EG20 V3.0.0.2�ǉ�

' EG20 V5.8.0.1�ǉ��J�n
    dllChangeAttributeContents (szTargetFolder)
' EG20 V5.8.0.1�ǉ��I��
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD Start
    '���k�t�@�C���w��̎��́A�𓀗p�ꎞ�t�H���_���폜����B
    If sFlag = "LZH" Then
        psDeleteFolder MELTED_FOLDER_FULLPASS
    End If
' EG20 V30.4.0.1�yHKRK_kansi02_001_01�z ADD End

    '�u�����ް�ޮ݁F�}�́�ܰ���߰�����ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_LZH_COPY_WRK_ERROR, lngErrCode)

' EG20 V6.9.0.1 �y�ʎY�Ή��FTOMAS�p�̈�R�s�[�Ή��zADD START
    Exit Sub    '�������I������

TomasErrorHandler:   ' TOMAS�����p�G���[�����B
' EG20 V6.9.0.1 �y�ʎY�Ή��FTOMAS�p�̈�R�s�[�Ή��zADD END
    
    Call pubfuncErrorOccur(MN_FOLD_WRK)
    
    '���X�g�{�b�N�X������������
    lstKan(0).Clear
    lstKan(1).Clear
    lstKan(2).Clear
    lstKan(3).Clear
    lstKan(4).Clear
    lstKan(5).Clear
  
    '�o�[�W������񃊃X�g�{�b�N�X���쐬����
    fMakeListbox

End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : SSTab1_Click
'//  �@�\����  : �R�[�i�^�u�I������
'//  �@�\�T�v  : �R�[�i�\����؂�ւ���
'//
'//              �^        ����             �Ӗ�
'//  ����      : Integer   PreviousTab      �I���^�u
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'/  REVISIONS    : (EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'/                  EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub SSTab1_Click(PreviousTab As Integer)
    
    On Error GoTo ErrorHandle
    
    '���X�g�{�b�N�X������������
    lstKan(0).Clear
    lstKan(1).Clear
    lstKan(2).Clear
    lstKan(3).Clear
    lstKan(4).Clear
    lstKan(5).Clear
    
    '�o�[�W������񃊃X�g�{�b�N�X���쐬����
    fMakeListbox
ErrorHandle:

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : pubfuncCommonGateCheck
'//  �@�\����  : ���D�@���ʔ��菈��
'//  �@�\�T�v  : �T���l�`�F�b�N�A�t�@�C�����ő�`�F�b�N�̎��s
'//
'//              �^         ����            �Ӗ�
'//  ����      : Integer    nKind           MN_FOLD_WRK(0):���[�N
'//                                         MN_FOLD_NOW(1):���s
'//                                         MN_FOLD_OLD(2):��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : BOOL      TRUE      ����
'//                        FALSE     �ُ�
'//
'//  ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                  EG20�t�F�[�Y�Q�Ή�
'//  REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//     REVISIONS :(EG20 V30.1.0.1) 2014-02-17  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pubfuncCommonGateCheck(nKind As Integer) As Boolean

    Dim lngSumRet As Long
    Dim lngCnt As Long
    Dim lngFileListCnt As Long               '�t�@�C�����X�g��
    Dim i As Integer
    Dim strWork     As String                '��ƃG���A
    Dim iFileNumber As Integer               '���g�p�t�@�C���ԍ�
    Dim bRet As Boolean
    Dim sGetFileListName As String           '�t�@�C�����X�g���L�ڃt�@�C����
    Dim myLen As Long                        '������̒���
    Dim sCorner As String                    '�R�[�i�[�ԍ�
    Dim sGatePath As String                  '�R�[�i�[�ԍ��t�t�@�C���p�X
    Dim sFilePath As String                  '�t�@�C���t�@�C���p�X
    Dim lTotalCount As Long                  ' ���ʌ���

    Dim lngPgmHanteiRcvErrSts   As Long     '�v���O���������M�ُ���
    Dim lngPgmHanteiSndErrSts   As Long     '�v���O��������z�M�ُ���
    Dim lngPgmHanteiErrSts      As Long     '�v���O��������ُ��ԁi���s�j
    Dim lngPgmHanteiErrStsOld   As Long     '�v���O��������ُ��ԁi���j
    Dim lngPgmHanteiElseErrSts  As Long     '�v���O�������肻�̑��ُ���

    
    On Error Resume Next

    ' �I�𒆂̃R�[�i�[�ԍ��擾
    iTab_index = SSTab1.Tab
    
    sCorner = Format(iTab_index + 1, "00")
    
    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sGatePath = PATH_N_GATE & sCorner


    ' /////////////////////////////////////////////////////
    ' // �T���l�`�F�b�N
    For lngCnt = 0 To UBound(FileList) - 1
        sFilePath = sGatePath & FolderName(nKind, FolderSyubetu)
        If pfFileSumChk(sFilePath & "\" & FileList(lngCnt), lngSumRet) <> True Then
            
            '�u�v���O���������M�ُ��ԁv�擾
            lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)
        
            '�Ď��ݒ�G���A�u�v���O���������M�ُ��ԁv���X�V
            Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_SumChk)
                    
            '�ă}�v���Z�X�Ɂu��ԕω��ʒm�v�𑗐M
            If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_SumChk Then
                Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_SumChk)
            End If
            
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
            '�T���l�ُ�
            If lngSumRet = SUM_CHK.SumErr Then
               'EG20 V30.1.0.1 DEL START
'               MsgBox "�T���l���ُ�ł��B" _
'                      & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
'                      vbOKOnly + vbExclamation, _
'                      "�������D�@ �o�[�W�����Ǘ�"
               'EG20 V30.1.0.1 DEL END
               'EG20 V30.1.0.1 ADD START
               MsgBox "�T���l���ُ�ł��B" _
                      & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
                      vbOKOnly + vbExclamation, _
                      "�V�����������D�@ �o�[�W�����Ǘ�"
               'EG20 V30.1.0.1 ADD END
            
            '�T���l�ُ�ȊO�ُ�
            ElseIf lngSumRet = SUM_CHK.SumErr_Else Then
               'EG20 V30.1.0.1 DEL START
'               MsgBox "�ُ�I�����܂����B", _
'                     vbOKOnly + vbExclamation, _
'                      "�������D�@ �o�[�W�����Ǘ�"
               'EG20 V30.1.0.1 DEL END
               'EG20 V30.1.0.1 ADD START
               MsgBox "�ُ�I�����܂����B", _
                     vbOKOnly + vbExclamation, _
                      "�V�����������D�@ �o�[�W�����Ǘ�"
               'EG20 V30.1.0.1 ADD END
            End If
            pubfuncCommonGateCheck = False
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_SUMCHK_ERROR, 0)
            Exit Function
        End If
    Next

    ' /////////////////////////////////////////////////////
    ' // �t�@�C�����ő�`�F�b�N
    If UBound(FileList) > FILECNT_MAX Then

        '�u�v���O���������M�ُ��ԁv
        lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)

        '�Ď��ݒ�G���A�u�v���O���������M�ُ��ԁv���X�V
        Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_FileMaxChk)
                
        '�ă}�v���Z�X�Ɂu��ԕω��ʒm�v�𑗐M
        If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_FileMaxChk Then
            Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_FileMaxChk)
        End If
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

        'EG20 V30.1.0.1 DEL START
'        MsgBox "�t�@�C����������𒴂��Ă��܂��B" _
'                & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
'                vbOKOnly + vbExclamation, _
'                "�������D�@ �o�[�W�����Ǘ�"
        'EG20 V30.1.0.1 ADD START
        MsgBox "�t�@�C����������𒴂��Ă��܂��B" _
                & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
                vbOKOnly + vbExclamation, _
                "�V�����������D�@ �o�[�W�����Ǘ�"
        'EG20 V30.1.0.1 ADD END
        pubfuncCommonGateCheck = False

        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)
        Exit Function
    End If

    'EG20 V30.1.0.1 DEL START �k���V�����ł͑S��ʂ̏���l�������Ă��Ȃ��̂Ń`�F�b�N�͕s�v�Ƃ���
'    ' /////////////////////////////////////////////////////
'    ' // �S�t�@�C�����ő�`�F�b�N�i���s�{�ǉ����j
'    bRet = True
'    lTotalCount = pfuncTotalListCount()
'    lTotalCount = lTotalCount + UBound(FileList)
'    If lTotalCount > TOTALFILECNT_MAX Then
'        bRet = False
'    End If
'    If bRet = False Then
'        '�u�v���O���������M�ُ��ԁv
'        lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)
'
'        '�Ď��ݒ�G���A�u�v���O���������M�ُ��ԁv���X�V
'        Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_FileMaxChk)
'
'        '�ă}�v���Z�X�Ɂu��ԕω��ʒm�v�𑗐M
'        If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_FileMaxChk Then
'            Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_FileMaxChk)
'        End If
'
'' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
'        '�v���O���X�o�[����������
'        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
'' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
'        MsgBox "�t�@�C����������𒴂��Ă��܂��B" _
'                & Chr(vbKeyReturn) & "�f�[�^���m�F���Ă��������B", _
'                vbOKOnly + vbExclamation, _
'                "�������D�@ �o�[�W�����Ǘ�"
'        pubfuncCommonGateCheck = False
'
'        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_MAXFILECHK_ERROR, 0)
'        Exit Function
'    End If
    'EG20 V30.1.0.1 DEL END

    pubfuncCommonGateCheck = True
    Exit Function

' �����{
'    ' /////////////////////////////////////////////////////
'    ' // �t�@�C�����T�C�Y�`�F�b�N
'    lngFileListCnt = UBound(FileList)
'
'    On Error GoTo FileGetError
'
'    iFileNumber = FreeFile          '���g�p�̃t�@�C���ԍ����擾����
'
'    sFilePath = sGatePath & FolderName(nKind, FolderSyubetu)
'    '�t�@�C�����X�g���I�[�v���B
'    Open sFilePath & "\" & MN_FILELIST For Input As #iFileNumber
'
'    bRet = True
'    For i = 0 To lngFileListCnt
'        If i = lngFileListCnt Then
'            Exit For
'        End If
'
'        '�t�@�C�������擾����B
'        Input #iFileNumber, strWork
'        If strWork <> "" And Left$(strWork, 1) <> "/" Then  '�t�@�C���������݂���
'            '�t�@�C������`�Ȃ�
'            If strWork = "" Then
'                '���[�v����
'                MsgBox "�t�@�C�������ُ�ł��B" _
'                        & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                        vbOKOnly + vbExclamation, _
'                        "�������D�@ �o�[�W�����Ǘ�"
'                bRet = False
'                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                Exit For
'            '�t�H�[�}�b�g�ُ�
'            ElseIf " " <> Mid(strWork, 2, 1) Then
'              '���[�v����
'                MsgBox "�t�@�C�������ُ�ł��B" _
'                        & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                        vbOKOnly + vbExclamation, _
'                        "�������D�@ �o�[�W�����Ǘ�"
'                bRet = False
'                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                Exit For
'            ElseIf (InStr(strWork, ".") - 1) = -1 Then
'                MsgBox "�t�@�C�������ُ�ł��B" _
'                        & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                        vbOKOnly + vbExclamation, _
'                        "�������D�@ �o�[�W�����Ǘ�"
'                bRet = False
'                Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                Exit For
'            Else
'                '�t�@�C�����݂̂𒊏o
'                sGetFileListName = Mid(strWork, 3, 16)
'                '�擾�t�@�C�����̃T�C�Y���擾
'                myLen = LenB(StrConv(sGetFileListName, vbFromUnicode))      '���p���Z�̃o�C�g�����擾
'                If FILE_NAME_MAX_SIZE < myLen Then
'                    '13�o�C�g�ȏ�̏ꍇ
'                    MsgBox "�t�@�C�������ُ�ł��B" _
'                            & Chr(vbKeyReturn) & "�t�@�C�����X�g���m�F���Ă��������B", _
'                            vbOKOnly + vbExclamation, _
'                            "�������D�@ �o�[�W�����Ǘ�"
'                    bRet = False
'                    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KAISATU_VERSION_KANRI_FILENAMESIZECHK_ERROR, 0)
'                    Exit For
'                End If
'            End If
'        End If
'    Next
'
'    If bRet = False Then
'        '�u�v���O���������M�ُ��ԁv
'        lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)
'
'        '�Ď��ݒ�G���A�u�v���O���������M�ُ��ԁv���X�V
'        Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_FileMaxChk)
'
'        '�ă}�v���Z�X�Ɂu��ԕω��ʒm�v�𑗐M
'        If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_FileMaxChk Then
'            Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_FileMaxChk)
'        End If
'    End If
'    '�t�@�C�����X�g���N���[�Y�B
'    Close #iFileNumber
'    pubfuncCommonGateCheck = bRet

FileGetError:
    If iFileNumber > 0 Then
       Close #iFileNumber
    End If
    iFileNumber = 0
    pubfuncCommonGateCheck = False
    
    '�u�v���O���������M�ُ��ԁv
    lngPgmHanteiRcvErrSts = pfGetKansiSet(IdKansiSet.PG_HANTEI_RCVERR_STS)

    '�Ď��ݒ�G���A�u�v���O���������M�ُ��ԁv���X�V
    Call gspfSetKansiSts(IdKansiSet.PG_HANTEI_RCVERR_STS, ErrCode.PgmHantei_FileMaxChk)
            
    '�ă}�v���Z�X�Ɂu��ԕω��ʒm�v�𑗐M
    If lngPgmHanteiRcvErrSts <> ErrCode.PgmHantei_FileMaxChk Then
        Call sSendMailStsChgInf(MailSts.stsErr, ErrCode.PgmHantei_FileMaxChk)
    End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfuncTotalListCount
'//  �@�\����  : �����X�g���̎擾
'//  �@�\�T�v  : �w���ʈȊO�̑��t�@�C�������Z�o����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l               �Ӗ�
'//  �߂�l    : LONG      lResultCount     ����
'//
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�FfrmJVer.fReadFileList���p
'///////////////////////////////////////////////////////////////////
Private Function pfuncTotalListCount() As Long
    Dim lResultCount As Long                ' ���ʌ���
    Dim iLoop As Integer                    ' ���[�v
    
    Dim iFileNumber As Integer              '�t�@�C���ԍ�
    Dim sFileName As String                 '�t�@�C����
    Dim sSrcFileName As String              '�t�@�C����
    Dim iListCnt As Integer                 '�t�@�C���i�[��
    Dim sCorner As String                   '�R�[�i�[�ԍ�
    Dim sGatePath As String                 '�R�[�i�[�ԍ��t�t�@�C���p�X
    Dim objFso As New FileSystemObject      ' �t�@�C���V�X�e���I�u�W�F�N�g

    On Error GoTo ErrorHandler      '�G���[�n���h���ݒ�
    
    
    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sCorner = Format(iTab_index + 1, "00")
    sGatePath = PATH_N_GATE & sCorner
    
    lResultCount = 0
    iFileNumber = FreeFile   '���g�p�̃t�@�C���ԍ����擾����
    For iLoop = 0 To 8
        
        iFileNumber = FreeFile   '���g�p�̃t�@�C���ԍ����擾����
        sSrcFileName = sGatePath & FolderName(1, iLoop) & "\" & MN_FILELIST
   
        If objFso.FileExists(sSrcFileName) = True Then
   
            Open sSrcFileName For Input Access Read As #iFileNumber     '�t�@�C�����X�g�̃I�[�v��
            iListCnt = 0
            Do While Not EOF(iFileNumber)                               '�t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
                Line Input #iFileNumber, sFileName                      '�f�[�^��ǂݍ��݂܂��B
                If sFileName <> "" And Left$(sFileName, 1) <> "/" Then  '�t�@�C���������݂���
                    iListCnt = iListCnt + 1                             '�t�@�C�����̃J�E���^���A�b�v����
                End If
            Loop
            Close #iFileNumber      '�t�@�C������܂��B
            iFileNumber = 0
            If iLoop <> FolderSyubetu Then
                lResultCount = lResultCount + iListCnt
            End If
        End If
    Next

    pfuncTotalListCount = lResultCount    '�߂�l��ݒ肷��
    Set objFso = Nothing

    Exit Function           '�������I������

'*********************
'* �G���[�n���h������ *
'*********************
ErrorHandler:   ' �G���[�������[�`���B
    If iFileNumber > 0 Then
        Close #iFileNumber
    End If
    pfuncTotalListCount = 0    '�߂�l��ݒ肷��
    Set objFso = Nothing
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : pfuncCopyPASSINF
'//  �@�\����  : ���s�t�H���_�ւ�PASSINF�R�s�[
'//  �@�\�T�v  : �w���ʈȊO�̑��t�@�C�������Z�o����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   nCorner   �R�[�i�ԍ��i0�`5�j
'//  ����      : Integer    nKind           MN_FOLD_WRK(0):���[�N
'//                                         MN_FOLD_NOW(1):���s
'//                                         MN_FOLD_OLD(2):��
'//
'//              �^        �l               �Ӗ�
'//  �߂�l    : BOOL      TRUE             ����
'//            : BOOL      FALSE            �ُ�
'//
'//     ORIGINAL  :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�FfrmJVer.fReadFileList���p
'///////////////////////////////////////////////////////////////////
Private Function pfuncCopyPASSINF(nCorner As Integer, nKind As Integer) As Boolean
    
    Dim objFso As New FileSystemObject      ' �t�@�C���V�X�e���I�u�W�F�N�g
    Dim szSrcFile As String                 ' �R�s�[���t�@�C��
    Dim szDstFile As String                 ' �R�s�[��t�@�C��
    Dim sCorner As String           '�R�[�i�[�ԍ�
    Dim sGatePath As String         '�R�[�i�[�ԍ��t�t�@�C���p�X

    On Error GoTo ErrorHandler              ' �G���[�n���h���̓o�^

    ' �Ώۂ�����f�[�^�̏ꍇ�̂ݏ������s��
    ' ��L�ɊY�����Ȃ��ꍇ�͐���I��
    If FolderSyubetu <> 0 Then
        pfuncCopyPASSINF = True
        Set objFso = Nothing
        Exit Function
    End If

    ' �R�[�i�[�ԍ��t�t�@�C���p�X�쐬
    sCorner = Format(nCorner + 1, "00")
    sGatePath = PATH_N_GATE & sCorner
    ' �R�s�[���t�@�C��
    szSrcFile = sGatePath & FolderName(nKind, 0) & "\" & "PASSINF"
    szDstFile = sGatePath & FolderName(MN_FOLD_NOW, 0) & "\" & "PASSINF"

    If objFso.FileExists(szSrcFile) = True Then
        '�t�@�C���R�s�[�i���ɑ��݂����ꍇ�͏㏑�����邷��j
        objFso.CopyFile szSrcFile, szDstFile, True
        pfuncCopyPASSINF = True
    Else
        pfuncCopyPASSINF = False
    End If

    Set objFso = Nothing
    Exit Function

ErrorHandler:
    pfuncCopyPASSINF = False
    Set objFso = Nothing
End Function

