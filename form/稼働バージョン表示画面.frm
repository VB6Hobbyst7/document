VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmKadoVerKanri 
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
   Begin VB.Timer tmrMail 
      Left            =   0
      Top             =   0
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
      Height          =   4620
      ItemData        =   "�ғ��o�[�W�����\�����.frx":0000
      Left            =   840
      List            =   "�ғ��o�[�W�����\�����.frx":0007
      TabIndex        =   16
      Top             =   3720
      Width           =   8175
   End
   Begin VB.ComboBox cmbGokiSelect 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "�ғ��o�[�W�����\�����.frx":0050
      Left            =   9360
      List            =   "�ғ��o�[�W�����\�����.frx":0052
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   15
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton cmdOutput 
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
      Left            =   9360
      TabIndex        =   3
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton cmdEject 
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
      Left            =   9360
      TabIndex        =   2
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton cmdModoru_Menu 
      Caption         =   "   �����e�i���X   ��ʂ֖߂�"
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
   Begin TabDlg.SSTab tabCorner 
      Height          =   8655
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   15266
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   706
      TabMaxWidth     =   3475
      TabCaption(0)   =   "   �������������@ ������������"
      TabPicture(0)   =   "�ғ��o�[�W�����\�����.frx":0054
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(1)=   "lblKan(5)"
      Tab(0).Control(2)=   "lblKan(4)"
      Tab(0).Control(3)=   "lblKan(3)"
      Tab(0).Control(4)=   "lblKan(2)"
      Tab(0).Control(5)=   "lblTogoVer_Data(0)"
      Tab(0).Control(6)=   "lblLDUVer_Data(0)"
      Tab(0).Control(7)=   "Label1(3)"
      Tab(0).Control(8)=   "lblIDUVer_Data(0)"
      Tab(0).Control(9)=   "Label1(2)"
      Tab(0).Control(10)=   "lblTakuVer_Data(0)"
      Tab(0).Control(11)=   "lblZenVer_Data(0)"
      Tab(0).Control(12)=   "Label3(0)"
      Tab(0).Control(13)=   "lblStationName(0)"
      Tab(0).Control(14)=   "Label1(1)"
      Tab(0).Control(15)=   "Label1(0)"
      Tab(0).Control(16)=   "lblZenVer(0)"
      Tab(0).Control(17)=   "Label4(0)"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "   �������������@ ������������"
      TabPicture(1)   =   "�ғ��o�[�W�����\�����.frx":0070
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblStationName(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(4)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(5)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblZenVer(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblZenVer_Data(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblTakuVer_Data(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(6)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblIDUVer_Data(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(7)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblLDUVer_Data(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblTogoVer_Data(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblKan(7)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label4(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lblKan(0)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lblKan(1)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lblKan(6)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label2(1)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "   �������������@ ������������"
      TabPicture(2)   =   "�ғ��o�[�W�����\�����.frx":008C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(2)"
      Tab(2).Control(1)=   "lblKan(10)"
      Tab(2).Control(2)=   "lblKan(9)"
      Tab(2).Control(3)=   "lblKan(8)"
      Tab(2).Control(4)=   "Label4(2)"
      Tab(2).Control(5)=   "lblZenVer_Data(2)"
      Tab(2).Control(6)=   "lblTogoVer_Data(2)"
      Tab(2).Control(7)=   "lblLDUVer_Data(2)"
      Tab(2).Control(8)=   "Label1(11)"
      Tab(2).Control(9)=   "lblIDUVer_Data(2)"
      Tab(2).Control(10)=   "Label1(10)"
      Tab(2).Control(11)=   "lblTakuVer_Data(2)"
      Tab(2).Control(12)=   "lblKan(11)"
      Tab(2).Control(13)=   "lblZenVer(2)"
      Tab(2).Control(14)=   "Label1(9)"
      Tab(2).Control(15)=   "Label1(8)"
      Tab(2).Control(16)=   "lblStationName(2)"
      Tab(2).Control(17)=   "Label3(2)"
      Tab(2).ControlCount=   18
      TabCaption(3)   =   "   �������������@ ������������"
      TabPicture(3)   =   "�ғ��o�[�W�����\�����.frx":00A8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label2(3)"
      Tab(3).Control(1)=   "lblZenVer_Data(3)"
      Tab(3).Control(2)=   "lblTogoVer_Data(3)"
      Tab(3).Control(3)=   "lblLDUVer_Data(3)"
      Tab(3).Control(4)=   "Label1(15)"
      Tab(3).Control(5)=   "lblIDUVer_Data(3)"
      Tab(3).Control(6)=   "Label1(14)"
      Tab(3).Control(7)=   "lblTakuVer_Data(3)"
      Tab(3).Control(8)=   "lblKan(15)"
      Tab(3).Control(9)=   "lblKan(14)"
      Tab(3).Control(10)=   "lblKan(13)"
      Tab(3).Control(11)=   "lblKan(12)"
      Tab(3).Control(12)=   "lblZenVer(3)"
      Tab(3).Control(13)=   "Label1(13)"
      Tab(3).Control(14)=   "Label1(12)"
      Tab(3).Control(15)=   "lblStationName(3)"
      Tab(3).Control(16)=   "Label3(3)"
      Tab(3).Control(17)=   "Label4(3)"
      Tab(3).ControlCount=   18
      TabCaption(4)   =   "   �������������@ ������������"
      TabPicture(4)   =   "�ғ��o�[�W�����\�����.frx":00C4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label2(4)"
      Tab(4).Control(1)=   "lblZenVer_Data(4)"
      Tab(4).Control(2)=   "lblTogoVer_Data(4)"
      Tab(4).Control(3)=   "lblLDUVer_Data(4)"
      Tab(4).Control(4)=   "Label1(19)"
      Tab(4).Control(5)=   "lblIDUVer_Data(4)"
      Tab(4).Control(6)=   "Label1(18)"
      Tab(4).Control(7)=   "lblTakuVer_Data(4)"
      Tab(4).Control(8)=   "lblKan(19)"
      Tab(4).Control(9)=   "lblKan(18)"
      Tab(4).Control(10)=   "lblKan(17)"
      Tab(4).Control(11)=   "lblKan(16)"
      Tab(4).Control(12)=   "lblZenVer(6)"
      Tab(4).Control(13)=   "Label1(17)"
      Tab(4).Control(14)=   "Label1(16)"
      Tab(4).Control(15)=   "lblStationName(4)"
      Tab(4).Control(16)=   "Label3(4)"
      Tab(4).Control(17)=   "Label4(4)"
      Tab(4).ControlCount=   18
      TabCaption(5)   =   "   �������������@ ������������"
      TabPicture(5)   =   "�ғ��o�[�W�����\�����.frx":00E0
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label2(5)"
      Tab(5).Control(1)=   "lblZenVer_Data(5)"
      Tab(5).Control(2)=   "lblTogoVer_Data(5)"
      Tab(5).Control(3)=   "lblLDUVer_Data(5)"
      Tab(5).Control(4)=   "Label1(23)"
      Tab(5).Control(5)=   "lblIDUVer_Data(5)"
      Tab(5).Control(6)=   "Label1(22)"
      Tab(5).Control(7)=   "lblTakuVer_Data(5)"
      Tab(5).Control(8)=   "lblKan(23)"
      Tab(5).Control(9)=   "lblKan(22)"
      Tab(5).Control(10)=   "lblKan(21)"
      Tab(5).Control(11)=   "lblKan(20)"
      Tab(5).Control(12)=   "lblZenVer(4)"
      Tab(5).Control(13)=   "Label1(21)"
      Tab(5).Control(14)=   "Label1(20)"
      Tab(5).Control(15)=   "lblStationName(5)"
      Tab(5).Control(16)=   "Label3(5)"
      Tab(5).Control(17)=   "Label4(5)"
      Tab(5).ControlCount=   18
      Begin VB.Label Label2 
         Caption         =   "�w"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -63600
         TabIndex        =   114
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "�w"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -63600
         TabIndex        =   113
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "�w"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -63600
         TabIndex        =   112
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "�w"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -63600
         TabIndex        =   111
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "�w"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   11400
         TabIndex        =   110
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "�w"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -63600
         TabIndex        =   109
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�N��������"
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
         Index           =   5
         Left            =   -68180
         TabIndex        =   5
         Top             =   3000
         Width           =   2195
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ғ��o�[�W����"
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
         Index           =   4
         Left            =   -71280
         TabIndex        =   6
         Top             =   3000
         Width           =   3105
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "Ұ���"
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
         Index           =   3
         Left            =   -72360
         TabIndex        =   7
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�@�햼"
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
         Index           =   2
         Left            =   -74160
         TabIndex        =   8
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ғ��o�[�W����"
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
         Index           =   6
         Left            =   3720
         TabIndex        =   39
         Top             =   3000
         Width           =   3105
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "Ұ���"
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
         Left            =   2640
         TabIndex        =   38
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�@�햼"
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
         Index           =   0
         Left            =   840
         TabIndex        =   37
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ғ��o�[�W����"
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
         Index           =   10
         Left            =   -71280
         TabIndex        =   49
         Top             =   3000
         Width           =   3105
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "Ұ���"
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
         Index           =   9
         Left            =   -72360
         TabIndex        =   48
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�@�햼"
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
         Index           =   8
         Left            =   -74160
         TabIndex        =   47
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "�ʘH�ғ��o�[�W����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -74160
         TabIndex        =   41
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "�ʘH�ғ��o�[�W����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   36
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label lblZenVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -72120
         TabIndex        =   108
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblTogoVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -72120
         TabIndex        =   107
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblLDUVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -65280
         TabIndex        =   106
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�k�c�t�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   23
         Left            =   -66600
         TabIndex        =   105
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblIDUVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -68760
         TabIndex        =   104
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�h�c�t�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   22
         Left            =   -70080
         TabIndex        =   103
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblTakuVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -72120
         TabIndex        =   102
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�N��������"
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
         Index           =   23
         Left            =   -68180
         TabIndex        =   101
         Top             =   3000
         Width           =   2195
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ғ��o�[�W����"
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
         Index           =   22
         Left            =   -71280
         TabIndex        =   100
         Top             =   3000
         Width           =   3105
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "Ұ���"
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
         Index           =   21
         Left            =   -72360
         TabIndex        =   99
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�@�햼"
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
         Index           =   20
         Left            =   -74160
         TabIndex        =   98
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblZenVer 
         Caption         =   "�����Ď��ՑS�́F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -74640
         TabIndex        =   97
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "�����Ď��Ձ@�@�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   21
         Left            =   -74640
         TabIndex        =   96
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "�����@�@�@�@�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   20
         Left            =   -74640
         TabIndex        =   95
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblStationName 
         Alignment       =   1  '�E����
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -65760
         TabIndex        =   94
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "���@�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -65640
         TabIndex        =   93
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "�ʘH�ғ��o�[�W����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -74160
         TabIndex        =   92
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label lblZenVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -72120
         TabIndex        =   91
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblZenVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -72120
         TabIndex        =   90
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblZenVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -72120
         TabIndex        =   89
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblTogoVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -72120
         TabIndex        =   88
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblLDUVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -65280
         TabIndex        =   87
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�k�c�t�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   19
         Left            =   -66600
         TabIndex        =   86
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblIDUVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -68760
         TabIndex        =   85
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�h�c�t�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   18
         Left            =   -70080
         TabIndex        =   84
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblTakuVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -72120
         TabIndex        =   83
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�N��������"
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
         Index           =   19
         Left            =   -68180
         TabIndex        =   82
         Top             =   3000
         Width           =   2195
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ғ��o�[�W����"
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
         Index           =   18
         Left            =   -71280
         TabIndex        =   81
         Top             =   3000
         Width           =   3105
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "Ұ���"
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
         Index           =   17
         Left            =   -72360
         TabIndex        =   80
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�@�햼"
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
         Index           =   16
         Left            =   -74160
         TabIndex        =   79
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblZenVer 
         Caption         =   "�����Ď��ՑS�́F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   -74640
         TabIndex        =   78
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "�����Ď��Ձ@�@�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   17
         Left            =   -74640
         TabIndex        =   77
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "�����@�@�@�@�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   -74640
         TabIndex        =   76
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblStationName 
         Alignment       =   1  '�E����
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -65760
         TabIndex        =   75
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "���@�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -65640
         TabIndex        =   74
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "�ʘH�ғ��o�[�W����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -74160
         TabIndex        =   73
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label lblTogoVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -72120
         TabIndex        =   72
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblLDUVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -65280
         TabIndex        =   71
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�k�c�t�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   -66600
         TabIndex        =   70
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblIDUVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -68760
         TabIndex        =   69
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�h�c�t�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   14
         Left            =   -70080
         TabIndex        =   68
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblTakuVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -72120
         TabIndex        =   67
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�N��������"
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
         Index           =   15
         Left            =   -68180
         TabIndex        =   66
         Top             =   3000
         Width           =   2195
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�ғ��o�[�W����"
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
         Index           =   14
         Left            =   -71280
         TabIndex        =   65
         Top             =   3000
         Width           =   3105
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "Ұ���"
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
         Index           =   13
         Left            =   -72360
         TabIndex        =   64
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�@�햼"
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
         Index           =   12
         Left            =   -74160
         TabIndex        =   63
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblZenVer 
         Caption         =   "�����Ď��ՑS�́F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -74640
         TabIndex        =   62
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "�����Ď��Ձ@�@�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   -74640
         TabIndex        =   61
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "�����@�@�@�@�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   -74640
         TabIndex        =   60
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblStationName 
         Alignment       =   1  '�E����
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -65760
         TabIndex        =   59
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "���@�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -65640
         TabIndex        =   58
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "�ʘH�ғ��o�[�W����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -74160
         TabIndex        =   57
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label lblTogoVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -72120
         TabIndex        =   56
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblLDUVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -65280
         TabIndex        =   55
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�k�c�t�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   -66600
         TabIndex        =   54
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblIDUVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -68760
         TabIndex        =   53
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�h�c�t�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   -70080
         TabIndex        =   52
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblTakuVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -72120
         TabIndex        =   51
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�N��������"
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
         Index           =   11
         Left            =   -68180
         TabIndex        =   50
         Top             =   3000
         Width           =   2195
      End
      Begin VB.Label lblZenVer 
         Caption         =   "�����Ď��ՑS�́F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -74640
         TabIndex        =   46
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "�����Ď��Ձ@�@�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   -74640
         TabIndex        =   45
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "�����@�@�@�@�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   -74640
         TabIndex        =   44
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblStationName 
         Alignment       =   1  '�E����
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -65760
         TabIndex        =   43
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "���@�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -65640
         TabIndex        =   42
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblKan 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�N��������"
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
         Index           =   7
         Left            =   6820
         TabIndex        =   40
         Top             =   3000
         Width           =   2195
      End
      Begin VB.Label lblTogoVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   35
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblLDUVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   9720
         TabIndex        =   34
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�k�c�t�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   8400
         TabIndex        =   33
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblIDUVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6240
         TabIndex        =   32
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�h�c�t�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   4920
         TabIndex        =   31
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblTakuVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   30
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblZenVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   29
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblZenVer 
         Caption         =   "�����Ď��ՑS�́F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   28
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "�����Ď��Ձ@�@�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   360
         TabIndex        =   27
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "�����@�@�@�@�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   360
         TabIndex        =   26
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblStationName 
         Alignment       =   1  '�E����
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   9360
         TabIndex        =   25
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "���@�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   9360
         TabIndex        =   24
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblTogoVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -72120
         TabIndex        =   23
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblLDUVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -65280
         TabIndex        =   22
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�k�c�t�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   -66600
         TabIndex        =   21
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblIDUVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -68760
         TabIndex        =   20
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "�h�c�t�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   -70080
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblTakuVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -72120
         TabIndex        =   18
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblZenVer_Data 
         Caption         =   "Z9.Z9.Z9.Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -72120
         TabIndex        =   17
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "���@�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -65640
         TabIndex        =   13
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblStationName 
         Alignment       =   1  '�E����
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -65760
         TabIndex        =   12
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "�����@�@�@�@�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   -74640
         TabIndex        =   11
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "�����Ď��Ձ@�@�F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   -74640
         TabIndex        =   10
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label lblZenVer 
         Caption         =   "�����Ď��ՑS�́F"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   9
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "�ʘH�ғ��o�[�W����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -74160
         TabIndex        =   14
         Top             =   2640
         Width           =   3015
      End
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�ғ�Ver�ꗗ�\��"
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
      TabIndex        =   0
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "frmKadoVerKanri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmKadoVerKanri.frm
'//  �p�b�P�[�W���F�ғ��o�[�W�����Ǘ����
'//
'//  �T�v�F�ғ��o�[�W�����Ǘ����
'//     ORIGINAL  :(EG20 V5.2.0.1) 2012-03-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000       '���C���^�C�}�̃C���^�[�o���l
Private Const MAX_GOUKI = 15                '�ő卆�@�l�i�P�R�[�i������j
Private mintCurCornerIdx As Integer         '�I���R�[�iIndex

Private Const PATH_DISP_FILE = PATH_WORK & "KadoVerDisp.csv"    '��ʏo�͗p�t�@�C��
Private Const FILE_KADOVER = "_KADOVER.txt"                     '�}�̏o�͗p�t�@�C��
Private Const LEN_KISHU = 15            '�@�햼����
Private Const LEN_MAKER = 9             '���[�J������
Private Const LEN_VERSION = 26          '�o�[�W��������
Private Const LEN_DATE = 17             '���t����

Private Enum mintDispDiv
    KADOVER_FILE_DISP = 0
    KADOVER_FILE_OUTPUT
End Enum
    


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : cmbGokiSelect_Click
'//  �@�\����  : ���@�I���R���{�{�b�N�X�N���b�N����
'//  �@�\�T�v  : ��ʂ̕\�����e���X�V����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.2.0.1) 2012-03-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmbGokiSelect_Click()

    On Error Resume Next
    
    '�\���X�V
    Call Change_Disp
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : cmdEject_Click
'//  �@�\����  : �}�̎�O�{�^����������
'//  �@�\�T�v  : �}�̎�O���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.2.0.1) 2012-03-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdEject_Click()

    On Error Resume Next
    
   '�u�}�̎�O�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '�}�̎�O����
    Call pfRemove(Me)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : cmdModoru_Menu_Click
'//  �@�\����  : �o�[�W�����Ǘ���ʂɖ߂�{�^����������
'//  �@�\�T�v  : �o�[�W�����Ǘ���ʂɖ߂�B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.2.0.1) 2012-03-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdModoru_Menu_Click()

    On Error Resume Next

    '�u�ғ��o�[�W�����Ǘ���ʁF�����v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KADOVER_KANSI_LOG_GAMEN_END, 0)
  
    '�ғ��o�[�W�����Ǘ���ʂ����
    Unload Me
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : cmdOutput_Click
'//  �@�\����  : �}�̏o�̓{�^������������
'//  �@�\�T�v  : �ғ��o�[�W�����t�@�C����}�̏o�͂���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.2.0.1) 2012-03-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V5.4.0.1) 2012-03-24   CODED   BY [TCC] M.Matsumoto
'//                 �y����No54�Ή��z
'//     REVISIONS :(EG20 V30.1.0.1) 2014-05-07  CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdOutPut_Click()

    Dim strDirName As String            '�o�͐�t�H���_
    Dim strOutputFile As String         '�o�̓t�@�C��
    Dim lngRet As Long                  '�֐��Ԃ�l
    Dim lngGokiNo As Long               '���@�ԍ�
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    On Error GoTo Err_Handler
    
    '�o�̓t�H���_�I��
    strDirName = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)
    If strDirName = "" Then
'        Unload Me          'EG20 V5.4.0.1 DELL                 'EG20 V5.4.0.1 DEL �y����No54�Ή��z
        Exit Sub  '�f�B���N�g�����w�肳��Ȃ���΁A�����I��
    End If
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_VERSION_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '�t�@�C�����ҏW
    strOutputFile = strDirName & lblStationName(mintCurCornerIdx).Caption & FILE_KADOVER
    
    '���@�ԍ��擾
    lngGokiNo = cmbGokiSelect.ItemData(cmbGokiSelect.ListIndex)
    
    '�t�@�C���o�͊֐���Call
    'EG20 V30.1.0.1 DEL START
'    lngRet = dllCreateKadoVersionFile(KADOVER_FILE_OUTPUT, CLng(mintCurCornerIdx + 1), _
'                                      lngGokiNo, strOutputFile, PATH_IDU_APP, PATH_LDU_APP)
    'EG20 V30.1.0.1 DEL END
    'EG20 V30.1.0.1 ADD START
    If gintCornerType(mintCurCornerIdx) = CORNER_TYPE_KANSEN Then
        lngRet = dllCreateKadoVersionFileKan(KADOVER_FILE_OUTPUT, CLng(mintCurCornerIdx + 1), _
                                        lngGokiNo, strOutputFile, PATH_IDU_APP, PATH_LDU_APP)
    Else
        lngRet = dllCreateKadoVersionFile(KADOVER_FILE_OUTPUT, CLng(mintCurCornerIdx + 1), _
                                        lngGokiNo, strOutputFile, PATH_IDU_APP, PATH_LDU_APP)
    End If
    'EG20 V30.1.0.1 ADD END
    
    '�ُ�I�����̓G���[������
    If lngRet = 0 Then
        GoTo Err_Handler
        Exit Sub
    End If
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    MsgBox "����I�����܂����B", vbInformation + vbOKOnly, "�o�͌���"
    
    Exit Sub

Err_Handler:
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
    '�u�ғ��o�[�W�����Ǘ���ʁF�ғ��o�[�W�������}�̏o�͏����ُ�v���O�o��
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KADOVER_INFO_OUTPUT_ERROR, lngErrCode)
                                      
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �ғ��o�[�W�����Ǘ����(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.2.0.1) 2012-03-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    '���C����M�p�̃^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �ғ��o�[�W�����Ǘ����(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.2.0.1) 2012-03-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    '���C����M�p�̃^�C�}���~�߂�B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �ғ��o�[�W�����Ǘ����(���[�h��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.2.0.1) 2012-03-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V6.8.0.1) 2012-08-28  CODED BY  [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V30.1.0.1) 2014-05-07  CODED BY  [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()

    Dim strCorner1 As String        '�R�[�i���i��i�j
    Dim strCorner2 As String        '�R�[�i���i���i�j
    Dim intCount As Integer         '�J�E���^
    
    On Error Resume Next
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
' EG20 V6.8.0.1 ADD START
    '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
' EG20 V6.8.0.1 ADD END
    
    Call gsGetGateInfo          '���@���擾
    Call gsGetCornerName        '�R�[�i���擾
    Call gsGetStationName       '�w���擾
    Call gsGetCornerType        '�R�[�i�^�C�v�擾   'EG20 V30.1.0.1 ADD
    
    tabCorner.Tab = 0
    
    For intCount = 0 To UBound(gblnCornerSet)
    
        '�w����\������
        lblStationName(intCount).Caption = gstrStationName(intCount)
        
        '�ݒ肠��̃R�[�i
        If gblnCornerSet(intCount) = True Then
            '�R�[�i�[���̕\��
            strCorner1 = MidB(gstrCornerName(intCount), 1, 12)
            strCorner2 = MidB(gstrCornerName(intCount), 13, 24)
            tabCorner.TabCaption(intCount) = strCorner1 & vbCrLf & strCorner2
        '�ݒ�Ȃ��̃R�[�i���\���ɂ���
        Else
            tabCorner.TabVisible(intCount) = False
        End If
    
    Next intCount
    
    '�\������
    Call tabCorner_Click(0)
   
    '�u�ғ��o�[�W�����Ǘ���ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KADOVER_KANSI_LOG_GAMEN_START, 0)
    
    Exit Sub

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : tabCorner_Click
'//  �@�\����  : �^�u�N���b�N������
'//  �@�\�T�v  : �I���R�[�i�̕\���ɍX�V����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.2.0.1) 2012-03-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub tabCorner_Click(PreviousTab As Integer)

    Dim intCount As Integer     '�J�E���^
    Dim intIndex As Integer     '�R���{�{�b�N�X�̃��X�g�C���f�b�N�X
    
    On Error Resume Next
    
    '�R�[�iIndex��ݒ�
    mintCurCornerIdx = tabCorner.Tab
    
    '���@�R���{�{�b�N�X���N���A����
    cmbGokiSelect.Clear
    intIndex = 0
    
    '�I�𒆂̃R�[�i�̍��@�����[�v����
    For intCount = 0 To UBound(gudtSettiCorner(mintCurCornerIdx).intGokiNo)
        '�L���ȍ��@�̏ꍇ
        If gudtSettiCorner(mintCurCornerIdx).intGokiNo(intCount) > 0 Then
            '�R���{�{�b�N�X�ɍ��@�ԍ���\��
            cmbGokiSelect.AddItem CStr(gudtSettiCorner(mintCurCornerIdx).intGokiNo(intCount)) & "���@"
            'ItemData�ɘ_�����@�ԍ����L�^����
            cmbGokiSelect.ItemData(intIndex) = gudtSettiCorner(mintCurCornerIdx).intGokiNo(intCount)
            intIndex = intIndex + 1
        End If
    Next
    
    '�f�t�H���g�͐擪���@
    cmbGokiSelect.ListIndex = 0
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
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
'//     ORIGINAL  :(EG20 V5.2.0.1) 2012-03-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    '�ėp���[����M�������s��
    If pfVersionDispMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmKadoVerKanri.Caption, False
        pfFormActive (frmKadoVerKanri.hwnd)     ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : Change_Disp
'//  �@�\����  : �\�����e�X�V
'//  �@�\�T�v  : �I�����ꂽ�R�[�i�A���@�ɂ���ʕ\�����e���X�V����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V5.2.0.1) 2012-03-05   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(EG20 V30.1.0.1) 2014-05-07  CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub Change_Disp()

    Dim bySyoAssort As Byte             '���O�p������
    Dim lngGokiNo As Long               '���@�ԍ�
    Dim lngRet As Long                  'DLL�Ԃ�l
    Dim intFileNo As Integer            '�t�@�C���ԍ�
    Dim intKishu As Integer             '�@�핪�ށi�t�@�C���ǂݍ��ݗp�j
    Dim intCorner As Integer            '�R�[�i���ށi�t�@�C���ǂݍ��ݗp�j
    Dim intGokiDiv As Integer           '���@���ށi�t�@�C���ǂݍ��ݗp�j
    Dim strName As String               '�@�햼�i�t�@�C���ǂݍ��ݗp�j
    Dim strMaker As String              '���[�J���i�t�@�C���ǂݍ��ݗp�j
    Dim strVer As String                '�o�[�W�����i�t�@�C���ǂݍ��ݗp�j
    Dim strDate As String               '�쐬���t�i�t�@�C���ǂݍ��ݗp�j
    Dim strDsp_Kishu As String          '�@�햼�i��ʕ\���p�j
    Dim strDsp_Maker As String          '���[�J���i��ʕ\���p�j
    Dim strDsp_Version As String        '�o�[�W�����i��ʕ\���p�j
    Dim strDsp_Date As String           '�쐬���t�i��ʕ\���p�j
    Dim objFs As FileSystemObject       '�t�@�C���V�X�e���I�u�W�F�N�g
    
    On Error Resume Next
    
    '�t�@�C���L���`�F�b�N
    Set objFs = New FileSystemObject
    
    '���@�ԍ��擾
    lngGokiNo = cmbGokiSelect.ItemData(cmbGokiSelect.ListIndex)
    
    '��ʕ\���p�t�@�C���쐬�֐���Call
    'EG20 V30.1.0.1 DEL START
'    lngRet = dllCreateKadoVersionFile(KADOVER_FILE_DISP, CLng(mintCurCornerIdx + 1), _
'                                      lngGokiNo, PATH_DISP_FILE, PATH_IDU_APP, PATH_LDU_APP)
    'EG20 V30.1.0.1 DEL END
    'EG20 V30.1.0.1 ADD START
    If gintCornerType(mintCurCornerIdx) = CORNER_TYPE_KANSEN Then
        lngRet = dllCreateKadoVersionFileKan(KADOVER_FILE_DISP, CLng(mintCurCornerIdx + 1), _
                                          lngGokiNo, PATH_DISP_FILE, PATH_IDU_APP, PATH_LDU_APP)
    Else
        lngRet = dllCreateKadoVersionFile(KADOVER_FILE_DISP, CLng(mintCurCornerIdx + 1), _
                                          lngGokiNo, PATH_DISP_FILE, PATH_IDU_APP, PATH_LDU_APP)
    End If
    'EG20 V30.1.0.1 ADD END
    '�ُ�I���̏ꍇ�̓G���[������
    If lngRet = 0 Then
        GoTo Err_Handler
        Exit Sub
    End If
    
    '�t�@�C�������݂��Ȃ��ꍇ�̓G���[������
    If objFs.FileExists(PATH_DISP_FILE) = False Then
        GoTo Err_Handler
        Exit Sub
    End If
    
    '��ʕ\���p�t�@�C�����I�[�v��
    intFileNo = FreeFile
    Open PATH_DISP_FILE For Input As #intFileNo
    
    lstKan.Clear
    '��ʕ\������
    Do While Not EOF(intFileNo)
    
        Input #intFileNo, intKishu, intCorner, intGokiDiv, strName, strMaker, strVer, strDate
        
        Select Case intKishu
        Case 1  '�S��
            lblZenVer_Data(mintCurCornerIdx).Caption = strVer
        Case 2  '�����Ď���
            lblTogoVer_Data(mintCurCornerIdx).Caption = strVer
        Case 3  '�h�c�t
            lblIDUVer_Data(mintCurCornerIdx).Caption = strVer
        Case 4  '�k�c�t
            lblLDUVer_Data(mintCurCornerIdx).Caption = strVer
        Case 5  '�����
            lblTakuVer_Data(mintCurCornerIdx).Caption = strVer
        Case 6  '�ʘH�ғ�
        
            '�e���ڂ��X�y�[�X�Ő��`����
            strDsp_Kishu = StrConv(MidB(StrConv(strName & Space(LEN_KISHU), vbFromUnicode), 1, LEN_KISHU), vbUnicode)
            strDsp_Maker = StrConv(MidB(StrConv(strMaker & Space(LEN_MAKER), vbFromUnicode), 1, LEN_MAKER), vbUnicode)
            strDsp_Version = StrConv(MidB(StrConv(strVer & Space(LEN_VERSION), vbFromUnicode), 1, LEN_VERSION), vbUnicode)
            strDsp_Date = StrConv(MidB(StrConv(strDate & Space(LEN_DATE), vbFromUnicode), 1, LEN_DATE), vbUnicode)
            
            '���X�g�\��
            lstKan.AddItem (strDsp_Kishu & strDsp_Maker & strDsp_Version & strDsp_Date)
            
        End Select
    Loop
    
    '�t�@�C���N���[�Y
    Close #intFileNo
    
    Set objFs = Nothing
    
    Exit Sub
    
Err_Handler:

    '�t�@�C���N���[�Y
    If intFileNo > 0 Then
        Close #intFileNo
    End If

    '�o�[�W���������N���A����
    lblZenVer_Data(mintCurCornerIdx).Caption = Empty
    lblTogoVer_Data(mintCurCornerIdx).Caption = Empty
    lblIDUVer_Data(mintCurCornerIdx).Caption = Empty
    lblLDUVer_Data(mintCurCornerIdx).Caption = Empty
    lblTakuVer_Data(mintCurCornerIdx).Caption = Empty
    lstKan.Clear
    Set objFs = Nothing
    
    '�G���[���O�̏o��
    Call sLogTraceReq(LTYP_ERROR, bySyoAssort, KADOVER_INFO_DISP_ERROR, 0)
    
End Sub
