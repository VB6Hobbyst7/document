VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTakuLogKanri 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�Ď��Ճ��O�Ǘ�"
   ClientHeight    =   9000
   ClientLeft      =   2445
   ClientTop       =   1395
   ClientWidth     =   12000
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �̊���l
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
      Height          =   975
      Left            =   9720
      TabIndex        =   36
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "  ���O�\��    (�e�L�X�g�\��)"
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
      Index           =   0
      Left            =   9720
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "  ���O�Ǘ�     ��ʂ֖߂�"
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
      Left            =   9720
      TabIndex        =   1
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Timer tmrMail 
      Left            =   9600
      Top             =   6840
   End
   Begin TabDlg.SSTab tabLog 
      Height          =   8620
      Left            =   0
      TabIndex        =   0
      Top             =   380
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   15214
      _Version        =   393216
      Tab             =   2
      TabHeight       =   706
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "�\���t�@�C���w��"
      TabPicture(0)   =   "���O�Ǘ�(�����)���.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdLogShushu"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdLzhFileWrite"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdLog(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdUpdateDisplay"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tabTakuCorner"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "�\�����ڎw��"
      TabPicture(1)   =   "���O�Ǘ�(�����)���.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraKoumoku(2)"
      Tab(1).Control(1)=   "fraKoumoku(1)"
      Tab(1).Control(2)=   "fraKoumoku(0)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "�\�����@�w��"
      TabPicture(2)   =   "���O�Ǘ�(�����)���.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraGouki"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdLogShushu 
         Caption         =   "���O���W"
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
         Left            =   -68400
         TabIndex        =   242
         Top             =   3480
         Width           =   2055
      End
      Begin VB.CommandButton cmdLzhFileWrite 
         Caption         =   "  ���O���k    �}�̏o��"
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
         Left            =   -68400
         TabIndex        =   241
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton cmdLog 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���O�}�̏o��"
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
         Index           =   1
         Left            =   -68400
         TabIndex        =   240
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton cmdUpdateDisplay 
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
         Left            =   -68400
         TabIndex        =   239
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Frame fraKoumoku 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   6135
         Index           =   2
         Left            =   -74895
         TabIndex        =   29
         Top             =   2400
         Width           =   9135
         Begin VB.Frame fraLogBunnrui 
            Caption         =   "�w�蕪��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   5800
            Left            =   2400
            TabIndex        =   38
            Top             =   240
            Width           =   6615
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   59
               Left            =   4470
               TabIndex        =   100
               Top             =   5400
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   58
               Left            =   4470
               TabIndex        =   99
               Top             =   5160
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   57
               Left            =   4470
               TabIndex        =   98
               Top             =   4920
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   56
               Left            =   4470
               TabIndex        =   97
               Top             =   4680
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   55
               Left            =   4470
               TabIndex        =   96
               Top             =   4440
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   54
               Left            =   4470
               TabIndex        =   95
               Top             =   4200
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   53
               Left            =   4470
               TabIndex        =   94
               Top             =   3960
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   52
               Left            =   4470
               TabIndex        =   93
               Top             =   3720
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   51
               Left            =   4470
               TabIndex        =   92
               Top             =   3480
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   50
               Left            =   4470
               TabIndex        =   91
               Top             =   3240
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   49
               Left            =   4470
               TabIndex        =   90
               Top             =   3000
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   48
               Left            =   4470
               TabIndex        =   89
               Top             =   2760
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   47
               Left            =   4470
               TabIndex        =   88
               Top             =   2520
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   46
               Left            =   4470
               TabIndex        =   87
               Top             =   2280
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   45
               Left            =   4470
               TabIndex        =   86
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   44
               Left            =   4470
               TabIndex        =   85
               Top             =   1800
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   43
               Left            =   4470
               TabIndex        =   84
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   42
               Left            =   4470
               TabIndex        =   83
               Top             =   1320
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   41
               Left            =   4470
               TabIndex        =   82
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   40
               Left            =   4470
               TabIndex        =   81
               Top             =   840
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   39
               Left            =   2295
               TabIndex        =   80
               Top             =   5400
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   38
               Left            =   2295
               TabIndex        =   79
               Top             =   5160
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   37
               Left            =   2295
               TabIndex        =   78
               Top             =   4920
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   36
               Left            =   2295
               TabIndex        =   77
               Top             =   4680
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   35
               Left            =   2295
               TabIndex        =   76
               Top             =   4440
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   34
               Left            =   2295
               TabIndex        =   75
               Top             =   4200
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   33
               Left            =   2295
               TabIndex        =   74
               Top             =   3960
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   32
               Left            =   2295
               TabIndex        =   73
               Top             =   3720
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   31
               Left            =   2295
               TabIndex        =   72
               Top             =   3480
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   30
               Left            =   2295
               TabIndex        =   71
               Top             =   3240
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   29
               Left            =   2295
               TabIndex        =   70
               Top             =   3000
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   28
               Left            =   2295
               TabIndex        =   69
               Top             =   2760
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   27
               Left            =   2295
               TabIndex        =   68
               Top             =   2520
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   26
               Left            =   2295
               TabIndex        =   67
               Top             =   2280
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   25
               Left            =   2295
               TabIndex        =   66
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   24
               Left            =   2295
               TabIndex        =   65
               Top             =   1800
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   23
               Left            =   2295
               TabIndex        =   64
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   22
               Left            =   2295
               TabIndex        =   63
               Top             =   1320
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   21
               Left            =   2295
               TabIndex        =   62
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   20
               Left            =   2295
               TabIndex        =   61
               Top             =   840
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   19
               Left            =   120
               TabIndex        =   60
               Top             =   5400
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   18
               Left            =   120
               TabIndex        =   59
               Top             =   5160
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   17
               Left            =   120
               TabIndex        =   58
               Top             =   4920
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   16
               Left            =   120
               TabIndex        =   57
               Top             =   4680
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   15
               Left            =   120
               TabIndex        =   56
               Top             =   4440
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   120
               TabIndex        =   55
               Top             =   4200
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   13
               Left            =   120
               TabIndex        =   54
               Top             =   3960
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   53
               Top             =   3720
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   52
               Top             =   3480
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   51
               Top             =   3240
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   50
               Top             =   3000
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   49
               Top             =   2760
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   48
               Top             =   2520
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   47
               Top             =   2280
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   46
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   45
               Top             =   1800
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   44
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   43
               Top             =   1320
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   42
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.CheckBox chkMod 
               Caption         =   "12345678901234"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   41
               Top             =   840
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.OptionButton optAll 
               Caption         =   "�S�Ė��I��"
               BeginProperty Font 
                  Name            =   "�l�r �o�S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   1
               Left            =   1920
               TabIndex        =   40
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton optAll 
               Caption         =   "�S�đI��"
               BeginProperty Font 
                  Name            =   "�l�r �o�S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   0
               Left            =   480
               TabIndex        =   39
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.Line Line11 
               BorderColor     =   &H00FFFFFF&
               X1              =   0
               X2              =   6550
               Y1              =   720
               Y2              =   720
            End
         End
         Begin VB.OptionButton optLogBunrui 
            Caption         =   "�w�蕪�ނ̂ݕ\��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Value           =   -1  'True
            Width           =   2275
         End
         Begin VB.OptionButton optLogBunrui 
            Caption         =   "�S�Ă̕��ނ�\��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   2275
         End
         Begin VB.Frame fraLogData 
            Caption         =   "�\���s"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   2295
            Left            =   120
            TabIndex        =   30
            Top             =   1800
            Width           =   2175
            Begin VB.OptionButton optLogData 
               Caption         =   "�P�s�ڂ̂ݕ\��"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   32
               Top             =   1560
               Value           =   -1  'True
               Width           =   2000
            End
            Begin VB.OptionButton optLogData 
               Caption         =   "�S�s�\��"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   31
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label lblLogData 
               Caption         =   "1����Ă������s�̂Ƃ�"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   33
               Top             =   480
               Width           =   1695
            End
         End
      End
      Begin VB.Frame fraKoumoku 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Index           =   1
         Left            =   -70800
         TabIndex        =   20
         Top             =   480
         Width           =   4935
         Begin VB.OptionButton optLogSyu 
            Caption         =   "�w���ʂ̂ݕ\��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton optLogSyu 
            Caption         =   "�S�Ă̎�ʂ�\��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2640
            TabIndex        =   27
            Top             =   360
            Width           =   2250
         End
         Begin VB.Frame fraLogSyu 
            Caption         =   "�w����"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1095
            Index           =   0
            Left            =   360
            TabIndex        =   21
            Top             =   720
            Width           =   4095
            Begin VB.CheckBox chkLogSyu 
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   480
               TabIndex        =   26
               Top             =   240
               Value           =   1  '����
               Width           =   1095
            End
            Begin VB.CheckBox chkLogSyu 
               Caption         =   "�ُ�"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   1560
               TabIndex        =   25
               Top             =   240
               Value           =   1  '����
               Width           =   1095
            End
            Begin VB.CheckBox chkLogSyu 
               Caption         =   "�x��"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   2640
               TabIndex        =   24
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox chkLogSyu 
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   480
               TabIndex        =   23
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox chkLogSyu 
               Caption         =   "�f�o�b�O"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   1560
               TabIndex        =   22
               Top             =   600
               Width           =   1300
            End
         End
      End
      Begin VB.Frame fraKoumoku 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Index           =   0
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   3855
         Begin VB.TextBox txtLogTime 
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   10
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txtLogTime 
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   9
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txtLogTime 
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   840
            MaxLength       =   2
            TabIndex        =   8
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txtLogTime 
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   7
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtLogTime 
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   6
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtLogTime 
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   840
            MaxLength       =   2
            TabIndex        =   5
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lblLogTime 
            Caption         =   "�܂�"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   3000
            TabIndex        =   19
            Top             =   1260
            Width           =   615
         End
         Begin VB.Label lblLogTime 
            Caption         =   "��"
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
            Index           =   7
            Left            =   2680
            TabIndex        =   18
            Top             =   1260
            Width           =   255
         End
         Begin VB.Label lblLogTime 
            Caption         =   "��"
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
            Index           =   6
            Left            =   1960
            TabIndex        =   17
            Top             =   1260
            Width           =   255
         End
         Begin VB.Label lblLogTime 
            Caption         =   "��"
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
            Index           =   5
            Left            =   1220
            TabIndex        =   16
            Top             =   1260
            Width           =   255
         End
         Begin VB.Label lblLogTime 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   3000
            TabIndex        =   15
            Top             =   900
            Width           =   615
         End
         Begin VB.Label lblLogTime 
            Caption         =   "��"
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
            Index           =   3
            Left            =   2680
            TabIndex        =   14
            Top             =   900
            Width           =   255
         End
         Begin VB.Label lblLogTime 
            Caption         =   "��"
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
            Left            =   1960
            TabIndex        =   13
            Top             =   900
            Width           =   255
         End
         Begin VB.Label lblLogTime 
            Caption         =   "��"
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
            Left            =   1220
            TabIndex        =   12
            Top             =   900
            Width           =   255
         End
         Begin VB.Label lblLogTime 
            Caption         =   "���O�f�[�^�Ώێ���"
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
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.Frame fraGouki 
         Caption         =   "�������@"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   7455
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   8895
         Begin VB.CommandButton cmdZSentaku 
            Caption         =   "  �S�R�[�i    �S���@ �I��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   240
            TabIndex        =   201
            Top             =   480
            Width           =   2000
         End
         Begin VB.CommandButton cmdZHisentaku 
            Caption         =   "  �S�R�[�i    �S���@ ��I��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   2400
            TabIndex        =   200
            Top             =   480
            Width           =   2000
         End
         Begin VB.CommandButton cmdHSentaku 
            Caption         =   " �\���R�[�i   �S���@  �I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   4560
            TabIndex        =   199
            Top             =   480
            Width           =   2000
         End
         Begin VB.CommandButton cmdHHisentaku 
            Caption         =   " �\���R�[�i   �S���@ ��I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   6720
            TabIndex        =   198
            Top             =   480
            Width           =   2000
         End
         Begin TabDlg.SSTab tabCorner 
            Height          =   2535
            Left            =   120
            TabIndex        =   101
            Top             =   1440
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   4471
            _Version        =   393216
            Tabs            =   6
            TabsPerRow      =   6
            TabHeight       =   794
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Tab 0"
            TabPicture(0)   =   "���O�Ǘ�(�����)���.frx":0054
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "chkLogGouki(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "chkLogGouki(1)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "chkLogGouki(2)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "chkLogGouki(3)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "chkLogGouki(4)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "chkLogGouki(5)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "chkLogGouki(6)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "chkLogGouki(7)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "chkLogGouki(8)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "chkLogGouki(9)"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "chkLogGouki(10)"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "chkLogGouki(11)"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "chkLogGouki(12)"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "chkLogGouki(13)"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "chkLogGouki(14)"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "chkLogGouki(15)"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).ControlCount=   16
            TabCaption(1)   =   "  "
            TabPicture(1)   =   "���O�Ǘ�(�����)���.frx":0070
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "chkLogGouki(31)"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "chkLogGouki(30)"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "chkLogGouki(29)"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "chkLogGouki(28)"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "chkLogGouki(27)"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "chkLogGouki(26)"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "chkLogGouki(25)"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "chkLogGouki(24)"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "chkLogGouki(23)"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "chkLogGouki(22)"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "chkLogGouki(21)"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).Control(11)=   "chkLogGouki(20)"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).Control(12)=   "chkLogGouki(19)"
            Tab(1).Control(12).Enabled=   0   'False
            Tab(1).Control(13)=   "chkLogGouki(18)"
            Tab(1).Control(13).Enabled=   0   'False
            Tab(1).Control(14)=   "chkLogGouki(17)"
            Tab(1).Control(14).Enabled=   0   'False
            Tab(1).Control(15)=   "chkLogGouki(16)"
            Tab(1).Control(15).Enabled=   0   'False
            Tab(1).ControlCount=   16
            TabCaption(2)   =   "  "
            TabPicture(2)   =   "���O�Ǘ�(�����)���.frx":008C
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "chkLogGouki(47)"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "chkLogGouki(46)"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "chkLogGouki(45)"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "chkLogGouki(44)"
            Tab(2).Control(3).Enabled=   0   'False
            Tab(2).Control(4)=   "chkLogGouki(43)"
            Tab(2).Control(4).Enabled=   0   'False
            Tab(2).Control(5)=   "chkLogGouki(42)"
            Tab(2).Control(5).Enabled=   0   'False
            Tab(2).Control(6)=   "chkLogGouki(41)"
            Tab(2).Control(6).Enabled=   0   'False
            Tab(2).Control(7)=   "chkLogGouki(40)"
            Tab(2).Control(7).Enabled=   0   'False
            Tab(2).Control(8)=   "chkLogGouki(39)"
            Tab(2).Control(8).Enabled=   0   'False
            Tab(2).Control(9)=   "chkLogGouki(38)"
            Tab(2).Control(9).Enabled=   0   'False
            Tab(2).Control(10)=   "chkLogGouki(37)"
            Tab(2).Control(10).Enabled=   0   'False
            Tab(2).Control(11)=   "chkLogGouki(36)"
            Tab(2).Control(11).Enabled=   0   'False
            Tab(2).Control(12)=   "chkLogGouki(35)"
            Tab(2).Control(12).Enabled=   0   'False
            Tab(2).Control(13)=   "chkLogGouki(34)"
            Tab(2).Control(13).Enabled=   0   'False
            Tab(2).Control(14)=   "chkLogGouki(33)"
            Tab(2).Control(14).Enabled=   0   'False
            Tab(2).Control(15)=   "chkLogGouki(32)"
            Tab(2).Control(15).Enabled=   0   'False
            Tab(2).ControlCount=   16
            TabCaption(3)   =   "  "
            TabPicture(3)   =   "���O�Ǘ�(�����)���.frx":00A8
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "chkLogGouki(63)"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).Control(1)=   "chkLogGouki(62)"
            Tab(3).Control(1).Enabled=   0   'False
            Tab(3).Control(2)=   "chkLogGouki(61)"
            Tab(3).Control(2).Enabled=   0   'False
            Tab(3).Control(3)=   "chkLogGouki(60)"
            Tab(3).Control(3).Enabled=   0   'False
            Tab(3).Control(4)=   "chkLogGouki(59)"
            Tab(3).Control(4).Enabled=   0   'False
            Tab(3).Control(5)=   "chkLogGouki(58)"
            Tab(3).Control(5).Enabled=   0   'False
            Tab(3).Control(6)=   "chkLogGouki(57)"
            Tab(3).Control(6).Enabled=   0   'False
            Tab(3).Control(7)=   "chkLogGouki(56)"
            Tab(3).Control(7).Enabled=   0   'False
            Tab(3).Control(8)=   "chkLogGouki(55)"
            Tab(3).Control(8).Enabled=   0   'False
            Tab(3).Control(9)=   "chkLogGouki(54)"
            Tab(3).Control(9).Enabled=   0   'False
            Tab(3).Control(10)=   "chkLogGouki(53)"
            Tab(3).Control(10).Enabled=   0   'False
            Tab(3).Control(11)=   "chkLogGouki(52)"
            Tab(3).Control(11).Enabled=   0   'False
            Tab(3).Control(12)=   "chkLogGouki(51)"
            Tab(3).Control(12).Enabled=   0   'False
            Tab(3).Control(13)=   "chkLogGouki(50)"
            Tab(3).Control(13).Enabled=   0   'False
            Tab(3).Control(14)=   "chkLogGouki(49)"
            Tab(3).Control(14).Enabled=   0   'False
            Tab(3).Control(15)=   "chkLogGouki(48)"
            Tab(3).Control(15).Enabled=   0   'False
            Tab(3).ControlCount=   16
            TabCaption(4)   =   "  "
            TabPicture(4)   =   "���O�Ǘ�(�����)���.frx":00C4
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "chkLogGouki(79)"
            Tab(4).Control(0).Enabled=   0   'False
            Tab(4).Control(1)=   "chkLogGouki(78)"
            Tab(4).Control(1).Enabled=   0   'False
            Tab(4).Control(2)=   "chkLogGouki(77)"
            Tab(4).Control(2).Enabled=   0   'False
            Tab(4).Control(3)=   "chkLogGouki(76)"
            Tab(4).Control(3).Enabled=   0   'False
            Tab(4).Control(4)=   "chkLogGouki(75)"
            Tab(4).Control(4).Enabled=   0   'False
            Tab(4).Control(5)=   "chkLogGouki(74)"
            Tab(4).Control(5).Enabled=   0   'False
            Tab(4).Control(6)=   "chkLogGouki(73)"
            Tab(4).Control(6).Enabled=   0   'False
            Tab(4).Control(7)=   "chkLogGouki(72)"
            Tab(4).Control(7).Enabled=   0   'False
            Tab(4).Control(8)=   "chkLogGouki(71)"
            Tab(4).Control(8).Enabled=   0   'False
            Tab(4).Control(9)=   "chkLogGouki(70)"
            Tab(4).Control(9).Enabled=   0   'False
            Tab(4).Control(10)=   "chkLogGouki(69)"
            Tab(4).Control(10).Enabled=   0   'False
            Tab(4).Control(11)=   "chkLogGouki(68)"
            Tab(4).Control(11).Enabled=   0   'False
            Tab(4).Control(12)=   "chkLogGouki(67)"
            Tab(4).Control(12).Enabled=   0   'False
            Tab(4).Control(13)=   "chkLogGouki(66)"
            Tab(4).Control(13).Enabled=   0   'False
            Tab(4).Control(14)=   "chkLogGouki(65)"
            Tab(4).Control(14).Enabled=   0   'False
            Tab(4).Control(15)=   "chkLogGouki(64)"
            Tab(4).Control(15).Enabled=   0   'False
            Tab(4).ControlCount=   16
            TabCaption(5)   =   "  "
            TabPicture(5)   =   "���O�Ǘ�(�����)���.frx":00E0
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "chkLogGouki(95)"
            Tab(5).Control(0).Enabled=   0   'False
            Tab(5).Control(1)=   "chkLogGouki(94)"
            Tab(5).Control(1).Enabled=   0   'False
            Tab(5).Control(2)=   "chkLogGouki(93)"
            Tab(5).Control(2).Enabled=   0   'False
            Tab(5).Control(3)=   "chkLogGouki(92)"
            Tab(5).Control(3).Enabled=   0   'False
            Tab(5).Control(4)=   "chkLogGouki(91)"
            Tab(5).Control(4).Enabled=   0   'False
            Tab(5).Control(5)=   "chkLogGouki(90)"
            Tab(5).Control(5).Enabled=   0   'False
            Tab(5).Control(6)=   "chkLogGouki(89)"
            Tab(5).Control(6).Enabled=   0   'False
            Tab(5).Control(7)=   "chkLogGouki(88)"
            Tab(5).Control(7).Enabled=   0   'False
            Tab(5).Control(8)=   "chkLogGouki(87)"
            Tab(5).Control(8).Enabled=   0   'False
            Tab(5).Control(9)=   "chkLogGouki(86)"
            Tab(5).Control(9).Enabled=   0   'False
            Tab(5).Control(10)=   "chkLogGouki(85)"
            Tab(5).Control(10).Enabled=   0   'False
            Tab(5).Control(11)=   "chkLogGouki(84)"
            Tab(5).Control(11).Enabled=   0   'False
            Tab(5).Control(12)=   "chkLogGouki(83)"
            Tab(5).Control(12).Enabled=   0   'False
            Tab(5).Control(13)=   "chkLogGouki(82)"
            Tab(5).Control(13).Enabled=   0   'False
            Tab(5).Control(14)=   "chkLogGouki(81)"
            Tab(5).Control(14).Enabled=   0   'False
            Tab(5).Control(15)=   "chkLogGouki(80)"
            Tab(5).Control(15).Enabled=   0   'False
            Tab(5).ControlCount=   16
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   95
               Left            =   -68640
               TabIndex        =   197
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   94
               Left            =   -70800
               TabIndex        =   196
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   93
               Left            =   -72840
               TabIndex        =   195
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   92
               Left            =   -74880
               TabIndex        =   194
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   91
               Left            =   -68640
               TabIndex        =   193
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   90
               Left            =   -70800
               TabIndex        =   192
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   89
               Left            =   -72840
               TabIndex        =   191
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   88
               Left            =   -74880
               TabIndex        =   190
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   87
               Left            =   -68640
               TabIndex        =   189
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   86
               Left            =   -70800
               TabIndex        =   188
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   85
               Left            =   -72840
               TabIndex        =   187
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   84
               Left            =   -74880
               TabIndex        =   186
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   83
               Left            =   -68640
               TabIndex        =   185
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   82
               Left            =   -70800
               TabIndex        =   184
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   81
               Left            =   -72840
               TabIndex        =   183
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   80
               Left            =   -74880
               TabIndex        =   182
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   79
               Left            =   -68640
               TabIndex        =   181
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   78
               Left            =   -70800
               TabIndex        =   180
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   77
               Left            =   -72840
               TabIndex        =   179
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   76
               Left            =   -74880
               TabIndex        =   178
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   75
               Left            =   -68640
               TabIndex        =   177
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   74
               Left            =   -70800
               TabIndex        =   176
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   73
               Left            =   -72840
               TabIndex        =   175
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   72
               Left            =   -74880
               TabIndex        =   174
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   71
               Left            =   -68640
               TabIndex        =   173
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   70
               Left            =   -70800
               TabIndex        =   172
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   69
               Left            =   -72840
               TabIndex        =   171
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   68
               Left            =   -74880
               TabIndex        =   170
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   67
               Left            =   -68640
               TabIndex        =   169
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   66
               Left            =   -70800
               TabIndex        =   168
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   65
               Left            =   -72840
               TabIndex        =   167
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   64
               Left            =   -74880
               TabIndex        =   166
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   63
               Left            =   -68640
               TabIndex        =   165
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   62
               Left            =   -70800
               TabIndex        =   164
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   61
               Left            =   -72840
               TabIndex        =   163
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   60
               Left            =   -74880
               TabIndex        =   162
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   59
               Left            =   -68640
               TabIndex        =   161
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   58
               Left            =   -70800
               TabIndex        =   160
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   57
               Left            =   -72840
               TabIndex        =   159
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   56
               Left            =   -74880
               TabIndex        =   158
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   55
               Left            =   -68640
               TabIndex        =   157
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   54
               Left            =   -70800
               TabIndex        =   156
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   53
               Left            =   -72840
               TabIndex        =   155
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   52
               Left            =   -74880
               TabIndex        =   154
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   51
               Left            =   -68640
               TabIndex        =   153
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   50
               Left            =   -70800
               TabIndex        =   152
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   49
               Left            =   -72840
               TabIndex        =   151
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   48
               Left            =   -74880
               TabIndex        =   150
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   47
               Left            =   -68640
               TabIndex        =   149
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   46
               Left            =   -70800
               TabIndex        =   148
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   45
               Left            =   -72840
               TabIndex        =   147
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   44
               Left            =   -74880
               TabIndex        =   146
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   43
               Left            =   -68640
               TabIndex        =   145
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   42
               Left            =   -70800
               TabIndex        =   144
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   41
               Left            =   -72840
               TabIndex        =   143
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   40
               Left            =   -74880
               TabIndex        =   142
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   39
               Left            =   -68640
               TabIndex        =   141
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   38
               Left            =   -70800
               TabIndex        =   140
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   37
               Left            =   -72840
               TabIndex        =   139
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   36
               Left            =   -74880
               TabIndex        =   138
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   35
               Left            =   -68640
               TabIndex        =   137
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   34
               Left            =   -70800
               TabIndex        =   136
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   33
               Left            =   -72840
               TabIndex        =   135
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   32
               Left            =   -74880
               TabIndex        =   134
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   31
               Left            =   -68640
               TabIndex        =   133
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   30
               Left            =   -70800
               TabIndex        =   132
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   29
               Left            =   -72840
               TabIndex        =   131
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   28
               Left            =   -74880
               TabIndex        =   130
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   27
               Left            =   -68640
               TabIndex        =   129
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   26
               Left            =   -70800
               TabIndex        =   128
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   25
               Left            =   -72840
               TabIndex        =   127
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   24
               Left            =   -74880
               TabIndex        =   126
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   23
               Left            =   -68640
               TabIndex        =   125
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   22
               Left            =   -70800
               TabIndex        =   124
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   21
               Left            =   -72840
               TabIndex        =   123
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   20
               Left            =   -74880
               TabIndex        =   122
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   19
               Left            =   -68640
               TabIndex        =   121
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   18
               Left            =   -70800
               TabIndex        =   120
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   17
               Left            =   -72840
               TabIndex        =   119
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "�P�Q�R�S�T���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   16
               Left            =   -74880
               TabIndex        =   118
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   15
               Left            =   6360
               TabIndex        =   117
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   4200
               TabIndex        =   116
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   13
               Left            =   2160
               TabIndex        =   115
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   114
               Top             =   2040
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   11
               Left            =   6360
               TabIndex        =   113
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   10
               Left            =   4200
               TabIndex        =   112
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   9
               Left            =   2160
               TabIndex        =   111
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   110
               Top             =   1560
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   6360
               TabIndex        =   109
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   4200
               TabIndex        =   108
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   2160
               TabIndex        =   107
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   106
               Top             =   1080
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   6360
               TabIndex        =   105
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   4200
               TabIndex        =   104
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   2160
               TabIndex        =   103
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkLogGouki 
               Caption         =   "Z9���@"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   102
               Top             =   600
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1815
            End
         End
      End
      Begin TabDlg.SSTab tabTakuCorner 
         Height          =   7815
         Left            =   -74640
         TabIndex        =   202
         Top             =   600
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   13785
         _Version        =   393216
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   794
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   " ������������ ������������"
         TabPicture(0)   =   "���O�Ǘ�(�����)���.frx":00FC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraLogFile(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   " ������������ ������������"
         TabPicture(1)   =   "���O�Ǘ�(�����)���.frx":0118
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1(1)"
         Tab(1).Control(1)=   "fraLogFile(1)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   " ������������ ������������"
         TabPicture(2)   =   "���O�Ǘ�(�����)���.frx":0134
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame1(2)"
         Tab(2).Control(1)=   "fraLogFile(2)"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   " ������������ ������������"
         TabPicture(3)   =   "���O�Ǘ�(�����)���.frx":0150
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame1(3)"
         Tab(3).Control(1)=   "fraLogFile(3)"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   " ������������ ������������"
         TabPicture(4)   =   "���O�Ǘ�(�����)���.frx":016C
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame1(4)"
         Tab(4).Control(1)=   "fraLogFile(4)"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   " ������������ ������������"
         TabPicture(5)   =   "���O�Ǘ�(�����)���.frx":0188
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame1(5)"
         Tab(5).Control(1)=   "fraLogFile(5)"
         Tab(5).ControlCount=   2
         Begin VB.Frame Frame1 
            BorderStyle     =   0  '�Ȃ�
            Caption         =   "Frame1"
            Height          =   975
            Index           =   5
            Left            =   -74760
            TabIndex        =   258
            Top             =   600
            Width           =   5775
            Begin VB.OptionButton optApp 
               Caption         =   "�A�v���P�[�V�������O"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   5
               Left            =   120
               TabIndex        =   260
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optHoshu 
               Caption         =   "��ʑ��샍�O"
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
               Index           =   5
               Left            =   120
               TabIndex        =   259
               Top             =   480
               Width           =   2655
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  '�Ȃ�
            Caption         =   "Frame1"
            Height          =   975
            Index           =   4
            Left            =   -74760
            TabIndex        =   255
            Top             =   600
            Width           =   5775
            Begin VB.OptionButton optApp 
               Caption         =   "�A�v���P�[�V�������O"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   257
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optHoshu 
               Caption         =   "��ʑ��샍�O"
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
               Left            =   120
               TabIndex        =   256
               Top             =   480
               Width           =   2655
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  '�Ȃ�
            Caption         =   "Frame1"
            Height          =   975
            Index           =   3
            Left            =   -74760
            TabIndex        =   252
            Top             =   600
            Width           =   5775
            Begin VB.OptionButton optApp 
               Caption         =   "�A�v���P�[�V�������O"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   254
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optHoshu 
               Caption         =   "��ʑ��샍�O"
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
               Left            =   120
               TabIndex        =   253
               Top             =   480
               Width           =   2655
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  '�Ȃ�
            Caption         =   "Frame1"
            Height          =   975
            Index           =   2
            Left            =   -74760
            TabIndex        =   249
            Top             =   600
            Width           =   5775
            Begin VB.OptionButton optApp 
               Caption         =   "�A�v���P�[�V�������O"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   251
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optHoshu 
               Caption         =   "��ʑ��샍�O"
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
               Left            =   120
               TabIndex        =   250
               Top             =   480
               Width           =   2655
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  '�Ȃ�
            Caption         =   "Frame1"
            Height          =   975
            Index           =   1
            Left            =   -74760
            TabIndex        =   246
            Top             =   600
            Width           =   5775
            Begin VB.OptionButton optApp 
               Caption         =   "�A�v���P�[�V�������O"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   248
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optHoshu 
               Caption         =   "��ʑ��샍�O"
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
               Index           =   1
               Left            =   120
               TabIndex        =   247
               Top             =   480
               Width           =   2655
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  '�Ȃ�
            Caption         =   "Frame1"
            Height          =   975
            Index           =   0
            Left            =   240
            TabIndex        =   243
            Top             =   600
            Width           =   5775
            Begin VB.OptionButton optHoshu 
               Caption         =   "��ʑ��샍�O"
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
               Left            =   120
               TabIndex        =   245
               Top             =   480
               Width           =   2655
            End
            Begin VB.OptionButton optApp 
               Caption         =   "�A�v���P�[�V�������O"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   244
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
         End
         Begin VB.Frame fraLogFile 
            Caption         =   "�Ď��Ճ��O�t�@�C��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Index           =   0
            Left            =   120
            TabIndex        =   233
            Top             =   1560
            Width           =   5895
            Begin VB.ListBox lstLogFile 
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5340
               Index           =   0
               Left            =   120
               MultiSelect     =   2  '�g��
               TabIndex        =   234
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�t�@�C����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   238
               Top             =   360
               Width           =   1680
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�ŏI�����N����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   1800
               TabIndex        =   237
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   " ���F��"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   3530
               TabIndex        =   236
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�T�C�Y "
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   4500
               TabIndex        =   235
               Top             =   360
               Width           =   1005
            End
         End
         Begin VB.Frame fraLogFile 
            Caption         =   "�Ď��Ճ��O�t�@�C��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Index           =   1
            Left            =   -74880
            TabIndex        =   227
            Top             =   1560
            Width           =   5895
            Begin VB.ListBox lstLogFile 
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5340
               Index           =   1
               Left            =   120
               MultiSelect     =   2  '�g��
               TabIndex        =   228
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�T�C�Y "
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   4500
               TabIndex        =   232
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   " ���F��"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   3530
               TabIndex        =   231
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�ŏI�����N����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   6
               Left            =   1800
               TabIndex        =   230
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�t�@�C����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   7
               Left            =   120
               TabIndex        =   229
               Top             =   360
               Width           =   1680
            End
         End
         Begin VB.Frame fraLogFile 
            Caption         =   "�Ď��Ճ��O�t�@�C��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Index           =   2
            Left            =   -74880
            TabIndex        =   221
            Top             =   1560
            Width           =   5895
            Begin VB.ListBox lstLogFile 
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5340
               Index           =   2
               Left            =   120
               MultiSelect     =   2  '�g��
               TabIndex        =   222
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�T�C�Y "
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   8
               Left            =   4500
               TabIndex        =   226
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   " ���F��"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   9
               Left            =   3530
               TabIndex        =   225
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�ŏI�����N����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   10
               Left            =   1800
               TabIndex        =   224
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�t�@�C����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   11
               Left            =   120
               TabIndex        =   223
               Top             =   360
               Width           =   1680
            End
         End
         Begin VB.Frame fraLogFile 
            Caption         =   "�Ď��Ճ��O�t�@�C��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Index           =   3
            Left            =   -74880
            TabIndex        =   215
            Top             =   1560
            Width           =   5895
            Begin VB.ListBox lstLogFile 
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5340
               Index           =   3
               Left            =   120
               MultiSelect     =   2  '�g��
               TabIndex        =   216
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�T�C�Y "
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   12
               Left            =   4500
               TabIndex        =   220
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   " ���F��"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   13
               Left            =   3530
               TabIndex        =   219
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�ŏI�����N����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   14
               Left            =   1800
               TabIndex        =   218
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�t�@�C����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   15
               Left            =   120
               TabIndex        =   217
               Top             =   360
               Width           =   1680
            End
         End
         Begin VB.Frame fraLogFile 
            Caption         =   "�Ď��Ճ��O�t�@�C��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Index           =   4
            Left            =   -74880
            TabIndex        =   209
            Top             =   1560
            Width           =   5895
            Begin VB.ListBox lstLogFile 
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5340
               Index           =   4
               Left            =   120
               MultiSelect     =   2  '�g��
               TabIndex        =   210
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�T�C�Y "
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   16
               Left            =   4500
               TabIndex        =   214
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   " ���F��"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   17
               Left            =   3530
               TabIndex        =   213
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�ŏI�����N����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   18
               Left            =   1800
               TabIndex        =   212
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�t�@�C����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   19
               Left            =   120
               TabIndex        =   211
               Top             =   360
               Width           =   1680
            End
         End
         Begin VB.Frame fraLogFile 
            Caption         =   "�Ď��Ճ��O�t�@�C��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6180
            Index           =   5
            Left            =   -74880
            TabIndex        =   203
            Top             =   1560
            Width           =   5895
            Begin VB.ListBox lstLogFile 
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5340
               Index           =   5
               Left            =   120
               MultiSelect     =   2  '�g��
               TabIndex        =   204
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�T�C�Y "
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   20
               Left            =   4500
               TabIndex        =   208
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   " ���F��"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   21
               Left            =   3530
               TabIndex        =   207
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�ŏI�����N����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   22
               Left            =   1800
               TabIndex        =   206
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label lblFile 
               Alignment       =   2  '��������
               BorderStyle     =   1  '����
               Caption         =   "�t�@�C����"
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   23
               Left            =   120
               TabIndex        =   205
               Top             =   360
               Width           =   1680
            End
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "����샍�O�Ǘ�"
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
      TabIndex        =   37
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmTakuLogKanri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmTakuLogKanri.frm
'//  �p�b�P�[�W���F����샍�O�Ǘ����
'//
'//  �T�v�F����샍�O�Ǘ����
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-23   CODED   BY [TCC] M.Matsumoto
'//                 �y�t�F�[�Y�Q�Ή��z����샍�O�Ǘ���ʂ𗬗p���ĐV�K�쐬
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-13   CODED   BY [TCC] M.Matsumoto
'//                 �y��-350�Ή��z
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.115�C���Ή��z
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05  CODED BY  [TCC] H.Sugimoto
'//                 �y���k�t�H���_�w��Ή��z
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ���O�}�̏o�͎��A������T�P�Q���Ƃ���
'//     REVISIONS :(EG20V5.10.0.1) 2012-05-09 REVISED BY [TCC] M.Chiwaki
'//                 ���O�}�̏o�͎��A�t�H���_���쐬����
'//     REVISIONS :(X.X.X.X) 0000-00-00   CODED   BY [ ]
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

'*****************************************************************************
'*      �萔
'*****************************************************************************
Private Const MN_COLOR_BLACK = &H80000008
Private Const MN_COLOR_RED = &HFF&
Private Const MN_COLOR_WHITE = &H80000005
Private Const MN_COLOR_YELLOW = &HFFFF&

Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l

'*****************************************************************************
'*      ���O���i�[�G���A
'*****************************************************************************
Private Type LogFileData
    sPath As String                 '���O�t�@�C���̃p�X
    sName As String                 '���O�t�@�C����
    dtFileDate As Date              '�쐬���t�E����
    lFileSize As Long               '�t�@�C���T�C�Y
    bSelect As Boolean              '�I���t���O
End Type

Private uLogfileData() As LogFileData
'*****************************************************************************
'*      �Ώۃt�@�C���t���p�X�i����̧�ق̎��A��߰�1�����ŋ�؂�B�j
'*****************************************************************************
Private sObjectFiles As String   '۸�̧��ؽ��ޯ���őI�𒆂�̧�ق����߽������
Private sObjectTopFile As String '����A�I�𒆂̐擪�i�ŋ��j̧�ٖ��B

'*****************************************************************************
'*      �C�x���g���O�R�s�[�p���[�N�t�@�C�����t���p�X
'*****************************************************************************
Private Const SAVEFILE_SYS As String = PATH_WORK & "SysEvent.Evt"
Private Const SAVEFILE_SEC As String = PATH_WORK & "ScuEvent.Evt"
Private Const SAVEFILE_APP As String = PATH_WORK & "AppEvent.Evt"

'���k�t�@�C���p
Private Type files
    sFileName(255) As String
End Type

'EG20 V3.6.0.1�y03����TR-No.115�C���Ή��z�폜�J�n
''EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
'Private Const CAB_LOG_FILE As String = PATH_WORK & "KANSI_LOG_TMP.CAB"
'Private Const DAT_LOG_FILE As String = PATH_WORK & "KANSI_LOG_TMP.DAT"
''EG20 V2.1.0.1 ADD END   �y�t�F�[�Y�Q�Ή��z
'EG20 V3.6.0.1�y03����TR-No.115�C���Ή��z�폜�I��
'EG20 V3.6.0.1�y03����TR-No.115�C���Ή��z�ǉ��J�n
Private Const CAB_LOG_FILE As String = PATH_WORK & "KLOGTEMP.CAB"
Private Const DAT_LOG_FILE As String = PATH_WORK & "KLOGTEMP.DAT"
'EG20 V3.6.0.1�y03����TR-No.115�C���Ή��z�ǉ��I��

'*****************************************************************************
'*      ���W���[�����i�[�G���A
'*****************************************************************************
Private Type ModFileData
    sName As String             '�v���Z�X��
    iProces As Integer          '�v���Z�XID
    iFuzokuId As Integer        '�t���v���Z�XID
    iFuzokuCnt As Integer       '�t���J�E���^
End Type
Private uModFileData(59) As ModFileData
Private iModCnt As Integer

Private Const ASRT_LOG = &H200         ' 10:���O��g���[�X
Private Const ASRT_HOSYU = &H400       ' 11:�ێ��ʐݒ�
Private Const ASRT_SYUKEI = &H800      ' 12:�W�v                'REV(03.00)�s�ǉ��B
Private Const ASRT_ALL = &H7FFFFFFF    '�S���ރ��O���W

'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
Private Const PATH_LOG_CORNER = "E:\\KANSI\\CORNER"
Private Const DIR_LOG_APL = "\\OPERATE_APL_LOG\\"
Private Const DIR_LOG_SOUSA = "\\OPERATE_SOUSA_LOG\\"

Private mintStatus(31) As Integer
'EG20 V2.1.0.1 ADD END


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
Private Sub cmdInstall_Click()
    On Error Resume Next
  
   '�u�}�̎�O�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '�}�̎�O����
    Call pfRemove(Me)
End Sub

Private Sub cmdLogShushu_Click()

    Dim intTabIdx As Integer
    Dim iResponse As Integer
    
    intTabIdx = tabTakuCorner.Tab

    '��ʑ��샍�O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KAKARISET_GAMEN_SAVE, 0)
    
    '�m�F���b�Z�[�W�{�b�N�X��\������B
    iResponse = MsgBox("����샍�O�f�[�^�����W���܂�����낵���ł����H", _
                        vbOKCancel, "���W")

    '�u�L�����Z���v�{�^�����������͏������I������
    If iResponse = vbCancel Then Exit Sub
    
    '�A�v�����O�^��ʑ��샍�O��ʂ�ݒ�
    If optApp(intTabIdx).Value = True Then
        glnglogKind = LOG_COL_KIND.LOG_APP
    Else
        glnglogKind = LOG_COL_KIND.LOG_SOUSA
    End If
    '�ΏۃR�[�i��ݒ�
    glngTargetCorner = intTabIdx + 1
    
    '��������ʂ�\������
    dlgLogShushuMessage.Show vbModal

    '���O�ꗗ�ĕ\��
    Call sSetListBox(intTabIdx)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : ����샍�O�Ǘ����(�A�N�e�B�u��)
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
    Dim bRet As Boolean                 '�߂�l
    Dim lId As Long                     '���[���h�c
    Dim bFlag As Boolean                '�t���O
    Dim lngErrCode As Long              '�G���[�R�[�h
    Dim udtMail As ML_KYOTU_INF         '�o�b�t�@�t���b�V���v��

    On Error Resume Next
    
    tmrMail.Enabled = True
            
   '�o�b�t�@�t���b�V���v�������O�v���Z�X�ɑ��M����
    udtMail.udtlHeader.dwId = ML_ID_LGBUFF_REQ
    udtMail.udtlHeader.dwSize = MlSize.BUFF_FLUSH_CMD
    udtMail.udtlHeader.dwProid = RHOSHU_ID
    udtMail.udtlHeader.dwSubArea = 0
    bRet = DssSendMail(MAIL_SLOT_LOG, MlSize.BUFF_FLUSH_CMD, udtMail.udtlHeader)
    If bRet = False Then
       '�u�o�b�t�@�t���b�V���v�����M�ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSI_LOG_BUFF_FLUSH_CMD_SEND, lngErrCode)
       Exit Sub
    Else
       '�u�o�b�t�@�t���b�V���v�����M����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSI_LOG_BUFF_FLUSH_CMD_SEND, 0)
    End If
  
    '�o�b�t�@�t���b�V���I���ʒm��M
    bFlag = False
    Do Until bFlag = True
        '���[����M�������s��
        lId = fMailRecieve()
        Select Case lId         '���[���h�c
        '�u�v���Z�X�I���w���v�̏ꍇ
        Case ML_ID_PROEND_ORD
             '�u�v���Z�X�I���w����M����v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
            '�������I������
            Exit Sub
        '�u�o�b�t�@�t���b�V���I���ʒm�v�̏ꍇ
        Case ML_ID_LGBUFF_ANS
            '�u�o�b�t�@�t���b�V���I���ʒm��M����v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, KANSI_LOG_BUFF_FLUSH_END_RECV, 0)
            '���[�v�𔲂���
            Exit Do
        Case Else
        End Select
        Sleep (MN_MAIL_INTERVAL)
    Loop
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : ����샍�O�Ǘ����(�f�B�A�N�e�B�u��)
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
Private Sub Form_Deactivate()
   On Error Resume Next
   
   If blnCabfrmOpenFlg = True Then
        Call fnTsbCabCallDiverge
        Exit Sub
    End If
    '���[����M�^�C�}���~����B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : ����샍�O�Ǘ����(���[�h��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
'//                 �y�t�F�[�Y�Q�Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim iRet As Integer             '�֐��̖߂�l
    Dim sKeyName As String          'INI�t�@�C���L�[��
    Dim iMozi As Integer            '�ǂݍ��ݕ�����
    Dim iKbn As Integer             '�ǂݍ��񂾕�����
    Dim sIni_Data As String * 128   'INI�t�@�C�����1�s���擾
    Dim iCnt As Integer             'INI�t�@�C���J�E���^
    Dim i As Integer                '�J�E���^
    Dim j As Integer                '�R���g���[���z��
    Dim MyName As String            'INI�L���`�F�b�N
    'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�z
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intIndex As Integer
    Dim strCorner1 As String
    Dim strCorner2 As String
    Dim bySyoAssort As Byte             '���O�p������
    'EG20 V2.1.0.1 ADD END

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�z
    '���@���擾
    Call gsGetGateInfo
    Call gsGetCornerName
    
    '�^�u����ݒu�R�[�i���Ƃ���
    tabCorner.Tab = 0
    
    '���W��ԏ�����
    Erase mintStatus
    
    '�����t�@�C���G���[�̃g���b�v
    On Error GoTo OtherError
    bySyoAssort = L3AN_FILE
    
    For intCount = 0 To UBound(gblnCornerSet)
    
        '�ݒ肠��̃R�[�i�������ɂ���
        If gblnCornerSet(intCount) = True Then
            '�R�[�i�[���̕\��
            strCorner1 = MidB(gstrCornerName(intCount), 1, 12)
            strCorner2 = MidB(gstrCornerName(intCount), 13, 24)
'            tabCorner.TabCaption(intCount) = strCorner1 & vbCrLf & strCorner2
            tabCorner.TabCaption(intCount) = Empty
            tabTakuCorner.TabCaption(intCount) = strCorner1 & vbCrLf & strCorner2
        End If
    
    Next intCount
    
    '�ݒu�R�[�i�������[�v
    For intCount = 0 To UBound(gblnCornerSet)

        If gblnCornerSet(intCount) = False Then
            tabCorner.TabVisible(intCount) = False
            tabTakuCorner.TabVisible(intCount) = False
            optApp(intCount).Value = True
        End If

        '�ő卆�@�������[�v
        For intCount2 = 0 To 15
            intIndex = (intCount * 16) + intCount2
            chkLogGouki(intIndex).Visible = False
            chkLogGouki(intIndex).Tag = "0"
        Next
        
        For intCount2 = 0 To 15
            intIndex = (intCount * 16) + (gudtSettiCorner(intCount).intGokiNo(intCount2) - 1)
            If gudtSettiCorner(intCount).intGokiNo(intCount2) > 0 Then
                chkLogGouki(intIndex).Caption = gudtSettiCorner(intCount).strDispGoki(intCount2) + "���@"
                'Tag�ɑΉ����鍆�@�ԍ����L�^�i1�`32���@�j
                chkLogGouki(intIndex).Tag = CStr(gudtSettiCorner(intCount).intGateNo(intCount2))
                mintStatus(gudtSettiCorner(intCount).intGateNo(intCount2) - 1) = CHECKBOX_ON
                chkLogGouki(intIndex).Visible = True
                chkLogGouki(intIndex).Value = CHECKBOX_ON
            End If
        Next intCount2
        
    Next intCount
    'EG20 V2.1.0.1 ADD END
    
    '�\���t�@�C���w���o�^����
'    sSetListBox            'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
    'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
    For intCount = 0 To 5
        Call sSetListBox(intCount)
    Next
    'EG20 V2.1.0.1 ADD END
    
    '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
     
    '�t�@�C���L���`�F�b�N
'    MyName = Dir(DISP_FILE, vbNormal)          'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
    MyName = Dir(DISP_FILE_TAKU, vbNormal)      'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
    If MyName = "" Then
        GoTo FileError
    End If
    
    For iCnt = 0 To 59
        sKeyName = DISP_KEY_NAME & Format(iCnt, "00")
        'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
'        iRet = GetPrivateProfileString(DISP_SECTION_NAME, _
'                                       sKeyName, _
'                                       DEFAILT, sIni_Data, Len(sIni_Data), _
'                                       DISP_FILE)
        'EG20 V2.1.0.1 DEL END
        'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
        iRet = GetPrivateProfileString(DISP_SECTION_NAME, _
                                       sKeyName, _
                                       DEFAILT, sIni_Data, Len(sIni_Data), _
                                       DISP_FILE_TAKU)
        'EG20 V2.1.0.1 ADD END
        iMozi = 1
        iKbn = 1
        Do
           '���W���[�����i�[�G���A��1�s���̃f�[�^��ێ�������B
            If Mid(sIni_Data, iMozi, 1) = "," Then
                Select Case iKbn
                    Case 1
                        uModFileData(iCnt).sName = Left(sIni_Data, iMozi - 1)
                        sIni_Data = Mid(sIni_Data, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 2
                        uModFileData(iCnt).iProces = Left(sIni_Data, iMozi - 1)
                        sIni_Data = Mid(sIni_Data, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 3
                        uModFileData(iCnt).iFuzokuId = Left(sIni_Data, iMozi - 1)
                        sIni_Data = Mid(sIni_Data, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 4
                        uModFileData(iCnt).iFuzokuCnt = Left(sIni_Data, iMozi - 1)
                        sIni_Data = Mid(sIni_Data, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                End Select
            End If
            iMozi = iMozi + 1
            If iMozi > Len(sIni_Data) Then
                Exit Do
            End If
        Loop
        
        '1�s���f�[�^�̕ێ�������A�\���������s���B
        If iKbn = 4 Then
            chkMod(iCnt).Visible = True
            chkMod(iCnt).Caption = uModFileData(iCnt).sName
           If uModFileData(iCnt).iFuzokuId = 0 Then
              Select Case iCnt
                Case 0 To 19
                  '�啪�ވ����F���ރJ�E���^�[0�`19�̏ꍇ
                  chkMod(iCnt).Left = 120
                Case 20 To 39
                  '�啪�ވ����F���ރJ�E���^�[20�`39�̏ꍇ
                  chkMod(iCnt).Left = 2295
                Case 40 To 59
                  '�啪�ވ����F���ރJ�E���^�[40�`59�̏ꍇ
                  chkMod(iCnt).Left = 4470
              End Select
          Else
              Select Case iCnt
                Case 0 To 19
                '�����ވ����F���ރJ�E���^�[0�`19�̏ꍇ
                  chkMod(iCnt).Left = 330
                Case 20 To 39
                '�����ވ����F���ރJ�E���^�[20�`39�̏ꍇ
                  chkMod(iCnt).Left = 2500
                Case 40 To 59
                '�啪�ވ����F���ރJ�E���^�[40�`59�̏ꍇ
                  chkMod(iCnt).Left = 4670
             End Select
          End If
            iModCnt = iCnt
        End If
    Next
          
   '�\�����ڎw�������������
    optLogSyu(0).Value = True               '���W�I�t�F�u�S�Ă̎�ʂ�\���v��L��
    j = chkLogSyu.UBound
    For i = 0 To j                          '��ʕ��J��Ԃ�
        chkLogSyu(i).Value = CHECKBOX_ON    '�u�H�H��ʁv��L���ɂ���
    Next
    
    optLogBunrui(0).Value = True            '���W�I�t�F�u�S�Ă̕��ނ�\���v��L��
    For i = 0 To iModCnt                    '���ޕ��J��Ԃ�
        If chkMod(i).Visible = True Then
            chkMod(i).Value = CHECKBOX_ON   '�u�H�H���ށv��L���ɂ���
        End If
    Next

    optLogData(1).Value = True             '�u�P�s�ڂ̂ݕ\���v��L���ɂ���

    'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
'    '�\���������@�w�������������
'    optLogGouki(0).Value = True            '���W�I�t�F�u�S�����v��L��
'    cmdChkAll.Enabled = False
'    cmdChkAllKai.Enabled = False
'
'    j = chkLogGouki.UBound
'    For i = 0 To j                         '���@���J��Ԃ�
'        chkLogGouki(i).Value = CHECKBOX_ON '�u�H�H���@�v��L���ɂ���
'        chkLogGouki(i).Enabled = False     '�S���@�����s��
'    Next
    'EG20 V2.1.0.1 DEL END
   
   tabLog.Tab = 0
   tabTakuCorner.Tab = 0        'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
   
   Call tabTakuCorner_Click(0)  'EG20 V2.1.0.1 ADD �y��-350�Ή��z
   
   '�u����샍�O�Ǘ���ʁF�\���v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_LOG_TAKU_GAMEN_START, 0)
   
   Exit Sub

FileError:
   '�u����샍�O�Ǘ��FINI�t�@�C���ُ�v���O�o��
   Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
 
   'INI�t�@�C���L���`�F�b�N�ُ펞�F�u�t�@�C���ُ�v�|�b�v�A�b�v��\��
   MsgBox "INI�t�@�C���̎擾�Ɏ��s���܂����", vbCritical, "�t�@�C���ُ�"
   Exit Sub
   
OtherError:
  '�u����샍�O�Ǘ��F���O�\���ُ�v���O�o��
   Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, KANSI_LOG_KANRI_LOG_ERROR, 0)
  '���X�g�{�b�N�X�̏�����
'   lstLogFile.Clear        'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
  'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
    For i = 0 To lstLogFile.UBound
        lstLogFile(i).Clear
    Next
  'EG20 V2.1.0.1 ADD END
   MsgBox "���O�ꗗ�̎擾�Ɏ��s���܂����B", vbCritical, "�\���ُ�"
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSetListBox
'//  �@�\����  : ���O�t�@�C���o�^����
'//  �@�\�T�v  : ���O�t�@�C�������X�g�{�b�N�X�ɓo�^����B
'//              �\���t�@�C���w�蕔�F��������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Private Sub sSetListBox()                      'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
Private Sub sSetListBox(intTabIdx As Integer)   'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
    Dim i As Integer            '�J�E���^
    Dim j As Integer            '�J�E���^
    Dim iCnt As Integer         '���O�t�@�C����
    Dim sEntry As String        '�ҏW������
    Dim uLogData As LogFileData '�o�[�W�������o�b�t�@

    On Error Resume Next
    
    '���O�t�@�C�������擾����
'    iCnt = fGetLogfileInf()            'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
    iCnt = fGetLogfileInf(intTabIdx)    'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z

    '�t�@�C�����Ń\�[�g����
    For i = 0 To iCnt - 2
        For j = i + 1 To iCnt - 1
            'If uLogfileData(i).sName > uLogfileData(j).sName Then              'V1.7.0.1 DEL
            If UCase(uLogfileData(i).sName) > UCase(uLogfileData(j).sName) Then 'V1.7.0.1 ADD
                uLogData = uLogfileData(i)
                uLogfileData(i) = uLogfileData(j)
                uLogfileData(j) = uLogData
            End If
        Next j
    Next i

    '�u���O�t�@�C���v���X�g�{�b�N�X���N���A����
'    lstLogFile.Clear               'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
    lstLogFile(intTabIdx).Clear     'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z

    '���O�t�@�C������ҏW����
    For i = 0 To iCnt - 1       '���O�t�@�C�������J��Ԃ�
        sEntry = Mid$(uLogfileData(i).sName & Space(14), 1, 14)
        sEntry = sEntry & "    " & Format(uLogfileData(i).dtFileDate, "yyyy/mm/dd  hh:nn")
        sEntry = sEntry & Format(uLogfileData(i).lFileSize, "@@@@@@@@@")
'        lstLogFile.AddItem sEntry       '���X�g�{�b�N�X�ɒǉ�����              'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
        lstLogFile(intTabIdx).AddItem sEntry       '���X�g�{�b�N�X�ɒǉ�����    'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
    Next
    If iCnt > 0 Then                    '���O�t�@�C�������݂���
        'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
'        lstLogFile.ListIndex = 0        '��s�ڂɃC���f�b�N�X���Z�b�g
'        lstLogFile.Selected(0) = True   '��s�ڂ�I���ςɂ���
        'EG20 V2.1.0.1 DEL END
        'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
        lstLogFile(intTabIdx).ListIndex = 0        '��s�ڂɃC���f�b�N�X���Z�b�g
        lstLogFile(intTabIdx).Selected(0) = True   '��s�ڂ�I���ςɂ���
        'EG20 V2.1.0.1 DEL END
    End If

End Sub

'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdZHisentaku_Click
'//  �@�\����  : �S�R�[�i�S���@��I���{�^����������
'//  �@�\�T�v  : ���ׂĂ̍��@���I����Ԃɂ���
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdZHisentaku_Click()

    Dim intLoop As Integer
    
    On Error Resume Next
    
    For intLoop = 0 To chkLogGouki.UBound
        chkLogGouki(intLoop).Value = CHECKBOX_OFF
    Next intLoop
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdZSentaku_Click
'//  �@�\����  : �S�R�[�i�S���@�I���{�^����������
'//  �@�\�T�v  : ���ׂĂ̍��@��I����Ԃɂ���
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdZSentaku_Click()

    Dim intLoop As Integer
    
    On Error Resume Next
    
    For intLoop = 0 To chkLogGouki.UBound
        chkLogGouki(intLoop).Value = CHECKBOX_ON
    Next intLoop
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdHHisentaku_Click
'//  �@�\����  : �\���R�[�i�S���@��I���{�^����������
'//  �@�\�T�v  : ���ׂĂ̍��@���I����Ԃɂ���
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdHHisentaku_Click()

    Dim intLoop As Integer
    Dim intStIndex As Integer
    Dim intEdIndex As Integer
    Dim intGokiIndex As Integer
    
    On Error Resume Next

    intStIndex = tabCorner.Tab * 16
    intEdIndex = intStIndex + 15
    
    For intLoop = intStIndex To intEdIndex
        chkLogGouki(intLoop).Value = CHECKBOX_OFF
    Next intLoop
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdHSentaku_Click
'//  �@�\����  : �\���R�[�i�S���@�I���{�^����������
'//  �@�\�T�v  : ���ׂĂ̍��@��I����Ԃɂ���
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdHSentaku_Click()

    Dim intLoop As Integer
    Dim intStIndex As Integer
    Dim intEdIndex As Integer
    Dim intGokiIndex As Integer
    
    On Error Resume Next

    intStIndex = tabCorner.Tab * 16
    intEdIndex = intStIndex + 15
    
    For intLoop = intStIndex To intEdIndex
        chkLogGouki(intLoop).Value = CHECKBOX_ON
    Next intLoop
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : chkLogGouki_Click
'//  �@�\����  : �w�荆�@�I�v�V�����{�^���N���b�N������
'//  �@�\�T�v  : �����ϐ���ON/OFF��؂�ւ���
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@�@Index  �@[IN]�I�v�V�����{�^���C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-22   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub chkLogGouki_Click(Index As Integer)

    Dim intGoki As Integer
    
    On Error Resume Next
    
    intGoki = CInt(chkLogGouki(Index).Tag) - 1
    
    mintStatus(intGoki) = chkLogGouki(Index).Value
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : optApp_Click
'//  �@�\����  : ���O�敪�I�v�V�����{�^���N���b�N������
'//  �@�\�T�v  : ���O�̎�ނ�؂�ւ���
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@�@Index  �@[IN]�I�v�V�����{�^���C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-23   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub optApp_Click(Index As Integer)

    Call sSetListBox(Index)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : optHoshu_Click
'//  �@�\����  : ���O�敪�I�v�V�����{�^���N���b�N������
'//  �@�\�T�v  : ���O�̎�ނ�؂�ւ���
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@�@Index  �@[IN]�I�v�V�����{�^���C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-23   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub optHoshu_Click(Index As Integer)

    Call sSetListBox(Index)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdUpdateDisplay_Click
'//  �@�\����  : �u�\���X�V�v�t����������
'//  �@�\�T�v  : ���O�t�@�C���̕\�����X�g�̓��e���ŐV��ԂɍX�V����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-23   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdUpdateDisplay_Click()

    Dim intTabIdx As Integer
    
    intTabIdx = tabTakuCorner.Tab

    Call sSetListBox(intTabIdx)
    
End Sub
'EG20 V2.1.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fGetLogfileInf
'//  �@�\����  : ���O�t�@�C�����擾����
'//  �@�\�T�v  : �S���O�t�@�C���̏����擾����B
'//              �\���t�@�C���w�蕔�F���O�t�@�C���o�^����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Integer�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
'Public Function fGetLogfileInf() As Integer                    'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
Public Function fGetLogfileInf(intIndex As Integer) As Integer  'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
    Dim MyPath As String       '�t�H���_��
    Dim MyName As String       '�t�@�C����
    Dim iLogfileCnt As Integer '�J�E���^�[
    Dim bSelectLogSousa As Boolean                      ' ���샍�O�I����ԁiTRUE:�I���j     ' EG20 V3.0.0.2�ǉ�
    Dim bFileOK As Boolean                              ' �t�@�C����������                  ' EG20 V3.0.0.2�ǉ�

    On Error Resume Next
    
    '���O�t�@�C����������������
    iLogfileCnt = 0
    
    '�ێ��ʑ��샍�O�t�@�C������������B
    'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
    MyPath = PATH_LOG_CORNER & CStr(intIndex + 1)
    If optHoshu(intIndex).Value = True Then
        MyPath = MyPath & DIR_LOG_SOUSA
        bSelectLogSousa = True                          ' ���샍�O�I����ԁi�I���j          ' EG20 V3.0.0.2�ǉ�
    Else
        MyPath = MyPath & DIR_LOG_APL                              ' �p�X��ݒ肵�܂��B
        bSelectLogSousa = False                         ' ���샍�O�I����ԁi��I���j        ' EG20 V3.0.0.2�ǉ�
    End If      'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
    
    MyName = Dir(MyPath & HOSHULOG_FILE, vbNormal) ' �ŏ��̃f�B���N�g������Ԃ��܂��B
    If MyName <> "" Then
      iLogfileCnt = iLogfileCnt + 1
      ReDim Preserve uLogfileData(iLogfileCnt)
      '���O�t�@�C�������i�[����
      uLogfileData(iLogfileCnt - 1).sPath = MyPath
      uLogfileData(iLogfileCnt - 1).sName = HOSHULOG_FILE
      uLogfileData(iLogfileCnt - 1).dtFileDate = FileDateTime(MyPath & HOSHULOG_FILE)
      uLogfileData(iLogfileCnt - 1).lFileSize = FileLen(MyPath & HOSHULOG_FILE)
      uLogfileData(iLogfileCnt - 1).bSelect = False
    End If
    
    '���O�g���[�X�t�@�C������������B
'    MyPath = PATH_LOG                           ' �p�X��ݒ肵�܂��B   'EG20 V2.1.0.1 DEL
'    MyName = Dir(MyPath & "L*.DAT", vbNormal)   ' �ŏ��̃f�B���N�g������Ԃ��܂��B     'EG20 V2.1.0.1 DEL �y��-331�Ή��z
'    MyName = Dir(MyPath & "L*.*", vbNormal)   ' �ŏ��̃f�B���N�g������Ԃ��܂��B        'EG20 V2.1.0.1 ADD �y��-331�Ή��z  EG20 V3.0.0.2 DEK
' EG20 V3.0.0.2�ǉ��J�n
    ' �������ׂ��t�@�C������ύX����B
    If bSelectLogSousa = True Then
        MyName = Dir(MyPath & "*.TXT", vbNormal)        ' �ŏ��̃f�B���N�g������Ԃ��܂��B
    Else
        MyName = Dir(MyPath & "L*.*", vbNormal)         ' �ŏ��̃f�B���N�g������Ԃ��܂��B
    End If
' EG20 V3.0.0.2�ǉ��I��
    Do While MyName <> ""                       ' ���[�v���J�n���܂��B
        ' ���݂̃f�B���N�g���Ɛe�f�B���N�g���͖������܂��B
        If MyName <> "." And MyName <> ".." Then
' EG20 V3.0.0.2�ǉ��J�n
            ' �I�v�V�����ɉ����Č������ׂ��t�@�C���̏������i�肱�݂���
            bFileOK = False
            If bSelectLogSousa = True Then
                ' ���샍�O�ɂ��Ă͌��󖳏���
                bFileOK = True
            Else
                ' �A�v���P�[�V�������O
                If Right(MyName, 3) = "IDU" Or Right(MyName, 3) = "DAT" Then        'EG20 V2.1.0.1 ADD �y��-331�Ή��z
                    bFileOK = True
                End If
            End If
' EG20 V3.0.0.2�ǉ��I��
            If bFileOK = True Then
                ' �r�b�g�P�ʂ̔�r���s���AMyName ���f�B���N�g�����ǂ����𒲂ׂ܂��B
                If (GetAttr(MyPath & MyName) And vbDirectory) <> vbDirectory Then
                    iLogfileCnt = iLogfileCnt + 1
                    ReDim Preserve uLogfileData(iLogfileCnt)
    
                    '���O�t�@�C�������i�[����
                    uLogfileData(iLogfileCnt - 1).sPath = MyPath
                    uLogfileData(iLogfileCnt - 1).sName = MyName
                    uLogfileData(iLogfileCnt - 1).dtFileDate = FileDateTime(MyPath & MyName)
                    uLogfileData(iLogfileCnt - 1).lFileSize = FileLen(MyPath & MyName)
                    uLogfileData(iLogfileCnt - 1).bSelect = False
    
                End If                      ' �����\�����܂��B
            End If
        End If          'EG20 V2.1.0.1 ADD �y��-331�Ή��z
        ' ���̃f�B���N�g������Ԃ��܂��B
        MyName = Dir
    Loop
    fGetLogfileInf = iLogfileCnt
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdLog_Click
'//  �@�\����  : �e�t����������
'//  �@�\�T�v  : �e�t���̂̏������s���B
'//              �u���O�\��(�e�L�X�g�\��)�v�u���O�}�̏o�́v
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.5.0.1) 2009-05-08   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�P�s��Ή�
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 ���O�t�@�C�������ݐ�f�B���N�g���ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ���O�}�̏o�͎��A������T�P�Q���Ƃ���
'//     REVISIONS :(EG20V5.10.0.1) 2012-05-09 REVISED BY [TCC] M.Chiwaki
'//                 ���O�}�̏o�͎��A�t�H���_���쐬����
'//     REVISIONS :(EG20V5.13.0.1) 2012-06-06 REVISED BY [TCC] H.Sugimoto
'//                 �y�}�̏o�̓t�H���_�쐬�Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdLog_Click(Index As Integer)
    Dim bRet As Boolean
    Dim lRetVal As Double
    Dim sCommand As String
    Dim sWriteDir As String    '�����݃f�B���N�g��
    Dim iObjFileNo As Integer  '�����ݑΏ�̧�ِ�
    On Error GoTo ErrorHandle:
    Dim lngErrCode As Long     '�G���[�R�[�h
    Dim fso As FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g       ' EG20 V5.10.0.1�y���O�t�H���_�쐬�Ή��zADD
    Dim szDefLogFolder As String    ' �o�̓��O�t�H���_                  ' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ�
    Dim szCornerFolder As String    ' �R�[�i                            ' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ�

    On Error Resume Next
 
 Select Case Index   '�{�^���C���f�b�N�X
   Case 0
     '�u����샍�O�Ǘ���ʁF���O�\���t�����v
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_TEXT_HYOUJI_BUTTOM, 0)
      '���O�����f�[�^�������`�F�b�N
      bRet = fLogSearchCheck
      If bRet = False Then    '���O�����f�[�^�ɃG���[������ꍇ
          Exit Sub            '�������I������
      End If

      '���O�e�L�X�g�t�@�C������������
       bRet = fWriteLogtxt
       If bRet = True Then         '���O�e�L�X�g�t�@�C��������ɍ쐬���ꂽ�ꍇ
           sCommand = MN_EXE_MEMO & MN_LOG_FILE        '���s�R�}���h���쐬����
           lRetVal = Shell(sCommand, vbMaximizedFocus) '�m�[�g�p�b�h���N������
           AppActivate lRetVal, True                   '�A�N�e�B�u�i�O�ʕ\���j�ɂ���
           SendKeys "{LEFT}", True
          '�u����샍�O�Ǘ���ʁF���O�\����������v
           Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LOG_KANRI_TEXT_HYOUJI_OK, 0)
       Else
          lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
          '�u����샍�O�Ǘ���ʁF���O�\�������ُ�v
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_KANRI_TEXT_HYOUJI_ERROR, lngErrCode)
       End If

    Case 1
       '�u����샍�O�Ǘ���ʁF���O�}�̏o�͖t�����v
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_OUTPUT_BUTTOM, 0)

        '���O�����f�[�^�������`�F�b�N
        iObjFileNo = fLogSelectCheck
        If iObjFileNo <= 0 Then
            Exit Sub            '�������I������
' EG20 V5.9.0.1�y���O�I������Ή��zADD START
        ElseIf iObjFileNo > LOG_FILECNT_MAX Then
            ' �x�������\��
            MsgBox "�I�����ꂽ�t�@�C����������𒴂��܂����B" _
                    & Chr(vbKeyReturn) & "�I���ł���t�@�C������[" & LOG_FILECNT_MAX & "]���܂łł��B", _
                    vbOKOnly + vbCritical, _
                    "�t�@�C���w��ُ�"
            Exit Sub
' EG20 V5.9.0.1�y���O�I������Ή��zADD END
        End If
        ' ��o����f�B���N�g����I������
'        sWriteDir = pfDirSelection("a:", "���O�t�@�C�������ݐ�̃f�B���N�g���I��")     'V1.12.0.1 DEL
        'sWriteDir = pfDirSelection("H:", "���O�t�@�C�������ݐ�̃f�B���N�g���I��")      'V1.12.0.1 ADD    'V1.20.0.1 DEL
        sWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)    'V1.20.0.1 ADD
        'V1.5.0.1 DEL START
        'frmDir.Caption = "���O�t�@�C�������ݐ�̃f�B���N�g���I��"
        'frmDir.Show 1
        'V1.5.0.1 DEL END
        If sWriteDir <> "" Then
        '�f�B���N�g�����w�肳���΁A���O�t�@�C������o��

' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ��J�n
            szDefLogFolder = fncCreateLogFolder()
            If sWriteDir Like ("*" & szDefLogFolder & "\") = False Then
                ' �t�H���_�����݂��邩�`�F�b�N����
                sWriteDir = sWriteDir & "\" & szDefLogFolder
                Set fso = New FileSystemObject
                If fso.FolderExists(sWriteDir) = False Then
                    ' �t�H���_�����݂��Ȃ��ꍇ�͍쐬����
                    fso.CreateFolder (sWriteDir)
                End If
                Set fso = Nothing
            End If
' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ��I��
' EG20 V5.10.0.1�y���O�t�H���_�쐬�Ή��zADD START
            szCornerFolder = "OPERATE_LOG" & CStr(tabTakuCorner.Tab + 1)
'            If sWriteDir Like "*OPERATE_LOG\" = False Then                 ' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�폜
            If sWriteDir Like ("*" & szCornerFolder & "\") = False Then     ' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ�
                ' �t�H���_�����݂��邩�`�F�b�N����
'                sWriteDir = sWriteDir & "\" & "OPERATE_LOG"                ' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�폜
                sWriteDir = sWriteDir & "\" & szCornerFolder                ' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ�
                Set fso = New FileSystemObject
                If fso.FolderExists(sWriteDir) = False Then
                    ' �t�H���_�����݂��Ȃ��ꍇ�͍쐬����
                    fso.CreateFolder (sWriteDir)
                End If
                Set fso = Nothing
            End If
' EG20 V5.10.0.1�y���O�t�H���_�쐬�Ή��zADD END
            sCopyLogFile sWriteDir, iObjFileNo
        End If
     Case Else
    
    End Select
ErrorHandle:
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdLzhFileWrite_Click
'//  �@�\����  : �u���O���k�}�̏o�́v�t����������
'//  �@�\�T�v  : ���O�̈��k�}�̏o�͂��s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 ���O�t�@�C�����k�����ݐ�f�B���N�g���ʒu�ύX
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 �u���O���k�}�̏o�́v�|�b�v�A�b�v��ʂ�ǉ�
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//                  �u���O���k�}�̏o�́v�t���������ł̕ێ烍�O�I�����t�@�C�����C��
'//     REVISIONS :(EG20 V5.4.0.1) 2012-03-24   REVISED BY [TCC] M.Matsumoto
'//                 �y�v���Y�~�[����-6�Ή��zPASSLOG.TXT�̏o�͂ɑΉ�
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05  CODED BY  [TCC] H.Sugimoto
'//                 �y���k�t�H���_�w��Ή��z
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ���O�}�̏o�͎��A������T�P�Q���Ƃ���
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20V5.13.0.1) 2012-06-06 REVISED BY [TCC] H.Sugimoto
'//                 �y�}�̏o�̓t�H���_�쐬�Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdLzhFileWrite_Click()
    Dim sLzhDirName As String    '.LZḨ�يi�[�f�B���N�g����
    Dim sLzhFileName As String   '.LZḨ�ٖ�
    Dim iObjFileNo As Integer    '���k�Ώ�̧�ِ�
    Dim nIndex As Integer        ' ������                    ' EG20 V5.6.0.1�ǉ�

    Dim fso As FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g       ' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ�
    Dim szDefLogFolder As String    ' �o�̓��O�t�H���_                  ' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ�
    Dim szCornerFolder As String    ' �R�[�i                            ' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ�

    On Error Resume Next
    
    '�u����샍�O�Ǘ���ʁF���O���k�}�̏o�͖t�����v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_LOG_LZH_OUTPUT_BUTTOM, 0)

    '���X�g�{�b�N�X�ŁA�t�@�C�����w�肳��Ă��邩�`�F�b�N����B
    iObjFileNo = fLogSelectCheck
    If iObjFileNo <= 0 Then       '�t�@�C���w�肳��Ă��Ȃ���΁A�����I��
        Exit Sub
' EG20 V5.9.0.1�y���O�I������Ή��zADD START
    ElseIf iObjFileNo > LOG_FILECNT_MAX Then
        ' �x�������\��
        MsgBox "�I�����ꂽ�t�@�C����������𒴂��܂����B" _
               & Chr(vbKeyReturn) & "�I���ł���t�@�C������[" & LOG_FILECNT_MAX & "]���܂łł��B", _
               vbOKOnly + vbCritical, _
               "�t�@�C���w��ُ�"
        Exit Sub
' EG20 V5.9.0.1�y���O�I������Ή��zADD END
    End If
    
    '�f�B���N�g���I����ʂ�\�������A���k�t�@�C���i�[�f�B���N�g�����𓾂�B�i��̫���ިڸ�؁��e�c�j
'    sLzhDirName = pfDirSelection("a:", "���O�t�@�C�����k�����ݐ�̃f�B���N�g���I��")   'V1.12.0.1 DEL
    'sLzhDirName = pfDirSelection("H:", "���O�t�@�C�����k�����ݐ�̃f�B���N�g���I��")    'V1.12.0.1 ADD  'V1.20.0.1 DEL
    sLzhDirName = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
    If sLzhDirName = "" Then Exit Sub  '�f�B���N�g�����w�肳��Ȃ���΁A�����I��
 
' EG20 V5.6.0.1�y���k�t�H���_�w��Ή��z�ǉ��J�n
    ' �o�̓t�H���_�ɔ��p�X�y�[�X���܂܂�Ă���ꍇ�A���k�ňُ킪�������Ă��܂�����
    ' ���k�O�Ƀ`�F�b�N���Ĉُ��\������B
    nIndex = InStr(sLzhDirName, " ")
    If nIndex <> 0 Then
        ' �x���|�b�v�A�b�v�E�B���h�E��\������B
        Call MsgBox(CABFOLDERSELECT_ERRORMESSAGE, vbCritical, CABFOLDERSELECT_ERRORTITLE)
        Exit Sub  '�f�B���N�g�����w�肳��Ȃ���΁A�����I��
    End If
' EG20 V5.6.0.1�y���k�t�H���_�w��Ή��z�ǉ��I��

 'V1.20.0.1 ADD START
     '�I�����ꂽ�t�@�C����HOSHU_LOG.dat���AL*.dat���̃`�F�b�N���s���B
     If sObjectTopFile = HOSHULOG_FILE Then
        sLzhFileName = Left$(sObjectTopFile, 9)
     'EG20 V5.4.0.1 ADD START �y�v���Y�~�[����-6�Ή��z
     '�I�����ꂽ�t�@�C����PASSLOG.txt�̏ꍇ
     ElseIf sObjectTopFile = PASSLOG_FILE Then
        sLzhFileName = Left$(sObjectTopFile, 7)
     'EG20 V5.4.0.1 ADD END
     Else
 'V1.20.0.1 ADD END
        '�P�Ԗڂ̃t�@�C��(�g���q���܂܂Ȃ��W����)���A.LZH�t�@�C�����p�Ɏ�o���B
        sLzhFileName = Left$(sObjectTopFile, 8)
    
     End If  'V1.20.0.1 ADD
    
    '.LZH�t�@�C��������������B
    If iObjFileNo >= 2 Then
        '�����I���Ȃ�A�I���t�@�C������t������B
        sLzhFileName = sLzhFileName & "." & CStr(iObjFileNo)
    End If
    
    '�g���q�́A.CAB�ł���B
    sLzhFileName = sLzhFileName & ".CAB"

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_LOG_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

' EG20 V5.9.0.1�y���k�t�H���_���Ή��z�ǉ��J�n
    ' ���k�Ώۃt�H���_�i���[�N�j�֑I���������O���R�s�[
    If funcCopyFileTemporary(PATH_LOGOUTTMP, iObjFileNo, sObjectFiles) = False Then
        Call subDeleteFolder(PATH_LOGOUTTMP)
        
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        ' ���O���k�}�̏o�͏������펞�F�u���O���k�}�̏o�́v�|�b�v�A�b�v��\��
        MsgBox "���O���k�}�̏o�͏����ُ͈�I�����܂����B", _
                vbOKOnly + vbInformation, _
                "���O���k�}�̏o��"
        Exit Sub
    End If
    sObjectFiles = PATH_LOGOUTTMP
' EG20 V5.9.0.1�y���k�t�H���_���Ή��z�ǉ��I��

' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ��J�n
    szDefLogFolder = fncCreateLogFolder()
    If sLzhDirName Like ("*" & szDefLogFolder & "\") = False Then
        ' �t�H���_�����݂��邩�`�F�b�N����
        sLzhDirName = sLzhDirName & "\" & szDefLogFolder
        Set fso = New FileSystemObject
        If fso.FolderExists(sLzhDirName) = False Then
            ' �t�H���_�����݂��Ȃ��ꍇ�͍쐬����
            fso.CreateFolder (sLzhDirName)
        End If
        Set fso = Nothing
    End If
    
    szCornerFolder = "OPERATE_LOG" & CStr(tabTakuCorner.Tab + 1)
    If sLzhDirName Like ("*" & szCornerFolder & "\") = False Then
        ' �t�H���_�����݂��邩�`�F�b�N����
        sLzhDirName = sLzhDirName & "\" & szCornerFolder
        Set fso = New FileSystemObject
        If fso.FolderExists(sLzhDirName) = False Then
            ' �t�H���_�����݂��Ȃ��ꍇ�͍쐬����
            fso.CreateFolder (sLzhDirName)
        End If
        Set fso = Nothing
        sLzhDirName = sLzhDirName & "\"
    End If
' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ��I��

    Call psCabReqest(CABREQEST.CAB_COMPRESSION, sLzhDirName & sLzhFileName, sObjectFiles)
    'V1.20.0.1 ADD START
    If (glngCabErrCd = 0) Then   '���k���ʂ�����(0)
        
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
        ' ���O���k�}�̏o�͏������펞�F�u���O���k�}�̏o�́v�|�b�v�A�b�v��\��
        MsgBox "���O���k�}�̏o�͏����͐���I�����܂����B", _
                vbOKOnly + vbInformation, _
                "���O���k�}�̏o��"
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        Call subDeleteFolder(PATH_LOGOUTTMP)
        Exit Sub
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    End If
    'V1.20.0.1 ADD END

' EG20 V5.9.0.1�y���k�t�H���_���Ή��z�ǉ��J�n
    Call subDeleteFolder(PATH_LOGOUTTMP)
' EG20 V5.9.0.1�y���k�t�H���_���Ή��z�ǉ��I��
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    
    On Error Resume Next

    '�u����샍�O�Ǘ���ʁF�����v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_LOG_TAKU_GAMEN_END, 0)
  
    '����샍�O�Ǘ���ʂ����
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fLogSearchCheck
'//  �@�\����  : ���O�����f�[�^�`�F�b�N����
'//  �@�\�T�v  : ���O�����f�[�^�̐��������`�F�b�N����B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F�u���O�\��(�e�L�X�g�\��)�t������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
'//                 �y�t�F�[�Y�Q�Ή��z
'//     REVISIONS :(EG20 V5.5.0.1) 2012-03-26   CODED   BY [TCC] M.Matsumoto
'//                 �y����No55�Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fLogSearchCheck() As Boolean
    Dim bRet As Boolean         '�֐��̖߂�l
    Dim i As Integer            '�J�E���^
    Dim j As Integer            '�R���g���[���z��
    Dim bFlag As Boolean        '�t���O
    Dim iSelectedLines As Integer '���X�g�{�b�N�X�őI�𒆂̍s��

    On Error Resume Next
    
    fLogSearchCheck = False     '�߂�l�ɏ����l�Ƃ��ăG���[���Z�b�g

    '���X�g�{�b�N�X�őI�𒆂�̧�ق����߽�������sObjectFiles�ɃZ�b�g����B�I�𒆍s���𓾂�B
    iSelectedLines = fSelectedFilesGet
    '�\���t�@�C���w��̃`�F�b�N���s��
    If iSelectedLines <= 0 Then
        '�\���t�@�C�����I�����F�u�\���t�@�C�����I���v�|�b�v�A�b�v��\��
        MsgBox "�\���t�@�C�����I������Ă��܂���B" _
               & Chr(vbKeyReturn) & "�I�����Ă��������B", _
               vbOKOnly + vbExclamation, _
               "����샍�O�Ǘ�"
        Exit Function                   '�������I������
    ElseIf iSelectedLines >= 2 Then
        '�����t�@�C���I�����F�u�����t�@�C���w��v�|�b�v�A�b�v��\��
        MsgBox "�����t�@�C�����I������Ă��܂��B" _
               & Chr(vbKeyReturn) & "������I�����Ă��������B", _
               vbOKOnly + vbExclamation, _
               "����샍�O�Ǘ�"
        Exit Function                   '�������I������
    End If

    '���O�f�[�^�Ώێ����̐������`�F�b�N
    bRet = fLogTimeCheck
    If bRet = False Then                '�G���[�����鎞�͏������I������B
        Exit Function
    End If

    '�w���ʂ̃`�F�b�N���s��
    If optLogSyu(1).Value = True Then   '�w���ʂ�I��������
        j = chkLogSyu.UBound
        bFlag = False
        For i = 0 To j                  '�w���ʕ��J��Ԃ�
            If chkLogSyu(i).Value = CHECKBOX_ON Then
                bFlag = True            '�w�肪��ł�����΁A�`�F�b�N�����I��
                Exit For
            End If
        Next
        If bFlag = False Then
        '�w���ʖ��I�����F�u�w���ʂȂ��v�|�b�v�A�b�v��\��
            MsgBox "�w���ʂ��I������Ă��܂���B" _
                   & Chr(vbKeyReturn) & "�I�����Ă��������B", _
                   vbOKOnly + vbExclamation, _
                   "����샍�O�Ǘ�"
            Exit Function               '�������I������
        End If
    End If

    '�w�蕪�ނ̃`�F�b�N���s��
'    If optLogBunrui(1).Value = True Then   '�w�蕪�ނ�I��������       'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�z
    If optLogBunrui(1).Value = True And optApp(tabTakuCorner.Tab).Value = True Then   '�w�蕪�ނ�I��������   'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�z
        bFlag = False
        For i = 0 To iModCnt             '�w�蕪�ޕ��J��Ԃ�
            If chkMod(i).Visible = True And _
               chkMod(i).Value = CHECKBOX_ON Then
                bFlag = True            '�w�肪�ЂƂł�����΁A�`�F�b�N�����I��
                Exit For
            End If
        Next
        If bFlag = False Then
        '�w�蕪�ޖ��I�����F�u�w�蕪�ނȂ��v�|�b�v�A�b�v��\��
            MsgBox "�w�蕪�ނ��I������Ă��܂���B" _
                   & Chr(vbKeyReturn) & "�I�����Ă��������B", _
                   vbOKOnly + vbExclamation, _
                   "����샍�O�Ǘ�"
            Exit Function               '�������I������
        End If
    End If

    '�w�荆�@�̃`�F�b�N���s��
'    If optLogGouki(1).Value = True Then   '�w�荆�@��I��������    'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
    If optApp(tabTakuCorner.Tab).Value = True Then    '�A�v�����O��I��������     'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
'        j = chkLogGouki.UBound     'EG20 V5.5.0.1 DEL �y����No55�Ή��z
        j = UBound(mintStatus)
        bFlag = False
        For i = 0 To j                 '�w�荆�@���J��Ԃ�
'            If chkLogGouki(i).Value = CHECKBOX_ON Then             'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
            'EG20 V5.5.0.1 DEL START �y����No55�Ή��z
'            If chkLogGouki(i).Visible = True And chkLogGouki(i).Value = CHECKBOX_ON Then    'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
            'EG20 V5.5.0.1 DEL END
            If mintStatus(i) = CHECKBOX_ON Then         'EG20 V5.5.0.1 ADD �y����No55�Ή��z
                bFlag = True            '�w�肪��ł�����ꍇ�A�`�F�b�N�����I��
                Exit For
            End If
        Next
        If bFlag = False Then
        '�w�荆�@���I�����F�u�w�荆�@�Ȃ��v�|�b�v�A�b�v�\��
            MsgBox "�w�荆�@���I������Ă��܂���B" _
                   & Chr(vbKeyReturn) & "�I�����Ă��������B", _
                   vbOKOnly + vbExclamation, _
                   "����샍�O�Ǘ�"
            Exit Function               '�������I������
        End If
    End If

    fLogSearchCheck = True              '�߂�l�ɐ�����Z�b�g
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fSelectedFilesGet
'//  �@�\����  : �I���t�@�C���擾����
'//  �@�\�T�v  : �I�𒆂̃t�@�C���̃t���p�X���擾����B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F���O�����f�[�^�`�F�b�N����
'//�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�u���O�}�̏o�́v
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Integer�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fSelectedFilesGet() As Integer
    Dim iLine As Integer         '۸�̧��ؽ��ޯ���̍s���ޯ��
    Dim iMaxLine As Integer      '۸�̧��ؽ��ޯ���̍s��
    Dim sLineFile As String      '۸�̧��ؽ��ޯ���w��s��̧�ٖ�
    Dim iFileCounter As Integer  '�Ώ�̧�ِ��J�E���^
    
    sObjectFiles = ""
    '���X�g�{�b�N�X�\�����̑S�s�ɂ��Ĉȉ������{����B
'    iMaxLine = lstLogFile.ListCount  '۸�̧��ؽ��ޯ���̍s���𓾂�B    'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
    iMaxLine = lstLogFile(tabTakuCorner.Tab).ListCount  '۸�̧��ؽ��ޯ���̍s���𓾂�B  'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
    
    iFileCounter = 0
    For iLine = 0 To iMaxLine - 1
'        If lstLogFile.Selected(iLine) = True Then      'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
        If lstLogFile(tabTakuCorner.Tab).Selected(iLine) = True Then    'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
        '�I�����ꂽ�s�Ȃ�΁A�Y���s�̃t�@�C���������X�g�{�b�N�X���瓾��B
            'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
'            sLineFile = Left$(lstLogFile.List(iLine), _
'                              InStr(lstLogFile.List(iLine), " ") - 1)
            'EG20 V2.1.0.1 DEL END
            'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
            sLineFile = Left$(lstLogFile(tabTakuCorner.Tab).List(iLine), _
                              InStr(lstLogFile(tabTakuCorner.Tab).List(iLine), " ") - 1)
            'EG20 V2.1.0.1 DEL END
            '�Ώ�̧�قƂ������߽���쐬���A������Ƃ��ĕۑ�����B
'            sObjectFiles = sObjectFiles & PATH_LOG & sLineFile & " "       'EG20 V2.1.0.1 DEL
            'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
            If optHoshu(tabTakuCorner.Tab).Value = True Then
                sObjectFiles = sObjectFiles & PATH_LOG_CORNER & CStr(tabTakuCorner.Tab + 1) & DIR_LOG_SOUSA & sLineFile & " "
            Else
                sObjectFiles = sObjectFiles & PATH_LOG_CORNER & CStr(tabTakuCorner.Tab + 1) & DIR_LOG_APL & sLineFile & " "
            End If
            'EG20 V2.1.0.1 ADD END
            If iFileCounter = 0 Then
            '�I���s���̐擪�i�ŋ��j̧�قł���΁A̧�ٖ��i�g���q���܂�12�����j��ۑ�����B
                sObjectTopFile = sLineFile
            End If
            iFileCounter = iFileCounter + 1
        End If
    Next
    '�I��̧�ق̐���Ԃ��B
    fSelectedFilesGet = iFileCounter
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fLogTimeCheck
'//  �@�\����  : ���O�Ώێ����`�F�b�N����
'//  �@�\�T�v  : ���O�Ώێ����̐������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F���O�����f�[�^�`�F�b�N����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fLogTimeCheck() As Boolean
    Dim i As Integer                '�J�E���^
    Dim bFlag As Boolean            '���̓t���O
    Dim bFromFlag As Boolean        '���̓t���O(�J�n������)
    Dim bToFlag As Boolean          '���̓t���O(�I��������)
    Dim iErrorIndex As Integer      '�G���[�̂���C���f�b�N�X

    fLogTimeCheck = True
    
    '�\���F�����ɖ߂�
    For i = 0 To 5
        txtLogTime(i).ForeColor = MN_COLOR_BLACK
        txtLogTime(i).BackColor = MN_COLOR_WHITE
    Next

    '���͂����邩�`�F�b�N���s��
    bFlag = False                   '�����ɂ���
    bFromFlag = False               '�����ɂ���
    bToFlag = False                 '�����ɂ���
    For i = 0 To 5
        If Not IsNull(txtLogTime(i)) And txtLogTime(i) <> "" Then
            bFlag = True            '�L���ɂ���
            If i >= 0 And i <= 2 Then
                bFromFlag = True    '�L���ɂ���
            Else
                bToFlag = True      '�L���ɂ���
            End If
            Select Case i
            Case 0, 3
                If Int(txtLogTime(i)) < 1 Or Int(txtLogTime(i)) > 31 Then
                    iErrorIndex = i
                    GoTo ErrorHandle
                End If
            Case 1, 4
                If Int(txtLogTime(i)) < 0 Or Int(txtLogTime(i)) > 23 Then
                    iErrorIndex = i
                    GoTo ErrorHandle
                End If
            Case 2, 5
                If Int(txtLogTime(i)) < 0 Or Int(txtLogTime(i)) > 59 Then
                    iErrorIndex = i
                    GoTo ErrorHandle
                End If
            End Select
        End If
    Next
    If bFlag = False Then           '���͂��ЂƂ��Ȃ�
        Exit Function               '�������I������
    End If

    '�J�n�������݂̂̃`�F�b�N���s��
    If bFromFlag = True And bToFlag = False Then
        If IsNull(txtLogTime(0)) Or txtLogTime(0) = "" Then
            iErrorIndex = 0
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(1)) Or txtLogTime(1) = "" Then
            iErrorIndex = 1
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(2)) Or txtLogTime(2) = "" Then
            iErrorIndex = 2
            GoTo ErrorHandle
        End If
    
    '�I���������݂̂̃`�F�b�N���s��
    ElseIf bFromFlag = False And bToFlag = True Then
        If IsNull(txtLogTime(3)) Or txtLogTime(3) = "" Then
            iErrorIndex = 3
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(4)) Or txtLogTime(4) = "" Then
            iErrorIndex = 4
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(5)) Or txtLogTime(5) = "" Then
            iErrorIndex = 5
            GoTo ErrorHandle
        End If
    
    '�����̃`�F�b�N���s��
    Else
        If IsNull(txtLogTime(0)) Or txtLogTime(0) = "" Then
            iErrorIndex = 0
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(1)) Or txtLogTime(1) = "" Then
            iErrorIndex = 1
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(2)) Or txtLogTime(2) = "" Then
            iErrorIndex = 2
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(4)) Or txtLogTime(4) = "" Then
            iErrorIndex = 4
            GoTo ErrorHandle
        ElseIf IsNull(txtLogTime(5)) Or txtLogTime(5) = "" Then
            iErrorIndex = 5
            GoTo ErrorHandle
        End If
        
        '�I�������X�y�[�X�̎��͊J�n���Ɠ����ɂ���
        If IsNull(txtLogTime(3)) Or txtLogTime(3) = "" Then
            txtLogTime(3) = txtLogTime(0)
        End If
        '�������̔�r���s��
        If CInt(txtLogTime(0)) > CInt(txtLogTime(3)) Then
            If CInt(txtLogTime(0)) < 20 _
            Or CInt(txtLogTime(3)) > 10 Then
                iErrorIndex = 0
                GoTo ErrorHandle
            End If
        ElseIf CInt(txtLogTime(0)) = CInt(txtLogTime(3)) Then
            If CInt(txtLogTime(1)) > CInt(txtLogTime(4)) Then
                iErrorIndex = 1
                GoTo ErrorHandle
            ElseIf CInt(txtLogTime(1)) = CInt(txtLogTime(4)) Then
                If CInt(txtLogTime(2)) > CInt(txtLogTime(5)) Then
                    iErrorIndex = 2
                    GoTo ErrorHandle
                End If
            End If
        End If
    End If
    Exit Function
    
ErrorHandle:
    tabLog.Tab = 1
    txtLogTime(iErrorIndex).SetFocus
    txtLogTime(iErrorIndex).BackColor = MN_COLOR_YELLOW
    fLogTimeCheck = False

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fLogSelectCheck
'//  �@�\����  : ���O�t�@�C����o���`�F�b�N����
'//  �@�\�T�v  : ��o���t�@�C���������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F�u���O�}�̏o�́v�u���O���k�}�̏o�́v
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Integer�@�@�@�@�@�@[OUT]�I�𒆃t�@�C����
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Function fLogSelectCheck() As Integer
    Dim bRet As Boolean                 '�߂�l
    Dim bFlag As Boolean                '�t���O
    Dim lId As Long                     '���[���h�c
    Dim udtMail As ML_KYOTU_INF         '�o�b�t�@�t���b�V���v��
    Dim lngErrCode As Long              '�G���[�R�[�h
    
    On Error Resume Next
    
    '���X�g�{�b�N�X�őI�𒆂�̧�ق����߽�������sObjectFiles�ɃZ�b�g����B�I�𒆍s���𓾂�B
    fLogSelectCheck = fSelectedFilesGet
    If fLogSelectCheck <= 0 Then
    '�t�@�C�����I�����F�u�t�@�C���w��Ȃ��v�|�b�v�A�b�v��\��
        MsgBox "��o���t�@�C�����I������Ă��܂���B" _
               & Chr(vbKeyReturn) & "�I�����Ă��������B", _
               vbOKOnly + vbExclamation, _
               "����샍�O�Ǘ�"
        Exit Function                   '�������I������
    End If

    ' ���ݏ������ݒ��̃t�@�C���i��ԐV�����t�@�C���j�͑ΏۊO�Ƃ���
'    If lstLogFile.Selected(lstLogFile.ListCount - 1) = True Then       'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
    If lstLogFile(tabTakuCorner.Tab).Selected(lstLogFile(tabTakuCorner.Tab).ListCount - 1) = True Then 'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
         '�o�b�t�@�t���b�V���v�������O�v���Z�X�ɑ��M����
          udtMail.udtlHeader.dwId = ML_ID_LGBUFF_REQ
          udtMail.udtlHeader.dwSize = MlSize.BUFF_FLUSH_CMD
          udtMail.udtlHeader.dwProid = RHOSHU_ID
          udtMail.udtlHeader.dwSubArea = 0
          bRet = DssSendMail(MAIL_SLOT_LOG, MlSize.BUFF_FLUSH_CMD, udtMail.udtlHeader)
          If bRet = False Then
            '�u�o�b�t�@�t���b�V���v�����M�ُ�v���O�o��
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
            Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSI_LOG_BUFF_FLUSH_CMD_SEND, lngErrCode)
            Exit Function
          Else
            '�u�o�b�t�@�t���b�V���v�����M����v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSI_LOG_BUFF_FLUSH_CMD_SEND, 0)
          End If
        
        '�o�b�t�@�t���b�V���I���ʒm��M
        bFlag = False
        Do Until bFlag = True
            '���[����M�������s��
            lId = fMailRecieve()
            Select Case lId         '���[���h�c
                Case ML_ID_PROEND_ORD
                    '�u�v���Z�X�I���w���v�̏ꍇ
                    '�u�v���Z�X�I���w����M����v���O�o��
                    Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
                    '�����I���������s��
                    pfAbortProc
                Case ML_ID_LGBUFF_ANS
                    '�u�o�b�t�@�t���b�V���I���v�̏ꍇ
                    '�u�o�b�t�@�t���b�V���I���ʒm��M����v���O�o��
                    Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, KANSI_LOG_BUFF_FLUSH_END_RECV, 0)
                    '���[�v�𔲂���
                    Exit Do
                Case Else
            End Select
            Sleep (MN_MAIL_INTERVAL)
        Loop
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sCopyLogFile
'//  �@�\����  : ���O�t�@�C����o������
'//  �@�\�T�v  : ���O�t�@�C���̎�o�����s���B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F�u���O�}�̏o�́v
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sCopyDir  [IN]�����ݐ�f�B���N�g��
'//  �@�@      : Integer�@ iFileNo   [IN]�����ݐ�t�@�C����
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�@���������@�s��C��
'//     REVISIONS :(1.20.0.1) 2010-03-09  REVISED BY [TCC] S.Yoshimori
'//                 �u���O�}�̏o�́v�|�b�v�A�b�v��ʂ�ǉ�
'//                 �u���O�}�̏o�́v�ł̃G���[���b�Z�[�W�\��
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub sCopyLogFile(sCopyDir As String, iFileNo As Integer)
    Dim sFileName As String
    Dim sCopyFileName As String
    Dim iResponse As Integer        'MsgBox�{�^���R�[�h
    Dim lSts As Long
    Dim iFile As Integer            '�t�@�C�����J�E���^
    Dim iIti As Integer             '�I��̧�����߽������(sObjectFiles)���̕����ʒu
    Dim iNext As Integer            '����A���̕����ʒu
    Dim lngErrCode As Long
    'V1.8.0.1 ADD START
    Dim slogPath    As String
    Dim sGetLogFile As String
    Dim bRet        As Boolean
    'V1.8.0.1 ADD END
        
On Error GoTo COPY_ERROR
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_LOG_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '���X�g�{�b�N�X�őI�𒆂̑S�Ẵt�@�C���ɂ��āA�ȉ������{����B
    iIti = 1
    For iFile = 0 To iFileNo - 1
        iNext = InStr(iIti, sObjectFiles, " ")  '�P�s���t�@�C���֏����ށB
        If iNext = 0 Then Exit For
        '�R�s�[���t�@�C�����t���p�X�i۸��ڰ�̧�فj���Z�b�g����B
        sFileName = Mid$(sObjectFiles, iIti, iNext - iIti)
        iIti = iNext + 1
        '�����ݐ�f�B���N�g���{�t�@�C���i�R�s�[���Ɠ����j�����Z�b�g����B
        'sCopyFileName = sCopyDir & "\" & Right$(sFileName, 12) 'V1.8.0.1 DEL
        'V1.8.0.1 ADD START
        '�t�@�C���p�X���A�t�@�C����(�ő�13�o�C�g)�݂̂��擾����B
        sGetLogFile = Right$(sFileName, 13)
        'L*.dat�@or�@HOSHU_LOG.dat�̃`�F�b�N���s���B
        '���f��́u\�v�̗L���ɂ��B
        If Left$(sGetLogFile, 1) = "\" Then
          '�u\�v������̂́uL*.dat�v�̂��߁A�u\�v���폜����B
           sGetLogFile = Right$(sFileName, 12)
        End If
        sCopyFileName = sCopyDir & "\" & sGetLogFile
        'V1.8.0.1 ADD END
        '���O�g���[�X�t�@�C�����w��t�@�C���ɏ����o���B
        'FileCopy sFileName, sCopyFileName              'V1.8.0.1 DEL
        'V1.8.0.1 ADD�@START
        lSts = CopyFile(sFileName, sCopyFileName, 0)
        If lSts = 0 Then
           GoTo COPY_ERROR
        End If
        'V1.8.0.1 ADD�@END
    Next
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    'V1.20.0.1 ADD START
    ' ���O�}�̏o�͏������펞�F�u���O�}�̏o�́v�|�b�v�A�b�v��\��
    MsgBox "���O�}�̏o�͏����͐���I�����܂����B", _
           vbOKOnly + vbInformation, _
           "���O�}�̏o��"
    'V1.20.0.1 ADD END
        
    '�u����샍�O�Ǘ���ʁF���O�}�̏o�͏�������v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LOG_OUTPUT_OK, 0)
    
    Exit Sub

COPY_ERROR:
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    'Select Case Err.Number        'V1.20.0.1 DEL
    Select Case Err.LastDllError   'V1.20.0.1 ADD
        'Case 61 ' �R�s�[��󂫗e�ʕs�����F�u�󂫗e�ʖ����v�|�b�v�A�b�v��\��  'V1.20.0.1 DEL
        Case 112 ' �R�s�[��󂫗e�ʕs�����F�u�󂫗e�ʖ����v�|�b�v�A�b�v��\��   'V1.20.0.1 ADD
            iResponse = MsgBox("�󂯑��̃h���C�u�̃f�B�X�N�������ς��ł��B" _
               & Chr(vbKeyReturn) & "�V�����f�B�X�N��}�����Ă��������B", _
               vbOKOnly, _
               "���O�}�̏o��")

        'Case 70 ' ���C�g�v���e�N�g���F�u�����݋֎~�v�|�b�v�A�b�v��\�� 'V1.20.0.1 DEL
         Case 19 ' ���C�g�v���e�N�g���F�u�����݋֎~�v�|�b�v�A�b�v��\�� 'V1.20.0.1 ADD
            lSts = CopyFile(sFileName, sCopyFileName, 0)
            If (lSts = 0) Then
                iResponse = MsgBox("�t�@�C�����쐬�܂��͒u���ł��܂���B���̃f�B�X�N�̓��C�g�v���e�N�g����Ă܂��B" _
                   & Chr(vbKeyReturn) & "���C�g�v���e�N�g���������邩�@�ʂ̃f�B�X�N���g���Ă��������B", _
                   vbOKOnly, _
                   "���O�}�̏o��")
            End If

        'Case 71 ' �f�B�X�N�𖢑}�����F�u�}�̖��}���v�|�b�v�A�b�v��\�� 'V1.20.0.1 DEL
        Case 21, 3    ' �f�B�X�N�𖢑}�����F�u�}�̖��}���v�|�b�v�A�b�v��\�� 'V1.20.0.1 ADD
            iResponse = MsgBox("�h���C�u�Ƀf�B�X�N�������Ă܂���B" _
               & Chr(vbKeyReturn) & "�f�B�X�N��}�����Ă����蒼���Ă��������B", _
               vbOKOnly, _
               "���O�}�̏o��")
'V1.20.0.1 DEL START
'        Case 75 ' �����Ȃ��^�p�X���ԈႢ���F�u�t�H���_�����ݕs�v�|�b�v�A�b�v��\��
'            iResponse = MsgBox("�R�s�[��̋󂫗e�ʂ��s�����Ă��܂��B" _
'               & Chr(vbKeyReturn) & "�s�v���t�@�C�����폜���邩�A�f�B�X�N�����ւ��Ă������� ", _
'               vbOKOnly, _
'               "���O�}�̏o��")
'V1.20.0.1 DEL END
        Case Else '��L�ȊO���F�u�t�@�C���o�ُ͈�v�|�b�v�A�b�v��\��
            iResponse = MsgBox("�\�����ʃG���[���������܂����B" _
               & Chr(vbKeyReturn) & "�������蒼���Ă��������B", _
               vbOKOnly, _
               "���O�}�̏o��")
    End Select
    
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '�u����샍�O�Ǘ���ʁF���O�}�̏o�͏����ُ�v
     Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_OUTPUT_ERROR, lngErrCode)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fWriteLogtxt
'//  �@�\����  : ���O�e�L�X�g�t�@�C�������ݏ���
'//  �@�\�T�v  : ���O�t�@�C�������O�e�L�X�g�t�@�C���ɏ������ށB
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u���O�\��(�e�L�X�g�\���j�v
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y03����TR-No.115�C���Ή��z
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Function fWriteLogtxt() As Boolean
    Dim uLogConv As LOGCONV             '���O�����f�[�^
    Dim bRet As Boolean                 '�߂�l
    Dim sFileName As String
    Dim lId As Long                     '���[���h�c
    Dim bFlag As Boolean                '�t���O
    Dim iResponse As Integer            'MsgBox�{�^���R�[�h
    Dim iStatus As Long
    Dim udtMail As ML_KYOTU_INF         '�o�b�t�@�t���b�V���v��
    Dim lngErrCode As Long              '�G���[�R�[�h
    fWriteLogtxt = False

    'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
    Dim lngRet As Long                  '�߂�l
    Dim iFilePathLen As Integer
    Dim iresult As Integer
    Dim iErrRet As Integer
    Dim sDatFileName As String
    Dim sSourceFileName As String
    Dim fso As New FileSystemObject

    iErrRet = 0
    'EG20 V2.1.0.1 ADD END   �y�t�F�[�Y�Q�Ή��z

    On Error Resume Next
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_LOG_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '���O�ϊ������쐬����
    sGetSearchData uLogConv
   
' EG20 V3.0.0.2 �폜�J�n�i����샍�O�ɂ͕s�v�j
'   '�o�b�t�@�t���b�V���v�������O�v���Z�X�ɑ��M����
'    udtMail.udtlHeader.dwId = ML_ID_LGBUFF_REQ
'    udtMail.udtlHeader.dwSize = MlSize.BUFF_FLUSH_CMD
'    udtMail.udtlHeader.dwProid = RHOSHU_ID
'    udtMail.udtlHeader.dwSubArea = 0
'    bRet = DssSendMail(MAIL_SLOT_LOG, MlSize.BUFF_FLUSH_CMD, udtMail.udtlHeader)
'    If bRet = False Then
'       '�u�o�b�t�@�t���b�V���v�����M�ُ�v���O�o��
'       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
'       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, KANSI_LOG_BUFF_FLUSH_CMD_SEND, lngErrCode)
'    Else
'       '�u�o�b�t�@�t���b�V���v�����M����v���O�o��
'       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, KANSI_LOG_BUFF_FLUSH_CMD_SEND, 0)
'    End If
'
'   If bRet = True Then
'
'       '�o�b�t�@�t���b�V���I���ʒm��M
'       bFlag = False
'       Do Until bFlag = True
'          '���[����M�������s��
'          lId = fMailRecieve()
'          Select Case lId         '���[���h�c
'            Case ML_ID_PROEND_ORD
'              '�u�v���Z�X�I���w���v�̏ꍇ
'              '�u�v���Z�X�I���w����M����v���O�o��
'               Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
'              '�������I������
'              Exit Function
'            Case ML_ID_LGBUFF_ANS
'              '�u�o�b�t�@�t���b�V���I���v�̏ꍇ
'              '�u�o�b�t�@�t���b�V���I���ʒm��M����v���O�o��
'              Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, KANSI_LOG_BUFF_FLUSH_END_RECV, 0)
'              '���[�v�𔲂���
'              Exit Do
'            Case Else
'            End Select
'          Sleep (MN_MAIL_INTERVAL)
'         Loop
'    End If
' EG20 V3.0.0.2 �폜�I���i����샍�O�ɂ͕s�v�j

    '���O�e�L�X�g�̍쐬
'    sFileName = PATH_LOG & sObjectTopFile      'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
    'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
    If optHoshu(tabTakuCorner.Tab).Value = True Then
        sFileName = PATH_LOG_CORNER & CStr(tabTakuCorner.Tab + 1) & DIR_LOG_SOUSA & sObjectTopFile
        
' EG20 V3.0.0.2 �ǉ��J�n�i���샍�O�͂��̂܂ܕ\���j
        ' �I�����ꂽ�t�@�C�������̂܂܃R�s�[
        If fso.FileExists(sFileName) = True Then
            '�t�@�C���R�s�[�i���ɑ��݂����ꍇ�͏㏑�����邷��j
            fso.CopyFile sFileName, MN_LOG_FILE, True
            fWriteLogtxt = True
        Else
            fWriteLogtxt = False
        End If
        Set fso = Nothing
        
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        Exit Function
' EG20 V3.0.0.2 �ǉ��I���i���샍�O�͂��̂܂ܕ\���j
    Else
        sFileName = PATH_LOG_CORNER & CStr(tabTakuCorner.Tab + 1) & DIR_LOG_APL & sObjectTopFile
    End If
    'EG20 V2.1.0.1 ADD END
    
        'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
    iFilePathLen = Len(sFileName)
    
    iresult = 0
    
    iresult = InStr(sFileName, "IDU")
    
    If iFilePathLen = ((iresult - 1) + 3) Then
    
' EG20 V3.6.0.1�y03����TR-No.115�C���Ή��z�폜�J�n
'        'IDU�t�@�C�� �� CAB�t�@�C���ϊ�
'        bRet = dllCreateDispLogFile2(lngErrCode, sFileName, CAB_LOG_FILE)
'
'        'CAB�t�@�C���ϊ�����H
'        If bRet = True Then
'
'            'CAB�t�@�C����
'             lngRet = pfCabKaito(CAB_LOG_FILE, PATH_WORK)
'
'             If lngRet = 0 Then
'
'                sDatFileName = Replace(sObjectTopFile, "IDU", "DAT")
'                sSourceFileName = PATH_WORK & sDatFileName
'
'                fso.DeleteFile (DAT_LOG_FILE)
'                Name sSourceFileName As DAT_LOG_FILE
'
'                sFileName = DAT_LOG_FILE
'
'             Else
'                fWriteLogtxt = False
'                iErrRet = 1
'             End If
'        Else
'            fWriteLogtxt = False
'            iErrRet = 1
'        End If
' EG20 V3.6.0.1�y03����TR-No.115�C���Ή��z�폜�I��
' EG20 V3.6.0.1�y03����TR-No.115�C���Ή��z�ǉ��J�n
        'IDU�t�@�C�� �� DAT�t�@�C���ϊ�
        bRet = dllCreateDispLogFile2(lngErrCode, sFileName, CAB_LOG_FILE, PATH_WORK)
        ' DAT�t�@�C���ϊ�����H
        If bRet <> True Then
            fWriteLogtxt = False
            iErrRet = 1
        End If
        sFileName = DAT_LOG_FILE
' EG20 V3.6.0.1�y03����TR-No.115�C���Ή��z�ǉ��I��
    
    End If
    'EG20 V2.1.0.1 ADD END   �y�t�F�[�Y�Q�Ή��z
   
    'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
    If iErrRet = 0 Then
    'EG20 V2.1.0.1 ADD END   �y�t�F�[�Y�Q�Ή��z
   
        iStatus = dllbLog2Text(sFileName, uLogConv)
        If iStatus = 2 Then
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
            
            '�ُ�T�C�Y���F�u�\���f�[�^�ʃI�[�o�[�v�|�b�v�A�b�v��\��
            iResponse = MsgBox("�f�[�^�ʂ��������āA�S�Ă�\���ł��܂���B" _
                        & Chr(vbKeyReturn) & "�ꕔ���݂̂ł��\�����܂����H", _
                        vbYesNo + vbExclamation, _
                        "�\���f�[�^�ʃI�[�o�[")
            If iResponse = vbYes Then
                fWriteLogtxt = True
            Else
                fWriteLogtxt = False
            End If
            Exit Function          ' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ�
        ElseIf iStatus = 1 Then    '����̂Ƃ�
            fWriteLogtxt = True
        Else                    '�G���[�̂Ƃ�
            fWriteLogtxt = False
        End If
    'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
    End If
    'EG20 V2.1.0.1 ADD END   �y�t�F�[�Y�Q�Ή��z
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sGetSearchData
'//  �@�\����  : ���O�ϊ����쐬����
'//  �@�\�T�v  : ����샍�O�Ǘ���ʂ��A���O�ϊ������쐬����B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F���O�e�L�X�g�t�@�C�������ݏ���
'//
'//              �^        ����      �Ӗ�
'//  ����      : LOGCONV�@uLogConv�@[OUT]���O�ϊ����
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.1.0.1) 2011-11-19   CODED   BY [TCC] M.Matsumoto
'//                 �y�t�F�[�Y�Q�Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub sGetSearchData(uLogConv As LOGCONV)
    Dim i As Integer                        '�J�E���^
    Dim j As Integer                        '�R���g���[���z��
    Dim sBuff As String                     '������o�b�t�@
    Dim byBuff() As Byte                    '�o�C�g�o�b�t�@
    Dim iProcessID As Integer               '�Ώۃv���Z�XID
    Dim iChangeCnt As Integer               '�ϊ��J�E���^�[(10�i��2�i(�r�b�g)��10�i)
    Dim sChangeProcessId1 As String         '�ϊ���ID[2�i]
    Dim lChangeProcessId2 As Long           '�ϊ���ID[10�i]
    Dim lSetId As Long                      '�G���A�Z�b�gID
    
    On Error Resume Next
     
    '�����͈͂̍쐬���s��
    sBuff = ""                              '����������
    For i = 0 To 5                          '�����͈̓G���A���J��Ԃ�
        If txtLogTime(i) = "" Then          '���͂��Ȃ��ꍇ
            sBuff = sBuff & "  "            '�u�󔒁v���Z�b�g
        Else                                '���͂�����ꍇ
                                            '�Q������������ɂ���
            sBuff = sBuff & Format(txtLogTime(i), "@@")
        End If
    Next
    byBuff = StrConv(sBuff, vbFromUnicode)  '�����ϊ�����
    For i = 0 To TIMEZONE_LEN - 1           '�o�C�g���J��Ԃ�
        uLogConv.byTimeZone(i) = byBuff(i)  '���O�ϊ����Ɋi�[����
    Next

    uLogConv.dw1stAssort = ASRT_NOTUSE      '�u���O���W�Ȃ��v���Z�b�g
    uLogConv.dw2stAssort = ASRT_NOTUSE     '�u���O���W�Ȃ��v���Z�b�g
    uLogConv.by2ndAssort = ASRT_NOTUSE      '�u���O���W�Ȃ��v���Z�b�g
    
    '���ނ̍쐬���s��
'    If optLogBunrui(0).Value = True Then        '���W�I�t�F�u�S�Ă̕��ނ�\���v���L��  'EG20 V2.1.0.1 DEL
    '���W�I�t�F�u�S�Ă̕��ނ�\���v���L���܂��͕ێ烍�O�I��
    If optLogBunrui(0).Value = True Or optHoshu(tabTakuCorner.Tab).Value = True Then      'EG20 V2.1.0.1 ADD
       Process_Settei_ALL uLogConv
    Else                                        '���W�I�t�F�u�w�蕪�ނ̂ݕ\���v���L��
       Process_Settei uLogConv
    End If

    '���O��ʂ̍쐬
    If optLogSyu(0).Value = True Then                 '���W�I�t�F�u�S�Ă̎�ʂ�\���v���L��
        uLogConv.byLogType = LTYP_ALL                 '�u�S��ʁv���Z�b�g�K�v
    Else                                              '���W�I�t�F�u�w���ʂ̂ݕ\���v���L��
        uLogConv.byLogType = LTYP_NOTUSE              '�u�����v���Z�b�g
        If chkLogSyu(0).Value = CHECKBOX_ON Then      '�u����v���L���ȏꍇ
            uLogConv.byLogType = uLogConv.byLogType + LTYP_NORMAL
        End If
        If chkLogSyu(1).Value = CHECKBOX_ON Then      '�u�ُ�v���L���ȏꍇ
            uLogConv.byLogType = uLogConv.byLogType + LTYP_ERROR
        End If
        If chkLogSyu(2).Value = CHECKBOX_ON Then      '�u�x���v���L���ȏꍇ
            uLogConv.byLogType = uLogConv.byLogType + LTYP_WARNING
        End If
        If chkLogSyu(4).Value = CHECKBOX_ON Then      '�u�f�o�b�O�v���L���ȏꍇ
            uLogConv.byLogType = uLogConv.byLogType + LTYP_DEBUG
        End If
    End If

    
    '�t�����t���O�̍쐬
    If optLogData(0).Value = True Then          '���W�I�t�F�u�S�s�\���v���L��
        uLogConv.byOptFlag = 1                  '�u�S�s�\���v���Z�b�g
    Else                                        '���W�I�t�F�u�P�s�ڂ̂ݕ\���v���L��
        uLogConv.byOptFlag = 0                  '�u�P�s�\���v���Z�b�g
    End If

    '�������@���̍쐬
    sBuff = ""

'    j = chkLogGouki.UBound         'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
    j = UBound(mintStatus)           'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z

'EG20 V5.4.0.1 DEL START �y����No49�Ή��z
'    If optLogGouki(0).Value = True Then         '���W�I�t�F�u�S���@�v���L��        'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
'    If optApp(tabTakuCorner.Tab).Value = True Then    '�A�v�����O�I����               'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
'
'        'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
''        For i = 0 To j                          '���@���J��Ԃ�
''            sBuff = sBuff & "1"                 '�Y�����@�Ɂu�L���v���Z�b�g
''        Next
''        For i = j + 1 To GATE_FLAGS_LEN - 1      '���@���J��Ԃ�
''            sBuff = sBuff & "0"                 '�Y�����@�Ɂu�����v���Z�b�g
''        Next
'        'EG20 V2.1.0.1 DEL END
'        'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
'        For i = 0 To j                          '���@���J��Ԃ�
'            If mintStatus(i) = CHECKBOX_ON Then
'                sBuff = sBuff & "1"                 '�Y�����@�Ɂu�L���v���Z�b�g
'            Else
'                sBuff = sBuff & "0"                 '�Y�����@�Ɂu�����v���Z�b�g
'            End If
'        Next
'        'EG20 V2.1.0.1 ADD END
'
'    Else
'EG20 V5.4.0.1 DEL END
    
        For i = 0 To j                          '���@���J��Ԃ�
'            If chkLogGouki(i).Value = CHECKBOX_ON Then '�u�H�H���@�v���L���ȏꍇ   'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
            If mintStatus(i) = CHECKBOX_ON Then '�u�H�H���@�v���L���ȏꍇ           'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
                sBuff = sBuff & "1"             '�Y�����@�Ɂu�L���v���Z�b�g
            Else
                sBuff = sBuff & "0"             '�Y�����@�Ɂu�����v���Z�b�g
            End If
        Next
        'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
'        For i = j + 1 To GATE_FLAGS_LEN - 1     '���@���J��Ԃ�
'            sBuff = sBuff & "0"                 '�Y�����@�Ɂu�����v���Z�b�g
'        Next
        'EG20 V2.1.0.1 DEL END
'    End If         'EG20 V5.4.0.1 DEL �y����No49�Ή��z
    byBuff = StrConv(sBuff, vbFromUnicode)      '�����ϊ�����
    For i = 0 To GATE_FLAGS_LEN - 1             '�o�C�g���J��Ԃ�
        uLogConv.byGateFlag(i) = byBuff(i)      '���O�ϊ����Ɋi�[����
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtLogTime_DblClick
'//  �@�\����  : ���O�f�[�^�������A�_�u���N���b�N������
'//  �@�\�T�v  : �[���e���L�[��ʂ�\��
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F���O�f�[�^�Ώێ����e�L�X�g�{�b�N�X
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@�@Index  �@[IN]�e�L�X�g�{�b�N�X�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub txtLogTime_DblClick(Index As Integer)
    gstrTenKeyData = txtLogTime(Index) ' ���ݐݒ肵�Ă������n��
    gstrTenKeySize = 4                 '���͉\���������w�肷��B
    ' �[���e���L�[��ʕ\��
    frmTenKey.Show 1
    ' �ݒ肵�������X�V����
    txtLogTime(Index) = gstrTenKeyData
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtLogTime_KeyPress
'//  �@�\����  : ���O�f�[�^�������A�L�[���͏���
'//  �@�\�T�v  : ���̓L�[�`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F���O�f�[�^�Ώێ����e�L�X�g�{�b�N�X
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@�@Index  �@[IN]�e�L�X�g�{�b�N�X�C���f�b�N�X
'//  �@�@      : Integer�@�@KeyAscii [IN]���̓L�[
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub txtLogTime_KeyPress(Index As Integer, KeyAscii As Integer)
    
    '�w�i�F�𔒐F�ɂ���
    txtLogTime(Index).BackColor = MN_COLOR_WHITE
    '�����̂ݗL���Ƃ���
    KeyAscii = pfKeyNumeric(KeyAscii)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfKeyNumeric
'//  �@�\����  : �������͏���
'//  �@�\�T�v  : �����ȊO�̕����𖳌��ɂ���B�B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F���O�f�[�^�Ώێ����e�L�X�g�{�b�N�X
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@�@KeyAscii [IN]���̓L�[
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Integer�@�@�@�@�@�@ [OUT]�L�[�R�[�h
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Function pfKeyNumeric(iKeyAscii As Integer) As Integer
    
    '�����l�Ƃ��Ĉ����̃R�[�h��߂�l�Ƃ���
    pfKeyNumeric = iKeyAscii
    
    '�o�b�N�X�y�[�X�L�[�͗L���Ƃ���
    If iKeyAscii = vbKeyBack Then
        Exit Function
    End If
    '�����ȊO�͖����Ƃ���
    If iKeyAscii < vbKey0 Or iKeyAscii > vbKey9 Then
        pfKeyNumeric = 0
        Beep
    End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtLogTime_Change
'//  �@�\����  : ���O�f�[�^�Ώێ������͏���
'//  �@�\�T�v  : �\�����ڎw�蕔�F���O�f�[�^�Ώێ����e�L�X�g�{�b�N�X
'//�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�����G���A�����̓��͒l�`�F�b�N
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub txtLogTime_Change(Index As Integer)
    
    '�K�茅������
    If Len(txtLogTime(Index)) = 2 Then
        Select Case Index
        Case 0, 3
            '���t(��)�̐��������`�F�b�N����
            If pfTextDay(txtLogTime(Index)) <> True Then
                '�O�ʐF���G���[�F�ɂ���
                txtLogTime(Index).ForeColor = MN_COLOR_RED
                Exit Sub
            End If
        Case 1, 4
            '���t(��)�̐��������`�F�b�N����
            If pfTextHour(txtLogTime(Index)) <> True Then
                '�O�ʐF���G���[�F�ɂ���
                txtLogTime(Index).ForeColor = MN_COLOR_RED
                Exit Sub
            End If
        Case 2, 5
            '���t(��)�̐��������`�F�b�N����
            If pfTextMin(txtLogTime(Index).Text) <> True Then
                '�O�ʐF���G���[�F�ɂ���
                txtLogTime(Index).ForeColor = MN_COLOR_RED
                Exit Sub
            End If
        Case Else
        End Select
        If Index < 5 Then
            '�G���[���Ȃ���Ύ��̍��ڂփt�H�[�J�X���ڂ�
            txtLogTime(Index + 1).SetFocus
        End If
    End If
    '�O�ʐF�����F�ɂ���
    txtLogTime(Index).ForeColor = MN_COLOR_BLACK

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfTextDay
'//  �@�\����  : ���t�������`�F�b�N����
'//  �@�\�T�v  : ���t�̐������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F���O�f�[�^�Ώێ����e�L�X�g�{�b�N�X
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sText�@�@[IN]���͓��l
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Function pfTextDay(sText As String) As Boolean
    
    pfTextDay = False
    '���������`�F�b�N����
    If LenB(StrConv(sText, vbFromUnicode)) > 2 Then
        Exit Function
    End If
    '���l�̐������`�F�b�N
    If IsNumeric(sText) = False Then
        Exit Function
    End If
    '�͈̓`�F�b�N���s��
    If CInt(sText) < 1 Or CInt(sText) > 31 Then
        Exit Function
    End If
    pfTextDay = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfTextHour
'//  �@�\����  : ���Ԑ������`�F�b�N����
'//  �@�\�T�v  : ���Ԃ̐������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F���O�f�[�^�Ώێ����e�L�X�g�{�b�N�X
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sText�@�@[IN]���͎��Ԓl
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Function pfTextHour(sText As String) As Boolean
    
    pfTextHour = False
    '���������`�F�b�N����
    If LenB(StrConv(sText, vbFromUnicode)) > 2 Then
        Exit Function
    End If
    '���l�̐������`�F�b�N
    If IsNumeric(sText) = False Then
        Exit Function
    End If
    '�͈̓`�F�b�N���s��
    If CInt(sText) < 0 Or CInt(sText) > 23 Then
        Exit Function
    End If
    pfTextHour = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfTextMin
'//  �@�\����  : �����������`�F�b�N����
'//  �@�\�T�v  : �����̐������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F���O�f�[�^�Ώێ����e�L�X�g�{�b�N�X
'//
'//              �^        ����      �Ӗ�
'//  ����      : String�@�@sText�@�@[IN]���͕����l
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Boolean�@�@�@�@�@�@[OUT]�߂�l
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Function pfTextMin(sText As String) As Boolean
    
    pfTextMin = False
    '���������`�F�b�N����
    If LenB(StrConv(sText, vbFromUnicode)) > 2 Then
        Exit Function
    End If
    '���l�̐������`�F�b�N
    If IsNumeric(sText) = False Then
        Exit Function
    End If
    '�͈̓`�F�b�N���s��
    If CInt(sText) < 0 Or CInt(sText) > 59 Then
        Exit Function
    End If
    pfTextMin = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : optLogSyu_Click
'//  �@�\����  : ��ʃ��W�I�t����������
'//  �@�\�T�v  : �w���ʂ̃A�N�e�B�u�E��A�N�e�B�u�̉�ʍX�V�������s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u��ʁv��
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@ [IN]�������W�I�t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub optLogSyu_Click(Index As Integer)
    '�������W�I�t�ɂ���ʕ\���X�V����
    sLogIndexChange
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : optlogbunrui_Click
'//  �@�\����  : ���ރ��W�I�t����������
'//  �@�\�T�v  : �w�蕪�ނ̃A�N�e�B�u�E��A�N�e�B�u�̉�ʍX�V�������s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u���ށv��
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@ [IN]�������W�I�t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub optlogbunrui_Click(Index As Integer)
    '�������W�I�t�ɂ���ʕ\���X�V����
    sLogIndexChange
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : optAll_Click
'//  �@�\����  : �u�S�đI���v�u�S�Ĕ�I���v�t����������
'//  �@�\�T�v  : �w�蕪�ނ̃`�F�b�NON/OFF���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�w�蕪�ށv��
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@ [IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub optAll_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To iModCnt
        If Index = 0 Then
        '�u�S�đI���v�t�������F�w�蕪�ނ�S�ă`�F�b�N����B
            chkMod(i).Value = vbChecked
        Else
        '�u�S�Ĕ�I���v�t�������F�w�蕪�ނ�S�ă`�F�b�N���Ȃ��B
            chkMod(i).Value = vbUnchecked
        End If
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sLogIndexChange
'//  �@�\����  : ���ڔF���ύX����
'//  �@�\�T�v  : ��ʁA���ނ̉�ʕ\�����X�V����B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�w���ʁv���u�w�蕪�ށv��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-30   REVISED BY [TCC] S.Terao
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub sLogIndexChange()
    Dim i As Integer        '�J�E���^
    Dim j As Integer        '�R���g���[���z��

    '***********************
    '* ��ʃG���A�{�b�N�X  *
    '***********************
    j = chkLogSyu.UBound
    '���W�I�t�F�u�S�Ă̎�ʂ�\���v���L��
    If optLogSyu(0).Value = True Then
        '�S�Ă̎�ʂ��A�N�e�B�u�\���ɂ���
        For i = 0 To j                      '�w���ʐ����J��Ԃ�
            chkLogSyu(i).Enabled = False
        Next
     '���W�I�t�F�u�w���ʂ̂ݕ\���v���L��
    Else
        '�S�Ă̎�ʂ��A�N�e�B�u�\���ɂ���
        For i = 0 To j                      '�w���ʐ����J��Ԃ�
            chkLogSyu(i).Enabled = True
        Next
    End If

    '***********************
    '* ���ރG���A�{�b�N�X  *
    '***********************
    j = iModCnt
    '���W�I�t�F�u�S�Ă̕��ނ�\���v���L��
    If optLogBunrui(0).Value = True Then
        '�S�Ă̕��ނ��A�N�e�B�u�\���ɂ���
        For i = 0 To j                      '�w�蕪�ސ����J��Ԃ�
             chkMod(i).Enabled = False
             'chkMod(i).Value = CHECKBOX_ON 'V1.7.0.1 DEL
        Next
        optAll(0).Enabled = False  '�u�S�đI���v�t���A�N�e�B�u�\���ɂ���B
        optAll(1).Enabled = False  '�u�S�Ĕ�I���v�t���A�N�e�B�u�\���ɂ���B
    '���W�I�t�F�u�w�蕪�ނ̂ݕ\���v���L��
    Else
        '�S�Ă̕��ނ��A�N�e�B�u�\���ɂ���
        For i = 0 To j                     '�w�蕪�ސ����J��Ԃ�
             chkMod(i).Enabled = True
        Next
        optAll(0).Enabled = True  '�u�S�đI���v�t���A�N�e�B�u�\���ɂ���B
        optAll(1).Enabled = True  '�u�S�Ĕ�I���v�t���A�N�e�B�u�\���ɂ���B
    End If
End Sub

'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : optLogGouki_Click
'//  �@�\����  : ���ڔF���ύX����
'//  �@�\�T�v  : ��ʁA���ނ̉�ʕ\�����X�V����B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�w���ʁv���u�w�蕪�ށv��
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@�@[IN]���W�I�t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
'Private Sub optLogGouki_Click(Index As Integer)
'    '�������W�I�t�ɂ���ʕ\���X�V����
'    sOptGoukiChange
'End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdChkAll_Click
'//  �@�\����  : �u�S���@�I���v�t����������
'//  �@�\�T�v  : �S�������@�̃`�F�b�N��ON�ɂ���B
'//�@�@�@�@�@�@�@�\���������@�w�蕔�F�u�������@�v��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
'Private Sub cmdChkAll_Click()
'
'    Dim i As Integer
'
'    For i = 0 To 17
'       chkLogGouki(i).Value = CHECKBOX_ON
'    Next
'End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdChkAllKai_Click
'//  �@�\����  : �u�S���@�����v�t����������
'//  �@�\�T�v  : �S�������@�̃`�F�b�N��OFF�ɂ���B
'//�@�@�@�@�@�@�@�\���������@�w�蕔�F�u�������@�v��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
'Private Sub cmdChkAllKai_Click()
'
'   Dim i As Integer
'
'    For i = 0 To 17
'       chkLogGouki(i).Value = CHECKBOX_OFF
'    Next
'
'End Sub
'EG20 V2.1.0.1 DEL END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : chkMod_Click
'//  �@�\����  : �w�蕪�ނ̊e�`�F�b�N�{�b�N�X��������
'//  �@�\�T�v  : �w�蕪�ނ̊e�`�F�b�N�{�b�N�X��ԍX�V���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�w�蕪�ށv��
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@ [IN]�e�`�F�b�N�{�b�N�X�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub chkMod_Click(Index As Integer)
    Dim iCnt As Integer
    Dim sDai As String
    Dim iChkType As Integer
    
    '�t���J�E���^�[��0���ǂ����`�F�b�N����B
    '�J�E���^�[0�F�啪�ވ����B�J�E���^�[0�ȊO�F�����ވ���
    If Int(uModFileData(Index).iFuzokuCnt) = 0 Then
        '�C���f�b�N�X�ԍ����ŏI�̏ꍇ�A�ŏI�ȍ~�͂Ȃ��̂ŏ����I��
        If Index = iModCnt Then
            Exit Sub
        End If
        
        '�����ވ����̃C���f�b�N�X�ԍ����쐬
        iCnt = Index + 1
        '�����ވ����A�啪�ވ����̃v���Z�XID���擾����B
        sDai = uModFileData(Index).iProces
        '�啪�ވ����̃`�F�b�N�{�b�N�X��Ԓl���擾����B
        iChkType = chkMod(Index).Value
        Do
            '�����ވ����̕t��ID�ƁA�啪�ވ�����ID�Ƃ���v���邩�ǂ����`�F�b�N����B
            '(�������ވ����Ƒ啪�ވ����̌q����m�F)
            If sDai = uModFileData(iCnt).iFuzokuId Then
                '��v�����ꍇ�A�啪�ނ̃`�F�b�N�{�b�N�X��Ԓl�𒆕��ވ����ɂ����f����B
                chkMod(iCnt).Value = iChkType
            Else
                '�s��v�̏ꍇ�A�����I���B
                Exit Do
            End If
            '�����ވ����̂��̂��܂����邩�`�F�b�N����B
            iCnt = iCnt + 1
            If iCnt > iModCnt Then
                '�C���f�b�N�X�ԍ����ŏI�ɂȂ�Ώ����I��
                Exit Sub
            End If
        Loop
    End If
End Sub

'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sOptGoukiChange
'//  �@�\����  : �\���������@�w��ύX����
'//  �@�\�T�v  : ���W�I�t�����ɂ��A��ʕ\�����X�V����B
'//�@�@�@�@�@�@�@�\���������@�w�蕔�F�u�������@�v��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
'Private Sub sOptGoukiChange()
'    Dim i As Integer            '�J�E���^
'    Dim j As Integer        '�R���g���[���z��
'
'    '�w�荆�@
'    j = chkLogGouki.UBound
'    '���W�I�t�F�u�S���@�v���L��
'    If optLogGouki(0).Value = True Then
'        cmdChkAll.Enabled = False
'        cmdChkAllKai.Enabled = False
'        '�S�Ă̍��@���A�N�e�B�u�\���ɂ���
'        For i = 0 To j                    '���@�����J��Ԃ�
'            chkLogGouki(i).Enabled = False
'        Next
'    '���W�I�t�F�u�w�荆�@�̂݁v���L��
'    Else
'         cmdChkAll.Enabled = True
'         cmdChkAllKai.Enabled = True
'        '�S�Ă̍��@���A�N�e�B�u�\���ɂ���
'         For i = 0 To j                   '���@�����J��Ԃ�
'            chkLogGouki(i).Enabled = True
'        Next
'    End If
'
'End Sub
'EG20 V2.1.0.1 DEL END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fMailRecieve
'//  �@�\����  : ���[����M����
'//  �@�\�T�v  : �ێ烁�[���E�X���b�g���烁�[������M����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Integer�@�@�@�@�@�@[OUT]���[��ID
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Function fMailRecieve() As Integer
    Dim lLen As Long                    '���[���T�C�Y
    Dim uMail As ML_KYOTU_INF           '���[��

    On Error Resume Next

    fMailRecieve = 0

    '���[����M
    lLen = DssMailRead(plMSlot_MN, uMail)
    If lLen > 0 Then                            '��M����̎�

      Select Case uMail.udtlHeader.dwId  '���[���h�c
        Case ML_ID_PROEND_ORD
             '�u�v���Z�X�I���w���v����M�����ꍇ
             '�u�v���Z�X�I���w����M����v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
             '�����I���������s��
             pfAbortProc
             '�߂�l�Ƀ��[���h�c���Z�b�g
             fMailRecieve = ML_ID_PROEND_ORD

        Case ML_ID_LGBUFF_ANS
             '�u�o�b�t�@�t���b�V���I���ʒm�v����M�����ꍇ
             '�u�o�b�t�@�t���b�V���I���ʒm��M����v���O�o��
              Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, KANSI_LOG_BUFF_FLUSH_END_RECV, 0)
             '�߂�l�Ƀ��[���h�c���Z�b�g
             fMailRecieve = ML_ID_LGBUFF_ANS

        Case ML_ID_HOSHU_ACTIVE_REQ
             '�ێ��ʃA�N�e�B�u�\���̏ꍇ
             '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
             AppActivate frmTakuLogKanri.Caption, False
             pfFormActive (frmTakuLogKanri.hwnd)
             fMailRecieve = ML_ID_HOSHU_ACTIVE_REQ

        Case ML_ID_LGCHGREQ_RES
             '���O�ؑ֗v��RES�̏ꍇ
             '�u���O�ؑ֗v��RES��M����v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, LOG_CHANGE_RES_RECV, 0)
             fMailRecieve = ML_ID_LGCHGREQ_RES

        Case Else
        '���[���h�c�s��
          '�u���[��ID�s���v���O�o��
          Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
End Function

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
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    Dim lId As Long         '���[���h�c
    '���[������M����'
    lId = fMailRecieve()
    If lId = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmTakuLogKanri.Caption, False
        pfFormActive (frmTakuLogKanri.hwnd)
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Process_Settei
'//  �@�\����  : �啪�ނ̃r�b�g�ݒ菈��
'//  �@�\�T�v  : �啪�ނ̐ݒ菈�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-13   CODED   BY [TCC] C.Terui
'//     REVISIONS :(V30.1.0.1) 2014-05-21   CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub Process_Settei(uLogConv As LOGCONV)
    Dim i As Integer                        '�J�E���^
    Dim iProcessID As Integer               '�Ώۃv���Z�XID
    Dim iChangeCnt As Integer               '�ϊ��J�E���^�[(10�i��2�i(�r�b�g)��10�i)
    Dim sChangeProcessId1 As String         '�ϊ���ID[2�i]
    Dim lChangeProcessId2 As Long           '�ϊ���ID[10�i]
    Dim lSetId As Long                      '�G���A�Z�b�gID
' V1.3.0.1 ADD START
    Dim bit(0 To 31) As Long
    '�r�b�g�錾
    'EG20 V30.1.0.1 DEL START
'    bit(0) = &H1
'    bit(1) = &H2
'    bit(2) = &H4
'    bit(3) = &H8
'    bit(4) = &H10
'    bit(5) = &H20
'    bit(6) = &H40
'    bit(7) = &H80
'    bit(8) = &H100
'    bit(9) = &H200
'    bit(10) = &H400
'    bit(11) = &H800
'    bit(12) = &H1000
'    bit(13) = &H2000
'    bit(14) = &H4000
'    bit(15) = &H8000
'    bit(16) = &H10000
'    bit(17) = &H20000
'    bit(18) = &H40000
'    bit(19) = &H80000
'    bit(20) = &H100000
'    bit(21) = &H200000
'    bit(22) = &H400000
'    bit(23) = &H800000
'    bit(24) = &H1000000
'    bit(25) = &H2000000
'    bit(26) = &H4000000
'    bit(27) = &H8000000
'    bit(28) = &H10000000
'    bit(29) = &H20000000
'    bit(30) = &H40000000
'    bit(31) = &H80000000
    'EG20 V30.1.0.1 DEL END
    'EG20 V30.1.0.1 ADD START
    '&Hxxxx&�ƌ���&�����Ȃ���LONG�^�Ƃ��ď�������Ȃ��̂ŏC���B&H8000���}�C�i�X�l�ɂȂ��Ă��܂��B
    '�r�b�g�錾
    bit(0) = &H1&
    bit(1) = &H2&
    bit(2) = &H4&
    bit(3) = &H8&
    bit(4) = &H10&
    bit(5) = &H20&
    bit(6) = &H40&
    bit(7) = &H80&
    bit(8) = &H100&
    bit(9) = &H200&
    bit(10) = &H400&
    bit(11) = &H800&
    bit(12) = &H1000&
    bit(13) = &H2000&
    bit(14) = &H4000&
    bit(15) = &H8000&
    bit(16) = &H10000
    bit(17) = &H20000
    bit(18) = &H40000
    bit(19) = &H80000
    bit(20) = &H100000
    bit(21) = &H200000
    bit(22) = &H400000
    bit(23) = &H800000
    bit(24) = &H1000000
    bit(25) = &H2000000
    bit(26) = &H4000000
    bit(27) = &H8000000
    bit(28) = &H10000000
    bit(29) = &H20000000
    bit(30) = &H40000000
    bit(31) = &H80000000
    'EG20 V30.1.0.1 ADD END
         
    '�w�蕪�ޕ����[�v����B
      For i = 0 To iModCnt
       '�w�蕪�ގw��L���`�F�b�N���s���B
       If chkMod(i).Value = CHECKBOX_ON Then
          '�Ώۃv���Z�XID���擾����
          iProcessID = uModFileData(i).iProces
          If (0 < iProcessID) And (iProcessID <= 31) Then
' V1.3.0.1 DEL START
'             '�v���Z�XID��2�i���ɕϊ�����B
'             sChangeProcessId1 = 0
'             iChangeCnt = 0
'             For iChangeCnt = 1 To iProcessID
'                If iChangeCnt = 1 Then
'                  '�r�b�g����������B
'                   sChangeProcessId1 = 1
'                Else
'                   sChangeProcessId1 = sChangeProcessId1 & 0
'                End If
'              Next
'
'              lChangeProcessId2 = 0
'              '2�i����10�i���ɕϊ�����B
'              For iChangeCnt = 0 To Len(sChangeProcessId1) - 1
'                 If Mid(sChangeProcessId1, iChangeCnt + 1, 1) <> 0 Then
'                    lChangeProcessId2 = lChangeProcessId2 + 2 ^ (Len(sChangeProcessId1) - iChangeCnt - 1)
'                 End If
'              Next iChangeCnt
'               uLogConv.dw1stAssort = uLogConv.dw1stAssort + lChangeProcessId2
' V1.3.0.1 DEL END
                uLogConv.dw1stAssort = uLogConv.dw1stAssort + bit(iProcessID)       ' V1.3.0.1 ADD

          ElseIf (31 < iProcessID) And (iProcessID < 63) Then
              iProcessID = iProcessID - 32
' V1.3.0.1 DEL START
'                '�v���Z�XID��2�i���ɕϊ�����B
'              iChangeCnt = 0
'              sChangeProcessId1 = 0
'               For iChangeCnt = 1 To iProcessID
'                  If iChangeCnt = 1 Then
'                    '�r�b�g����������
'                     sChangeProcessId1 = 1
'                  Else
'                     sChangeProcessId1 = sChangeProcessId1 & 0
'                  End If
'               Next
'
'               lChangeProcessId2 = 0
'               '2�i����10�i���ɕϊ�����
'               For iChangeCnt = 0 To Len(sChangeProcessId1) - 1
'                  If Mid(sChangeProcessId1, iChangeCnt + 1, 1) <> 0 Then
'                     lChangeProcessId2 = lChangeProcessId2 + 2 ^ (Len(sChangeProcessId1) - iChangeCnt - 1)
'                  End If
'               Next iChangeCnt
'                uLogConv.dw2stAssort = uLogConv.dw2stAssort + lChangeProcessId2
' V1.3.0.1 DEL END
                 uLogConv.dw2stAssort = uLogConv.dw2stAssort + bit(iProcessID)       ' V1.3.0.1 ADD
          End If
        End If
      Next
End Sub
'V1.7.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Process_Settei_ALL
'//  �@�\����  : �啪�ނ̃r�b�g�ݒ菈��(�������S����)
'//  �@�\�T�v  : �啪�ނ̐ݒ菈�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.7.0.1) 2009-07-30   CODED   BY [TCC] S.Terao
'//     REVISIONS :(V30.1.0.1) 2014-05-21   CODED   BY [TCC] T.Nakajima
'//                 �k���V�����J�ƑΉ�
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub Process_Settei_ALL(uLogConv As LOGCONV)
    Dim i As Integer                        '�J�E���^
    Dim iProcessID As Integer               '�Ώۃv���Z�XID
    Dim iChangeCnt As Integer               '�ϊ��J�E���^�[(10�i��2�i(�r�b�g)��10�i)
    Dim sChangeProcessId1 As String         '�ϊ���ID[2�i]
    Dim lChangeProcessId2 As Long           '�ϊ���ID[10�i]
    Dim lSetId As Long                      '�G���A�Z�b�gID
    
    Dim bit(0 To 31) As Long
    '�r�b�g�錾
    'EG20 V30.1.0.1 DEL START
'    bit(0) = &H1
'    bit(1) = &H2
'    bit(2) = &H4
'    bit(3) = &H8
'    bit(4) = &H10
'    bit(5) = &H20
'    bit(6) = &H40
'    bit(7) = &H80
'    bit(8) = &H100
'    bit(9) = &H200
'    bit(10) = &H400
'    bit(11) = &H800
'    bit(12) = &H1000
'    bit(13) = &H2000
'    bit(14) = &H4000
'    bit(15) = &H8000
'    bit(16) = &H10000
'    bit(17) = &H20000
'    bit(18) = &H40000
'    bit(19) = &H80000
'    bit(20) = &H100000
'    bit(21) = &H200000
'    bit(22) = &H400000
'    bit(23) = &H800000
'    bit(24) = &H1000000
'    bit(25) = &H2000000
'    bit(26) = &H4000000
'    bit(27) = &H8000000
'    bit(28) = &H10000000
'    bit(29) = &H20000000
'    bit(30) = &H40000000
'    bit(31) = &H80000000
    'EG20 V30.1.0.1 DEL END
    
    'EG20 V30.1.0.1 ADD START
    '&Hxxxx&�ƌ���&�����Ȃ���LONG�^�Ƃ��ď�������Ȃ��̂ŏC���B&H8000���}�C�i�X�l�ɂȂ��Ă��܂��B
    '�r�b�g�錾
    bit(0) = &H1&
    bit(1) = &H2&
    bit(2) = &H4&
    bit(3) = &H8&
    bit(4) = &H10&
    bit(5) = &H20&
    bit(6) = &H40&
    bit(7) = &H80&
    bit(8) = &H100&
    bit(9) = &H200&
    bit(10) = &H400&
    bit(11) = &H800&
    bit(12) = &H1000&
    bit(13) = &H2000&
    bit(14) = &H4000&
    bit(15) = &H8000&
    bit(16) = &H10000
    bit(17) = &H20000
    bit(18) = &H40000
    bit(19) = &H80000
    bit(20) = &H100000
    bit(21) = &H200000
    bit(22) = &H400000
    bit(23) = &H800000
    bit(24) = &H1000000
    bit(25) = &H2000000
    bit(26) = &H4000000
    bit(27) = &H8000000
    bit(28) = &H10000000
    bit(29) = &H20000000
    bit(30) = &H40000000
    bit(31) = &H80000000
    'EG20 V30.1.0.1 ADD END
         
    '�w�蕪�ޕ����[�v����B
      For i = 0 To iModCnt
       '�w�蕪�ގw��L���`�F�b�N���s���B
       '�Ώۃv���Z�XID���擾����
       iProcessID = uModFileData(i).iProces
       If (0 < iProcessID) And (iProcessID <= 31) Then
          uLogConv.dw1stAssort = uLogConv.dw1stAssort + bit(iProcessID)
       ElseIf (31 < iProcessID) And (iProcessID < 63) Then
          iProcessID = iProcessID - 32
          uLogConv.dw2stAssort = uLogConv.dw2stAssort + bit(iProcessID)
       End If
      Next
End Sub
'V1.7.0.1 ADD END
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : SetEnableFalse
'//  �@�\����  : ��ʃ��b�N����
'//  �@�\�T�v  : ��ʂ̃��b�N������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()

    On Error Resume Next
  
    '�^�u��False�ɂ���B
    tabLog.Enabled = False
    
    '�u���O�}�̏o�́v�t��False�ɂ���B
    cmdLog(1).Enabled = False
    
    '�u���k�}�̏o�́v�t��False�ɂ���B
    cmdLzhFileWrite.Enabled = False
       
    '�u�\���X�V�v�t��False�ɂ���B
    cmdUpdateDisplay.Enabled = False
        
    '�u�������\���v�t��False�ɂ���B
    cmdLog(0).Enabled = False

    '�u�}�̎�O�v�t��False�ɂ���B
    cmdInstall.Enabled = False
    
    '�u�ێ��ʂ֖߂�v�t��False�ɂ���B
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
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()
  
    On Error Resume Next

    '�^�u��True�ɂ���B
    tabLog.Enabled = True
    
    '�u���O�}�̏o�́v�t��True�ɂ���B
    cmdLog(1).Enabled = True
    
    '�u���O���k�}�̏o�́v�t��True�ɂ���B
    cmdLzhFileWrite.Enabled = True
        
    '�u�\���X�V�v�t��True�ɂ���B
    cmdUpdateDisplay.Enabled = True
    
    '�u���O�\��(�e�L�X�g�\��)�v�t��True�ɂ���B
    cmdLog(0).Enabled = True

    '�u�}�̎�O�v�t��True�ɂ���B
    cmdInstall.Enabled = True
    
    '�u�ێ��ʂ֖߂�v�t��True�ɂ���B
    cmdReturn.Enabled = True

End Sub

'EG20 V2.1.0.1 ADD START �y��-350�Ή��z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : tabTakuCorner_Click
'//  �@�\����  : �R�[�i�I���^�u�N���b�N������
'//  �@�\�T�v  : �\�����@�w���I���R�[�i�݂̂ɂ���
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@�@Index  �@[IN]�e�L�X�g�{�b�N�X�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-13   CODED   BY [TCC] M.Matsumoto
'//                 �y��-350�Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub tabTakuCorner_Click(PreviousTab As Integer)

    Dim intIndex As Integer
    Dim intStIndex As Integer
    Dim intEdIndex As Integer
    Dim intGoki As Integer
    
    intStIndex = tabTakuCorner.Tab * 16
    intEdIndex = intStIndex + 15
    
    '�\�����@�w��̃R�[�i�^�u��I���R�[�i�݂̂̕\���ɂ���
    For intIndex = 0 To tabTakuCorner.Tabs - 1
        If intIndex = tabTakuCorner.Tab Then
            tabCorner.TabVisible(intIndex) = True
            tabCorner.Tab = intIndex
        Else
            tabCorner.TabVisible(intIndex) = False
        End If
    Next
    
    '�I����ԁi�����ϐ��j�͑I���R�[�i�̍��@�̂ݗL���Ƃ���
    For intIndex = 0 To chkLogGouki.UBound
        intGoki = CInt(chkLogGouki(intIndex).Tag) - 1
        If intGoki >= 0 Then
            If intIndex >= intStIndex And intIndex <= intEdIndex Then
                    mintStatus(intGoki) = chkLogGouki(intIndex).Value
            Else
                mintStatus(intGoki) = CHECKBOX_OFF
            End If
        End If
    Next
    
End Sub
'EG20 V2.1.0.1 ADD END

