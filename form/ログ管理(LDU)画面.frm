VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLDULogkanri 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "                                                                  LD���[�e�B���e�B���O�Ǘ�"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   9480
      Top             =   3240
   End
   Begin VB.CommandButton CmdRemove 
      Caption         =   "�}�̎�O"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9240
      TabIndex        =   135
      Top             =   6540
      Width           =   2600
   End
   Begin VB.TextBox txtDummy 
      Height          =   375
      Left            =   11400
      TabIndex        =   0
      Text            =   "Text11"
      Top             =   15420
      Width           =   2175
   End
   Begin VB.CommandButton cmdLogHyouzi 
      Caption         =   "  ���O�\��     (�e�L�X�g�\��)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9240
      TabIndex        =   75
      Top             =   540
      Width           =   2600
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "���O�}�̏o��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9240
      TabIndex        =   76
      Top             =   1980
      Width           =   2600
   End
   Begin VB.CommandButton cmdSqllog 
      Caption         =   "   SQL���O     �}�̏o��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9240
      TabIndex        =   77
      Top             =   5400
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "  ���O�Ǘ�    ��ʂ֖߂�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9240
      TabIndex        =   78
      Top             =   7920
      Width           =   2600
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   8535
      Left            =   120
      TabIndex        =   79
      Top             =   420
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   15055
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "�\���t�@�C���w��"
      TabPicture(0)   =   "���O�Ǘ�(LDU)���.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LstFile"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "optApp"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "optHoshu"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdRefresh"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblFile"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblStart"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblEnd"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblSize"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "�\�����ڎw��"
      TabPicture(1)   =   "���O�Ǘ�(LDU)���.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frmMod"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frmShubetu"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frmKekka"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "frmOpt"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "frmHani"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "�\�����@�w��"
      TabPicture(2)   =   "���O�Ǘ�(LDU)���.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdHHisentaku"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdHSentaku"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdZHisentaku"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdZSentaku"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "tabCorner"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
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
         Left            =   -68280
         TabIndex        =   236
         Top             =   1200
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
         Left            =   -70440
         TabIndex        =   235
         Top             =   1200
         Width           =   2000
      End
      Begin VB.CommandButton cmdZHisentaku 
         Caption         =   "  �S�R�[�i    �S���@ ��I��"
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
         Left            =   -72600
         TabIndex        =   234
         Top             =   1200
         Width           =   2000
      End
      Begin VB.CommandButton cmdZSentaku 
         Caption         =   "  �S�R�[�i    �S���@ �I��"
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
         Left            =   -74760
         TabIndex        =   233
         Top             =   1200
         Width           =   2000
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
         Height          =   5820
         Left            =   -74640
         MultiSelect     =   2  '�g��
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   2040
         Width           =   8055
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
         Left            =   -74280
         TabIndex        =   1
         Top             =   600
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton optHoshu 
         Caption         =   "�ێ�v���O�������O"
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
         Left            =   -74280
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "���O�ؑ�"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -69240
         TabIndex        =   3
         Top             =   600
         Width           =   2295
      End
      Begin VB.Frame frmHani 
         Caption         =   "�\���͈͎w��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   85
         Top             =   360
         Width           =   8415
         Begin VB.OptionButton optHaninasi 
            Caption         =   "�͈͎w�薳"
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
            Left            =   360
            TabIndex        =   6
            Top             =   600
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optHaniari 
            Caption         =   "�͈͎w��L"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtStNen 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  '�̌Œ�
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   7
            Text            =   "9999"
            Top             =   210
            Width           =   615
         End
         Begin VB.TextBox txtStTuki 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  '�̌Œ�
            Left            =   4560
            MaxLength       =   2
            TabIndex        =   8
            Text            =   "99"
            Top             =   210
            Width           =   375
         End
         Begin VB.TextBox txtStZi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  '�̌Œ�
            Left            =   6240
            MaxLength       =   2
            TabIndex        =   10
            Text            =   "99"
            Top             =   210
            Width           =   375
         End
         Begin VB.TextBox txtStHi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  '�̌Œ�
            Left            =   5400
            MaxLength       =   2
            TabIndex        =   9
            Text            =   "99"
            Top             =   210
            Width           =   375
         End
         Begin VB.TextBox txtStFun 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  '�̌Œ�
            Left            =   7080
            MaxLength       =   2
            TabIndex        =   11
            Text            =   "99"
            Top             =   210
            Width           =   375
         End
         Begin VB.TextBox txtEdNen 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  '�̌Œ�
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   12
            Text            =   "9999"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtEdTuki 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  '�̌Œ�
            Left            =   4560
            MaxLength       =   2
            TabIndex        =   13
            Text            =   "99"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtEdZi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  '�̌Œ�
            Left            =   6240
            MaxLength       =   2
            TabIndex        =   15
            Text            =   "99"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtEdHi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  '�̌Œ�
            Left            =   5400
            MaxLength       =   2
            TabIndex        =   14
            Text            =   "99"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtEdFun 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  '�̌Œ�
            Left            =   7080
            MaxLength       =   2
            TabIndex        =   16
            Text            =   "99"
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblSt 
            Caption         =   "�J�n"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   97
            Top             =   270
            Width           =   495
         End
         Begin VB.Label lblStNen 
            Caption         =   "�N"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   96
            Top             =   270
            Width           =   255
         End
         Begin VB.Label lblStTuki 
            Caption         =   "��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5040
            TabIndex        =   95
            Top             =   270
            Width           =   255
         End
         Begin VB.Label lblStHi 
            Caption         =   "��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   94
            Top             =   270
            Width           =   255
         End
         Begin VB.Label lblStZi 
            Caption         =   "��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   93
            Top             =   270
            Width           =   255
         End
         Begin VB.Label lblStFun 
            Caption         =   "��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   92
            Top             =   270
            Width           =   255
         End
         Begin VB.Label lblEd 
            Caption         =   "�I��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   91
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lblEdNen 
            Caption         =   "�N"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   90
            Top             =   660
            Width           =   255
         End
         Begin VB.Label lblEdTuki 
            Caption         =   "��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5040
            TabIndex        =   89
            Top             =   660
            Width           =   255
         End
         Begin VB.Label lblEdHi 
            Caption         =   "��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   88
            Top             =   660
            Width           =   255
         End
         Begin VB.Label lblEdZi 
            Caption         =   "��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   87
            Top             =   660
            Width           =   255
         End
         Begin VB.Label lblEdFun 
            Caption         =   "��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   86
            Top             =   660
            Width           =   255
         End
      End
      Begin VB.Frame frmOpt 
         Caption         =   "�\���I�v�V����"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   84
         Top             =   1500
         Width           =   2775
         Begin VB.OptionButton optShousai 
            Caption         =   "�ڍו\��"
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
            Left            =   360
            TabIndex        =   18
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton optSam 
            Caption         =   "�T�}���[�\��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Frame frmKekka 
         Caption         =   "�������ʎw��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3120
         TabIndex        =   83
         Top             =   1500
         Width           =   2895
         Begin VB.CheckBox chkSeijou 
            Caption         =   "����"
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
            Left            =   360
            TabIndex        =   19
            Top             =   240
            Value           =   1  '����
            Width           =   855
         End
         Begin VB.CheckBox chkIjou 
            Caption         =   "�ُ�"
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
            Left            =   1680
            TabIndex        =   20
            Top             =   240
            Value           =   1  '����
            Width           =   855
         End
         Begin VB.CheckBox chkReigai 
            Caption         =   "��O"
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
            Left            =   360
            TabIndex        =   21
            Top             =   600
            Value           =   1  '����
            Width           =   855
         End
         Begin VB.CheckBox chkKeikoku 
            Caption         =   "�x��"
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
            Left            =   1680
            TabIndex        =   22
            Top             =   600
            Value           =   1  '����
            Width           =   855
         End
      End
      Begin VB.Frame frmShubetu 
         Caption         =   "���ڎ�ʎw��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6120
         TabIndex        =   82
         Top             =   1500
         Width           =   2535
         Begin VB.CheckBox chkKey 
            Caption         =   "�L�[����"
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
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Value           =   1  '����
            Width           =   1335
         End
         Begin VB.CheckBox chkDeb 
            Caption         =   "�f�o�b�O����"
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
            Left            =   240
            TabIndex        =   24
            Top             =   600
            Value           =   1  '����
            Width           =   1815
         End
      End
      Begin VB.Frame frmMod 
         Caption         =   "���W���[���w��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   240
         TabIndex        =   80
         Top             =   2520
         Width           =   8415
         Begin VB.CommandButton cmdModSen 
            Caption         =   "�S�đI��"
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
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdModHi 
            Caption         =   "�S�Ĕ�I��"
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
            Left            =   1800
            TabIndex        =   26
            Top             =   360
            Width           =   1455
         End
         Begin VB.Frame frmModMeisai 
            Height          =   5055
            Left            =   120
            TabIndex        =   81
            Top             =   720
            Width           =   8175
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
               Index           =   79
               Left            =   6360
               TabIndex        =   133
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
               Index           =   78
               Left            =   6360
               TabIndex        =   132
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
               Index           =   77
               Left            =   6360
               TabIndex        =   131
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
               Index           =   76
               Left            =   6360
               TabIndex        =   130
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
               Index           =   75
               Left            =   6360
               TabIndex        =   129
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
               Index           =   74
               Left            =   6360
               TabIndex        =   128
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
               Index           =   73
               Left            =   6360
               TabIndex        =   127
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
               Index           =   72
               Left            =   6360
               TabIndex        =   126
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
               Index           =   71
               Left            =   6360
               TabIndex        =   125
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
               Index           =   70
               Left            =   6360
               TabIndex        =   124
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
               Index           =   69
               Left            =   6360
               TabIndex        =   123
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
               Index           =   68
               Left            =   6360
               TabIndex        =   122
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
               Index           =   67
               Left            =   6360
               TabIndex        =   121
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
               Index           =   66
               Left            =   6360
               TabIndex        =   120
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
               Index           =   65
               Left            =   6360
               TabIndex        =   119
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
               Index           =   64
               Left            =   6360
               TabIndex        =   118
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
               Index           =   63
               Left            =   6360
               TabIndex        =   117
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
               Index           =   62
               Left            =   6360
               TabIndex        =   116
               Top             =   600
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
               Index           =   61
               Left            =   6360
               TabIndex        =   115
               Top             =   360
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
               Index           =   60
               Left            =   6360
               TabIndex        =   114
               Top             =   120
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
               Index           =   59
               Left            =   4320
               TabIndex        =   113
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
               Index           =   58
               Left            =   4320
               TabIndex        =   112
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
               Index           =   57
               Left            =   4320
               TabIndex        =   111
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
               Index           =   56
               Left            =   4320
               TabIndex        =   110
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
               Index           =   55
               Left            =   4320
               TabIndex        =   109
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
               Index           =   54
               Left            =   4320
               TabIndex        =   108
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
               Index           =   53
               Left            =   4320
               TabIndex        =   107
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
               Index           =   52
               Left            =   4320
               TabIndex        =   106
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
               Index           =   51
               Left            =   4320
               TabIndex        =   105
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
               Index           =   50
               Left            =   4320
               TabIndex        =   104
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
               Index           =   49
               Left            =   4320
               TabIndex        =   103
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
               Index           =   48
               Left            =   4320
               TabIndex        =   102
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
               Index           =   47
               Left            =   4320
               TabIndex        =   74
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
               Index           =   46
               Left            =   4320
               TabIndex        =   73
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
               Index           =   45
               Left            =   4320
               TabIndex        =   72
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
               Index           =   44
               Left            =   4320
               TabIndex        =   71
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
               Index           =   43
               Left            =   4320
               TabIndex        =   70
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
               Index           =   42
               Left            =   4320
               TabIndex        =   69
               Top             =   600
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
               Left            =   4320
               TabIndex        =   68
               Top             =   360
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
               Left            =   4320
               TabIndex        =   67
               Top             =   120
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
               Left            =   2280
               TabIndex        =   66
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
               Index           =   38
               Left            =   2280
               TabIndex        =   65
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
               Index           =   37
               Left            =   2280
               TabIndex        =   64
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
               Index           =   36
               Left            =   2280
               TabIndex        =   63
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
               Index           =   35
               Left            =   2280
               TabIndex        =   62
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
               Index           =   34
               Left            =   2280
               TabIndex        =   61
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
               Index           =   33
               Left            =   2280
               TabIndex        =   60
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
               Index           =   32
               Left            =   2280
               TabIndex        =   59
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
               Index           =   31
               Left            =   2295
               TabIndex        =   58
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
               Index           =   30
               Left            =   2295
               TabIndex        =   57
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
               Index           =   29
               Left            =   2295
               TabIndex        =   56
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
               Index           =   28
               Left            =   2295
               TabIndex        =   55
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
               Index           =   27
               Left            =   2295
               TabIndex        =   54
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
               Index           =   26
               Left            =   2295
               TabIndex        =   53
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
               Index           =   25
               Left            =   2295
               TabIndex        =   52
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
               Index           =   24
               Left            =   2295
               TabIndex        =   51
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
               Index           =   23
               Left            =   2295
               TabIndex        =   50
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
               Index           =   22
               Left            =   2295
               TabIndex        =   49
               Top             =   600
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
               TabIndex        =   48
               Top             =   360
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
               TabIndex        =   47
               Top             =   120
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
               Left            =   375
               TabIndex        =   46
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
               Index           =   18
               Left            =   375
               TabIndex        =   45
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
               Index           =   17
               Left            =   375
               TabIndex        =   44
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
               Index           =   16
               Left            =   375
               TabIndex        =   43
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
               Index           =   15
               Left            =   360
               TabIndex        =   42
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
               Index           =   14
               Left            =   360
               TabIndex        =   41
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
               Index           =   13
               Left            =   360
               TabIndex        =   40
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
               Index           =   12
               Left            =   360
               TabIndex        =   39
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
               Index           =   11
               Left            =   360
               TabIndex        =   38
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
               Index           =   10
               Left            =   360
               TabIndex        =   37
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
               Index           =   9
               Left            =   360
               TabIndex        =   36
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
               Index           =   8
               Left            =   360
               TabIndex        =   35
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
               Index           =   7
               Left            =   360
               TabIndex        =   34
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
               Index           =   6
               Left            =   360
               TabIndex        =   33
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
               Index           =   5
               Left            =   360
               TabIndex        =   32
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
               Index           =   4
               Left            =   360
               TabIndex        =   31
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
               Index           =   3
               Left            =   360
               TabIndex        =   30
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
               Index           =   2
               Left            =   360
               TabIndex        =   29
               Top             =   600
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
               Left            =   360
               TabIndex        =   28
               Top             =   360
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
               Left            =   360
               TabIndex        =   27
               Top             =   120
               Value           =   1  '����
               Visible         =   0   'False
               Width           =   1770
            End
         End
      End
      Begin TabDlg.SSTab tabCorner 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   136
         Top             =   2280
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
         TabPicture(0)   =   "���O�Ǘ�(LDU)���.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "chkEGXGoki(15)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "chkEGXGoki(14)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkEGXGoki(13)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkEGXGoki(12)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "chkEGXGoki(11)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkEGXGoki(10)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "chkEGXGoki(9)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "chkEGXGoki(8)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "chkEGXGoki(7)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "chkEGXGoki(6)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "chkEGXGoki(5)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "chkEGXGoki(4)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "chkEGXGoki(3)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "chkEGXGoki(2)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "chkEGXGoki(1)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "chkEGXGoki(0)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "  "
         TabPicture(1)   =   "���O�Ǘ�(LDU)���.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkEGXGoki(16)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "chkEGXGoki(17)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "chkEGXGoki(18)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "chkEGXGoki(19)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "chkEGXGoki(20)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "chkEGXGoki(21)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "chkEGXGoki(22)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "chkEGXGoki(23)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "chkEGXGoki(24)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "chkEGXGoki(25)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "chkEGXGoki(26)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "chkEGXGoki(27)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "chkEGXGoki(28)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "chkEGXGoki(29)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "chkEGXGoki(30)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "chkEGXGoki(31)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).ControlCount=   16
         TabCaption(2)   =   "  "
         TabPicture(2)   =   "���O�Ǘ�(LDU)���.frx":008C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkEGXGoki(32)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "chkEGXGoki(33)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "chkEGXGoki(34)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "chkEGXGoki(35)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "chkEGXGoki(36)"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "chkEGXGoki(37)"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "chkEGXGoki(38)"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "chkEGXGoki(39)"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "chkEGXGoki(40)"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "chkEGXGoki(41)"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "chkEGXGoki(42)"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "chkEGXGoki(43)"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "chkEGXGoki(44)"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "chkEGXGoki(45)"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "chkEGXGoki(46)"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).Control(15)=   "chkEGXGoki(47)"
         Tab(2).Control(15).Enabled=   0   'False
         Tab(2).ControlCount=   16
         TabCaption(3)   =   "  "
         TabPicture(3)   =   "���O�Ǘ�(LDU)���.frx":00A8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "chkEGXGoki(48)"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "chkEGXGoki(49)"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "chkEGXGoki(50)"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "chkEGXGoki(51)"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "chkEGXGoki(52)"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "chkEGXGoki(53)"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "chkEGXGoki(54)"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).Control(7)=   "chkEGXGoki(55)"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).Control(8)=   "chkEGXGoki(56)"
         Tab(3).Control(8).Enabled=   0   'False
         Tab(3).Control(9)=   "chkEGXGoki(57)"
         Tab(3).Control(9).Enabled=   0   'False
         Tab(3).Control(10)=   "chkEGXGoki(58)"
         Tab(3).Control(10).Enabled=   0   'False
         Tab(3).Control(11)=   "chkEGXGoki(59)"
         Tab(3).Control(11).Enabled=   0   'False
         Tab(3).Control(12)=   "chkEGXGoki(60)"
         Tab(3).Control(12).Enabled=   0   'False
         Tab(3).Control(13)=   "chkEGXGoki(61)"
         Tab(3).Control(13).Enabled=   0   'False
         Tab(3).Control(14)=   "chkEGXGoki(62)"
         Tab(3).Control(14).Enabled=   0   'False
         Tab(3).Control(15)=   "chkEGXGoki(63)"
         Tab(3).Control(15).Enabled=   0   'False
         Tab(3).ControlCount=   16
         TabCaption(4)   =   "  "
         TabPicture(4)   =   "���O�Ǘ�(LDU)���.frx":00C4
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "chkEGXGoki(64)"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "chkEGXGoki(65)"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "chkEGXGoki(66)"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "chkEGXGoki(67)"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "chkEGXGoki(68)"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).Control(5)=   "chkEGXGoki(69)"
         Tab(4).Control(5).Enabled=   0   'False
         Tab(4).Control(6)=   "chkEGXGoki(70)"
         Tab(4).Control(6).Enabled=   0   'False
         Tab(4).Control(7)=   "chkEGXGoki(71)"
         Tab(4).Control(7).Enabled=   0   'False
         Tab(4).Control(8)=   "chkEGXGoki(72)"
         Tab(4).Control(8).Enabled=   0   'False
         Tab(4).Control(9)=   "chkEGXGoki(73)"
         Tab(4).Control(9).Enabled=   0   'False
         Tab(4).Control(10)=   "chkEGXGoki(74)"
         Tab(4).Control(10).Enabled=   0   'False
         Tab(4).Control(11)=   "chkEGXGoki(75)"
         Tab(4).Control(11).Enabled=   0   'False
         Tab(4).Control(12)=   "chkEGXGoki(76)"
         Tab(4).Control(12).Enabled=   0   'False
         Tab(4).Control(13)=   "chkEGXGoki(77)"
         Tab(4).Control(13).Enabled=   0   'False
         Tab(4).Control(14)=   "chkEGXGoki(78)"
         Tab(4).Control(14).Enabled=   0   'False
         Tab(4).Control(15)=   "chkEGXGoki(79)"
         Tab(4).Control(15).Enabled=   0   'False
         Tab(4).ControlCount=   16
         TabCaption(5)   =   "  "
         TabPicture(5)   =   "���O�Ǘ�(LDU)���.frx":00E0
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "chkEGXGoki(80)"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).Control(1)=   "chkEGXGoki(81)"
         Tab(5).Control(1).Enabled=   0   'False
         Tab(5).Control(2)=   "chkEGXGoki(82)"
         Tab(5).Control(2).Enabled=   0   'False
         Tab(5).Control(3)=   "chkEGXGoki(83)"
         Tab(5).Control(3).Enabled=   0   'False
         Tab(5).Control(4)=   "chkEGXGoki(84)"
         Tab(5).Control(4).Enabled=   0   'False
         Tab(5).Control(5)=   "chkEGXGoki(85)"
         Tab(5).Control(5).Enabled=   0   'False
         Tab(5).Control(6)=   "chkEGXGoki(86)"
         Tab(5).Control(6).Enabled=   0   'False
         Tab(5).Control(7)=   "chkEGXGoki(87)"
         Tab(5).Control(7).Enabled=   0   'False
         Tab(5).Control(8)=   "chkEGXGoki(88)"
         Tab(5).Control(8).Enabled=   0   'False
         Tab(5).Control(9)=   "chkEGXGoki(89)"
         Tab(5).Control(9).Enabled=   0   'False
         Tab(5).Control(10)=   "chkEGXGoki(90)"
         Tab(5).Control(10).Enabled=   0   'False
         Tab(5).Control(11)=   "chkEGXGoki(91)"
         Tab(5).Control(11).Enabled=   0   'False
         Tab(5).Control(12)=   "chkEGXGoki(92)"
         Tab(5).Control(12).Enabled=   0   'False
         Tab(5).Control(13)=   "chkEGXGoki(93)"
         Tab(5).Control(13).Enabled=   0   'False
         Tab(5).Control(14)=   "chkEGXGoki(94)"
         Tab(5).Control(14).Enabled=   0   'False
         Tab(5).Control(15)=   "chkEGXGoki(95)"
         Tab(5).Control(15).Enabled=   0   'False
         Tab(5).ControlCount=   16
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   232
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   231
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   230
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   229
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   228
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   227
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   226
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   225
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   224
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   223
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   222
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   221
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   220
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   219
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   218
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   217
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   216
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   215
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   214
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   213
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   212
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   211
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   210
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   209
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   208
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   207
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   206
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   205
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   204
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   203
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   202
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   201
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   200
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   199
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   198
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   197
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   196
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   195
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   194
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   193
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   192
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   191
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   190
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   189
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   188
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   187
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   186
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   185
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   184
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   183
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   182
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   181
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   180
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   179
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   178
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   177
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   176
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   175
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   174
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   173
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   172
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   171
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   170
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   169
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   168
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   166
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   165
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   164
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   163
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   162
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   161
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   160
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   159
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   158
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   157
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   156
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   155
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   154
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   153
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   152
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   151
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   150
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   149
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   148
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   147
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   146
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   145
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   144
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   143
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   142
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   141
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   140
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   139
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   138
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkEGXGoki 
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
            TabIndex        =   137
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
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
         Left            =   -74640
         TabIndex        =   101
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblStart 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "���O�J�n����"
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
         Left            =   -73080
         TabIndex        =   100
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblEnd 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "���O�I������"
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
         Left            =   -70560
         TabIndex        =   99
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblSize 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "�T�C�Y"
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
         Left            =   -68040
         TabIndex        =   98
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H0000C000&
      Caption         =   "LDU�A�v���P�[�V�������O�Ǘ�"
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
      TabIndex        =   134
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmLDULogkanri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmLDULogKanri.frm
'//  �p�b�P�[�W���FLD���[�e�B���e�B���O�Ǘ����
'//
'//  �T�v�FLD���[�e�B���e�B���O�Ǘ����
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 �ELD���[�e�B���e�B�A���O�Ǘ����(frmLogKanri.frm)�𗬗p
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή�
'//     REVISIONS :(1.9.0.1) 2009-09-10   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή� ���b�Z�[�W�{�b�N�X�C��
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02 REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή��y�Ď�D-122�z
'//                   �E���O�}�̏o�͌��ʃ��b�Z�[�W�{�b�N�X�����ύX
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-07 REVISED BY [TCC] M.Matsumoto
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 3.6.0.1) 2012-02-23   REVISED BY [TCC] T.Furuya
'//                 EG20 �t�F�[�Y�Q �c�����
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ���O�}�̏o�͎��A������T�P�Q���Ƃ���
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l

Public iNowChk1 As Integer
Public iNowChk2 As Integer

'�A�v���E�ێ��ʒ�`
Public mlngLogType        As Long

'///////////////////////////////////////////////////////////////////
'�Ώۃt�@�C���t���p�X�i����̧�ق̎��A��߰�1�����ŋ�؂�B�j
'///////////////////////////////////////////////////////////////////
Private sObjectFiles As String   '۸�̧��ؽ��ޯ���őI�𒆂�̧�ق����߽������
Private sObjectTopFile As String '����A�I�𒆂̐擪�i�ŋ��j̧�ٖ��B(12����)�B

'///////////////////////////////////////////////////////////////////
'���O���i�[�G���A
'///////////////////////////////////////////////////////////////////
Private Type LogFileData
    sPath As String                 '���O�t�@�C���̃p�X
    sName As String                 '���O�t�@�C����
    dtFileDate As Date              '�쐬���t�E����
    dtFileDate2 As Date              '�쐬���t�E����
    lFileSize As Long               '�t�@�C���T�C�Y
    bSelect As Boolean              '�I���t���O
End Type

Private uLogfileData() As LogFileData

'///////////////////////////////////////////////////////////////////
'���W���[�����i�[�G���A
'///////////////////////////////////////////////////////////////////
Private Type ModFileData
    sName As String             '���W���[����
    sDai  As String             '�區��
    sShou As String             '������
    sType As String             '���W���[���^�C�v
    iBit  As Integer            '�r�b�g�ԍ�
End Type

Private uModFileData(79) As ModFileData
Private iModCnt As Integer

'///////////////////////////////////////////////////////////////////
'EGX���i�[�G���A
'///////////////////////////////////////////////////////////////////
Private Type EgxFileData
    iRonri As Integer               '�_�����@
    iHyozi As Integer               '�\�����@
    iIndex As Integer               'chkCorner��INDEX
End Type

Private uEgxFileData(15) As EgxFileData
Private iEgxCnt As Integer

'///////////////////////////////////////////////////////////////////
'�C�x���g���O�R�s�[�p���[�N�t�@�C�����t���p�X
'///////////////////////////////////////////////////////////////////
Private SAVEFILE_SYS As String
Private SAVEFILE_SEC As String
Private SAVEFILE_APP As String
Private SAVEFILE_LOG As String

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias _
    "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
        lpSectorsPerCluster As Long, _
        lpBytesPerSector As Long, _
        lpNumberOfFreeClusters As Long, _
        lpTtoalNumberOfClusters As Long) As Long
        
Private mintStatus(31) As Integer       'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z

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
'//  �@�\����  : LD���[�e�B���e�B���O�Ǘ����(�A�N�e�B�u��)
'//  �@�\�T�v  : ��ʂ̍őO�ʕ\�����s���B
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
'//  �@�\����  : LD���[�e�B���e�B���O�Ǘ����(�f�B�A�N�e�B�u��)
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
'//  �@�\����  : LD���[�e�B���e�B���O�Ǘ����(���[�h��)
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
'//     REVISIONS :(1.9.0.1) 2009-09-10   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή� ���b�Z�[�W�{�b�N�X�C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim intModulesFileNo As Integer
    Dim sModules As String * LDU_LOG_SIZE    '�P�s���t�@�C�����e�擾�p
    Dim Cnt As Integer
    Dim iMozi As Integer
    Dim iKbn As Integer
    Dim iRet As Integer
    Dim sType As String * LDU_LOG_MOZI_TYPE_SIZE         '�ݒu�^�C�v
    Dim sKeyName As String
    Dim str As String
    Dim iLoop As Integer
    Dim MyName As String
    Dim iErr As Integer
    'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
    Dim Cnt2 As Integer
    Dim strCorner1 As String
    Dim strCorner2 As String
    Dim intIndex As Integer
    'EG20 V2.1.0.1 ADD END
                   
    '�p�X�w��
    LDU_PROFILE_NAME_EGX = PATH_LDU_APP & LDU_EGX_FILE

    SAVEFILE_SYS = PATH_LDU_APP & PATH_LDU_WORK & "SysEvent.Evt"
    SAVEFILE_SEC = PATH_LDU_APP & PATH_LDU_WORK & "SecEvent.Evt"
    SAVEFILE_APP = PATH_LDU_APP & PATH_LDU_WORK & "AppEvent.Evt"
    SAVEFILE_LOG = PATH_LDU_APP & PATH_LDU_WORK & "drwtsn32.log"
    
    gStrCurrentForm = sFormName_LDULog
    
    cmdLogHyouzi.Caption = "���O�\��" & Chr(13) & "(�e�L�X�g�\���j"
    cmdCancel.Caption = "���O�Ǘ�" & Chr(13) & "��ʂ֖߂�"
    
    'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
'    cmdEGXZSentaku.Caption = "���D�@�S���@" & Chr(13) & "�I��"
'    cmdEGXZHisentaku.Caption = "���D�@�S���@" & Chr(13) & "��I��"
    'EG20 V2.1.0.1 DEL END

    '�z�u�ݒ�
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
         
    '������
    tabMain.Tab = 0
    
    LstFile.Clear
    
    txtStNen.Text = ""
    txtStTuki.Text = ""
    txtStHi.Text = ""
    txtStZi.Text = ""
    txtStFun.Text = ""
    txtEdNen.Text = ""
    txtEdTuki.Text = ""
    txtEdHi.Text = ""
    txtEdZi.Text = ""
    txtEdFun.Text = ""
    
   'V1.3.0.1 ADD START
   '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END
   
    'INI�t�@�C�����̐ݒ�擾
    '���W���[���w��擾
    On Error GoTo FileError
    iErr = 1
    
    '�t�@�C���L���`�F�b�N
    MyName = Dir(PATH_LDU_APP & LDU_MODULES_FILE_FULLPASS, vbNormal)
    If MyName = "" Then
        'MsgBox "DICPLOG.INI�̎擾�Ɏ��s���܂����", vbCritical 'V1.9.0.1 DEL
        GoTo FileError
    End If
    
    Cnt = 0
    
    For Cnt = 0 To 79
        sKeyName = "ID" & Format(Cnt, "000")
        iRet = GetPrivateProfileString(LDU_PROFILE_SECTION_NAME_ID, _
                                       sKeyName, _
                                       DEFAILT, sModules, Len(sModules), _
                                       PATH_LDU_APP & LDU_MODULES_FILE_FULLPASS)

        iMozi = 1
        iKbn = 1
        Do
            If Mid(sModules, iMozi, 1) = "," Then
                Select Case iKbn
                    Case 1
                        uModFileData(Cnt).sName = Left(sModules, iMozi - 1)
                        sModules = Mid(sModules, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 2
                        uModFileData(Cnt).sDai = Left(sModules, iMozi - 1)
                        sModules = Mid(sModules, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 3
                        uModFileData(Cnt).sShou = Left(sModules, iMozi - 1)
                        sModules = Mid(sModules, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 4
                        uModFileData(Cnt).sType = Left(sModules, iMozi - 1)
                        sModules = Mid(sModules, iMozi + 1)
                        iMozi = 0
                        iKbn = iKbn + 1
                    Case 5
                        uModFileData(Cnt).iBit = Left(sModules, iMozi - 1)
                        Exit Do
                End Select
            End If
            iMozi = iMozi + 1
            If iMozi > Len(sModules) Then
                Exit Do
            End If
        Loop

        If iKbn = 5 Then
            chkMod(Cnt).Visible = True
            chkMod(Cnt).Caption = uModFileData(Cnt).sName
            If LenB(StrConv(uModFileData(Cnt).sName, vbFromUnicode)) > 14 Then
                str = uModFileData(Cnt).sName
                For iLoop = 0 To Len(uModFileData(Cnt).sName)
                    str = Left(str, Len(str) - 1)
                    If LenB(StrConv(str, vbFromUnicode)) <= 14 Then
                        chkMod(Cnt).Caption = str
                        Exit For
                    End If
                Next
            End If
            If Int(uModFileData(Cnt).sShou) = 0 Then
                chkMod(Cnt).Left = chkMod(Cnt).Left - 240
            End If
            iModCnt = Cnt
        End If
    Next
    
    iErr = 2
    
'    �t�@�C���L���`�F�b�N
    MyName = Dir(LDU_PROFILE_NAME_EGX, vbNormal)
    If MyName = "" Then
        'MsgBox "EGX.INI�̎擾�Ɏ��s���܂����", vbCritical 'V1.9.0.1 DEL
        iErr = 3
        GoTo FileError
    End If
    
    'EGX���擾
    If False = GetCpuEgxInfo() Then
        GoTo FileError
    End If
   
    'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
    Call gsGetCornerName_LDU
    
    '�^�u����ݒu�R�[�i���Ƃ���
    tabCorner.Tab = 0
    
    '���W��ԏ�����
    Erase mintStatus
    
    For Cnt = 0 To UBound(gblnCornerSet)
    
        '�ݒ肠��̃R�[�i�������ɂ���
        If gblnCornerSet(Cnt) = True Then
            '�R�[�i�[���̕\��
            strCorner1 = MidB(gstrCornerName(Cnt), 1, 12)
            strCorner2 = MidB(gstrCornerName(Cnt), 13, 12)
            tabCorner.TabCaption(Cnt) = strCorner1 & vbCrLf & strCorner2
            
        End If
    
    Next Cnt
    
    '�ݒu�R�[�i�������[�v
    For Cnt = 0 To UBound(gblnCornerSet)

        If gblnCornerSet(Cnt) = False Then
            tabCorner.TabVisible(Cnt) = False
        End If

        '�ő卆�@�������[�v
        For Cnt2 = 0 To 15
            intIndex = (Cnt * 16) + Cnt2
            chkEGXGoki(intIndex).Visible = False
            chkEGXGoki(intIndex).Tag = "0"
        Next
        
        For Cnt2 = 0 To 15
            intIndex = (Cnt * 16) + (gudtSettiCorner(Cnt).intGokiNo(Cnt2) - 1)
            If gudtSettiCorner(Cnt).intGokiNo(Cnt2) > 0 Then
                chkEGXGoki(intIndex).Caption = gudtSettiCorner(Cnt).strDispGoki(Cnt2) + "���@"
                'Tag�ɑΉ����鍆�@�ԍ����L�^�i1�`32���@�j
                chkEGXGoki(intIndex).Tag = CStr(gudtSettiCorner(Cnt).intGateNo(Cnt2))
                mintStatus(gudtSettiCorner(Cnt).intGateNo(Cnt2) - 1) = CHECKBOX_ON
                chkEGXGoki(intIndex).Visible = True
                chkEGXGoki(intIndex).Value = CHECKBOX_ON
            End If
        Next Cnt2
        
    Next Cnt
    'EG20 V2.1.0.1 ADD END
       
    On Error GoTo OtherError
       
    iNowChk1 = 1
    iNowChk2 = 2
    
    '���X�g�̏����\��
    If sSetListBox = False Then
        '�uLD���[�e�B���e�B���O�Ǘ��F�A�v�����O�\���ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
        LstFile.Clear
        MsgBox "���O�ꗗ�̎擾�Ɏ��s���܂����B", vbCritical, "�\���ُ�"
    End If

   '�uLD���[�e�B���e�B���O�Ǘ��F�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LDU_LOG_KANRI_GAMEN_START, 0)

  Exit Sub
     
FileError:
    Select Case iErr
    Case 1:
        'MsgBox "CASE1", vbCritical 'V1.9.0.1 DEL
        '�uLD���[�e�B���e�B���O�Ǘ��FINI�t�@�C���ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
    Case 2:
        'MsgBox "CASE2", vbCritical  'V1.9.0.1 DEL
        '�uLD���[�e�B���e�B���O�Ǘ��FINI�t�@�C���ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
    Case 3:
        'MsgBox "CASE3", vbCritical  'V1.9.0.1 DEL
        '�uLD���[�e�B���e�B���O�Ǘ��FINI�t�@�C���ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
    Case 4:
        'MsgBox "CASE4", vbCritical  'V1.9.0.1 DEL
        '�uLD���[�e�B���e�B���O�Ǘ��FINI�t�@�C���ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
    End Select
    
    MsgBox "INI�t�@�C���̎擾�Ɏ��s���܂����", vbCritical, "�t�@�C���ُ�"

   Exit Sub
OtherError:
   '�uLD���[�e�B���e�B���O�Ǘ��F�A�v�����O�\���ُ�v���O�o��
   Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
   '���X�g�{�b�N�X�̏�����
   LstFile.Clear
   MsgBox "���O�ꗗ�̎擾�Ɏ��s���܂����B", vbCritical, "�\���ُ�"
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : GetCpuEgxInfo
'//  �@�\����  : EGX.INI���擾����
'//  �@�\�T�v  : EGX.INI�t�@�C���������擾����B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F��������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    :Boolean             [OUT]�߂�l
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function GetCpuEgxInfo() As Boolean

    Dim sKeyName As String                 '�L�[��
    Dim iRet As Integer                    'INI�t�@�C���擾�����߂�l
    Dim sEgxData As String * LDU_LOG_SIZE  '�P�s���t�@�C�����e�擾�p
    Dim i As Integer                       '���[�v�J�E���^
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim intCnrIdx As Integer    'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z

    On Error Resume Next

    iEgxCnt = -1

    'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
    Erase gudtDisp
    Erase gudtSettiCorner
    'EG20 V2.1.0.1 ADD END
    
    'EGX���擾
'    For i = 1 To 16                'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
    For i = 1 To MAX_GATE_NO        'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
       sKeyName = "egx" & Format(i, "00")
        iRet = GetPrivateProfileString(LDU_PROFILE_SECTION_NAME_EGX, _
                                       sKeyName, _
                                       DEFAILT, sEgxData, Len(sEgxData), _
                                       LDU_PROFILE_NAME_EGX)
        If iRet = 0 Then
            GetCpuEgxInfo = False
            Exit Function
        End If

        '�f�[�^�̎擾
        ReDim sFData(14)
        iFCnt = 1

        For iFLoop = 1 To Len(sEgxData)
            If Mid(sEgxData, iFLoop, 1) <> " " Or Mid(sEgxData, iFLoop, 1) <> "," Then
                iFLoop2 = iFLoop
                Do
                    iFLoop2 = iFLoop2 + 1
                    If iFLoop2 > Len(sEgxData) Then
                        sFData(iFCnt) = Mid(sEgxData, iFLoop, iFLoop2 - iFLoop)
                        iFCnt = iFCnt + 1
                        If iFCnt >= 15 Then
                            Exit For
                        End If
                        iFLoop = iFLoop2
                        Exit Do
                    End If

                    If Mid(sEgxData, iFLoop2, 1) = " " Or Mid(sEgxData, iFLoop2, 1) = "," Then
                        sFData(iFCnt) = Mid(sEgxData, iFLoop, iFLoop2 - iFLoop)
                        If Len(Trim(sFData(iFCnt))) <> 0 Then
                            iFCnt = iFCnt + 1
                            If iFCnt >= 15 Then
                                Exit For
                            End If
                        End If
                        iFLoop = iFLoop2
                        Exit Do
                    End If
                Loop
            End If
        Next

        'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
        '�ʘH��ʂ����ݒu�̎��͏�������
'        If Trim(sFData(5)) <> MISETI Then
'            iEgxCnt = iEgxCnt + 1
'            uEgxFileData(iEgxCnt).iIndex = i
'            uEgxFileData(iEgxCnt).iRonri = i
'            uEgxFileData(iEgxCnt).iHyozi = Trim(sFData(1))
'            chkEGXGoki(uEgxFileData(iEgxCnt).iIndex).Visible = True
'            chkEGXGoki(uEgxFileData(iEgxCnt).iIndex).Caption = uEgxFileData(iEgxCnt).iHyozi & "���@"
'        End If
        'EG20 V2.1.0.1 DEL END
        
        'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
        '���@���������̎�
        If sFData(5) = MISETI Then
            gudtDisp(i - 1).intJiso = JissouUmu.Mijissou                        '������
               
        '���@����������Ă��鎞
        Else
            gudtDisp(i - 1).intJiso = JissouUmu.jissou                          '����
            
            '�\�����@
            gudtDisp(i - 1).strJikaiDipNo = sFData(1)
            '�_���R�[�i�[�ԍ�
            gudtDisp(i - 1).intRonriCorner = CInt(sFData(3))
            '�_���R�[�i�[���@�ԍ�
            gudtDisp(i - 1).intRonriCornerGoki = sFData(4)
            
            '�R�[�i�ʍ��@�����J�E���g
            intCnrIdx = gudtDisp(i - 1).intRonriCorner - 1
            gudtSettiCorner(intCnrIdx).intGokiNum = gudtSettiCorner(intCnrIdx).intGokiNum + 1
            '�_�����@�ԍ��A�\�����@�ԍ����i�[
            gudtSettiCorner(intCnrIdx).intGateNo(gudtSettiCorner(intCnrIdx).intGokiNum - 1) = i - 1 + 1
            gudtSettiCorner(intCnrIdx).intGokiNo(gudtSettiCorner(intCnrIdx).intGokiNum - 1) = gudtDisp(i - 1).intRonriCornerGoki
            gudtSettiCorner(intCnrIdx).strDispGoki(gudtSettiCorner(intCnrIdx).intGokiNum - 1) = gudtDisp(i - 1).strJikaiDipNo
        End If
        'EG20 V2.1.0.1 ADD END
  Next

  GetCpuEgxInfo = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdLogHyouzi_Click
'//  �@�\����  : �u���O�\��(�e�L�X�g�\��)�v�t����������
'//  �@�\�T�v  : �I���t�@�C�����A�e�L�X�g�ɂďo�͕\������B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F�u���O�\��(�e�L�X�g�\��)�v
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
Private Sub cmdLogHyouzi_Click()
    Dim bRet As Boolean
    Dim lRetVal As Double
    Dim sCommand As String
    Dim sWriteDir As String
    Dim iObjFileNo As Integer
    Dim sFileName As String
    Dim lngErrCode As Long
    
   '�uLD���[�e�B���e�B���O�Ǘ��F���O�\���t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_TEXT_HYOUJI_BUTTOM, 0)

    '���O�����f�[�^�������`�F�b�N
    bRet = fLogSearchCheck
    If bRet = False Then
    '���O�����f�[�^�ɃG���[������ꍇ�A�����I��
        Exit Sub
    End If

    '���O�e�L�X�g�t�@�C������������
    bRet = fWriteLogtxt
    If bRet = True Then                                 '���O�e�L�X�g�t�@�C��������ɍ쐬���ꂽ�ꍇ
      '�uLD���[�e�B���e�B���O�Ǘ��F���O�e�L�X�g�t�@�C���쐬����v���O�o��
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_CREATE_TEXT_HYOUJI, 0)
      '�t�@�C���R�s�[
      sFileName = Trim(Left(LstFile.List(LstFile.ListIndex), 12))
      sFileName = PATH_LDU_APP & PATH_LDU_WORK & "\\" & Left(sFileName, Len(sFileName) - 4) & ".txt"
      
      '�t�@�C���I�[�v��
      On Error GoTo FileError
      sCommand = MN_EXE_MEMO & sFileName              '���s�R�}���h���쐬����
      lRetVal = Shell(sCommand, vbMaximizedFocus)     '�m�[�g�p�b�h���N������
      AppActivate lRetVal, True                       '�A�N�e�B�u�i�O�ʕ\���j�ɂ���
      SendKeys "{LEFT}", True
      On Error GoTo 0
       '�uLD���[�e�B���e�B���O�Ǘ��F���O�e�L�X�g�\������v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LOG_KANRI_TEXT_HYOUJI_OK, 0)
    Else
       '�uLD���[�e�B���e�B���O�Ǘ��F�o�̓f�[�^�쐬�ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_KANRI_CREATE_TEXT_ERROR, lngErrCode)
       MsgBox "�}�̏o�͂���f�[�^�̍쐬�Ɏ��s���܂����B", vbCritical, "�f�[�^�o�͎��s"
    End If
    Exit Sub

FileError:
   '�uLD���[�e�B���e�B���O�Ǘ��F���O�e�L�X�g�\�������ُ�v���O�o��
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_KANRI_TEXT_HYOUJI_ERROR, lngErrCode)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdLog_Click
'//  �@�\����  : �u���O�}�̏o�́v�t����������
'//  �@�\�T�v  : �I���t�@�C�����A�w��t�H���_�֏o�͂���B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F�u���O�}�̏o�́v
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή�
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02 REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή��y�Ď�D-122�z
'//                   �E���O�}�̏o�͌��ʃ��b�Z�[�W�{�b�N�X�����ύX
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 ���O�}�̏o�͎��A������T�P�Q���Ƃ���
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 �y�v���O���X�o�[�\���@�\�������Ή��z
'//     REVISIONS :(EG20V5.13.0.1) 2012-06-06 REVISED BY [TCC] H.Sugimoto
'//                 �y�}�̏o�̓t�H���_�쐬�Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdLog_Click()
    Dim sWriteDir
    Dim sFileName As String
    Dim dFileSize As Double
    Dim MyPath As String
    Dim MyName As String
    Dim iRet As Integer
    Dim Sekuta As Long      '�Z�N�^�i�N���X�^����j
    Dim nByte As Long       '�o�C�g���i�Z�N�^����j
    Dim Kurasuta As Long    '�t���[�N���X�^��
    Dim Drive As Long       '�h���C�u�̃N���X�^���i���v�j
    Dim FreeSpace As Double '�f�B�X�N�̋󂫗e��
    Dim lngErrCode As Long  '�G���[�R�[�h
    Dim objFso         As New FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g 'V1.6.0.1 ADD
    Dim iFileCounter As Integer  '�Ώ�̧�ِ��J�E���^    ' EG20 V5.9.0.1�y���O�I������Ή��zADD

    Dim fso As FileSystemObject     '�t�@�C���V�X�e���I�u�W�F�N�g       ' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ�
    Dim szDefLogFolder As String    ' �o�̓��O�t�H���_                  ' EG20V5.13.0.1�y�}�̏o�̓t�H���_�쐬�Ή��z�ǉ�

    '�uLD���[�e�B���e�B���O�Ǘ���ʁF���O�}�̏o�͖t�����v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_OUTPUT_BUTTOM, 0)

    Dim bFrmShow As Boolean
    bFrmShow = False

    txtDummy.SetFocus

    On Error GoTo EVENTLOG_ERROR
    If iNowChk1 = 1 Then
        MyPath = PATH_LDU_LOG & PATH_LDU_LOG_APP
    Else
        MyPath = PATH_LDU_LOG & PATH_LDU_LOG_HOSHU
    End If

    '�t�@�C���L���`�F�b�N
    Dim i
    Dim Chk
    Chk = False
    dFileSize = 0
    iFileCounter = 0                                                                            ' EG20 V5.9.0.1�y���O�I������Ή��zADD
    For i = 0 To LstFile.ListCount - 1
        If LstFile.Selected(i) Then
            Chk = True
            MyName = Trim(Left(LstFile.List(i), 12))
            MyName = Dir(MyPath & MyName, vbNormal)
            If MyName = "" Then ' ���[�v���J�n���܂��B
                MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
                Exit Sub
            End If
            dFileSize = dFileSize + FileLen(MyPath & MyName)
            iFileCounter = iFileCounter + 1                                                     ' EG20 V5.9.0.1�y���O�I������Ή��zADD
        End If
    Next

    If Chk = False Then
        '�\���t�@�C�����I������Ă��Ȃ���΁A�G���[���b�Z�[�W��\������
        MsgBox "���ڎw��Ɉُ킪����܂��B�w�肵���\�����ڂ��m�F���Ă��������B", _
               vbCritical, _
               "���ڎw��ُ�"
        Exit Sub
    End If

' EG20 V5.9.0.1�y���O�I������Ή��zADD START
    If iFileCounter > LOG_FILECNT_MAX Then
        ' �x�������\��
        MsgBox "�I�����ꂽ�t�@�C����������𒴂��܂����B" _
               & Chr(vbKeyReturn) & "�I���ł���t�@�C������[" & LOG_FILECNT_MAX & "]���܂łł��B", _
               vbOKOnly + vbCritical, _
               "�t�@�C���w��ُ�"
        Exit Sub
    End If
' EG20 V5.9.0.1�y���O�I������Ή��zADD END

DirSelect:
    '�t�H���_�w��_�C�A���O�̕\��
'    sWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", "")                         'V1.12.0.1 DEL
    sWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)    'V1.12.0.1 ADD

    If Len(sWriteDir) = 0 Then
        Exit Sub
    End If

    If UCase(Left(sWriteDir, 1)) = "A" Then
        iRet = MsgBox("�e�c��}�����Ă��������B", vbQuestion + vbOKCancel, "�}�̏����m�F")
        If iRet = vbOK Then
            frmLDULogkanri.Refresh

             '�f�B�X�N�����擾
            iRet = GetDiskFreeSpace("a:\", Sekuta, nByte, Kurasuta, Drive)

            If Drive = 0 Then
                iRet = MsgBox(" FD���}������Ă��܂���B", _
                    vbCritical, _
                    "�w��}�̏o�ُ͈�")
                GoTo DirSelect
            End If

            '�󂫗e�ʂ��擾
            FreeSpace = Sekuta * nByte * Kurasuta
            If dFileSize > FreeSpace Then
               iRet = MsgBox("�o�̓t�@�C���̃T�C�Y���w��}�̂��傫�����ߏo�͂ł��܂���B", _
                            vbCritical, _
                            "�w��}�̏o�ُ͈�")

                GoTo DirSelect
            End If
        Else
          Exit Sub
        End If

    End If
    
' EG20V5.9.0.1�y������w�E����No.10�C���Ή��z�폜�J�n
'    '�����ԍ��i�[�i�������j
'    glShoriNo = SHORI_NO.NO_MEDIA_OUT
'
'    Load frmSyorityu
'    frmSyorityu.lblLogMessage.Caption = "�}�̏o�͒�"
'    frmSyorityu.Caption = "�}�̏o�͒�"
'    frmSyorityu.Show vbModal
'    frmSyorityu.Refresh
' EG20V5.9.0.1�y������w�E����No.10�C���Ή��z�폜�I��

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_LOG_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

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

'V1.6.0.1 ADD START
    '�R�s�[��t�H���_�p�X�쐬(�w��t�H���_��LDULOG)
    sWriteDir = sWriteDir & "\" & LDU_LOGKANRI_LDULOG
    
    '�t�@�C���V�X�e���I�u�W�F�N�g����
    Set objFso = CreateObject("Scripting.FileSystemObject")

    '�R�s�[��t�H���_�̗L���m�F
    If objFso.FolderExists(sWriteDir) = False Then
    
        '�R�s�[��t�H���_�쐬
        objFso.CreateFolder (sWriteDir)
    
    End If
    
    '�t�@�C���V�X�e���I�u�W�F�N�g���
    Set objFso = Nothing
'V1.6.0.1 ADD END
    For i = 0 To LstFile.ListCount - 1
        If LstFile.Selected(i) Then
            MyName = Trim(Left(LstFile.List(i), 12))
            FileCopy MyPath & MyName, sWriteDir & "\" & MyName
        End If
    Next

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    If UCase(Left(sWriteDir, 1)) = "A" Then
        MsgBox "�e�c�o�͂͐���I�����܂����B", vbInformation + vbOKOnly, "�o�͌���"
    Else
'EG20 V2.0.1.1�y�Ď�D-122�zDEL START
'        MsgBox "�g�c�c���ꎞ�t�H���_�ւ̏o�͂͐���I�����܂����B", vbInformation + vbOKOnly, "�o�͌���"
'EG20 V2.0.1.1�y�Ď�D-122�zDEL END
'EG20 V2.0.1.1�y�Ď�D-122�zADD START
        MsgBox "����I�����܂����B", vbInformation + vbOKOnly, "�o�͌���"
'EG20 V2.0.1.1�y�Ď�D-122�zADD END
    End If
    
    '�uLD���[�e�B���e�B���O�Ǘ����O�Ǘ���ʁF���O�}�̏o�͏�������v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LOG_OUTPUT_OK, 0)

    Exit Sub

EVENTLOG_ERROR:
    'V1.6.0.1 ADD START
     '�t�@�C���V�X�e���I�u�W�F�N�g���
     Set objFso = Nothing
     '�uLD���[�e�B���e�B���O�Ǘ���ʁF�t�H���_�쐬�ُ�v
     Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_CREATE_LOGFOLDER_ERROR, 0)
    'V1.6.0.1 ADD END
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    If UCase(Left(sWriteDir, 1)) = "A" Then
        MsgBox "�e�c�o�ُ͈͂�I�����܂����B", vbCritical, "�o�͌���"
    Else
'EG20 V2.0.1.1�y�Ď�D-122�zDEL START
'        MsgBox "�g�c�c���ꎞ�t�H���_�ւ̏o�ُ͈͂�I�����܂����B", vbCritical, "�o�͌���"
'EG20 V2.0.1.1�y�Ď�D-122�zDEL END
'EG20 V2.0.1.1�y�Ď�D-122�zADD START
        MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
'EG20 V2.0.1.1�y�Ď�D-122�zADD END
    End If
    
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '�uLD���[�e�B���e�B���O�Ǘ���ʁF���O�}�̏o�͏����ُ�v
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_OUTPUT_ERROR, lngErrCode)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdRefresh_Click
'//  �@�\����  : �u���O�ؑցv�t����������
'//  �@�\�T�v  : ���O���ŐV�̏�ԂɍX�V����B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F�u���O�ؑցv
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
    Dim bRet As Boolean                     '���[�����M�����̖߂�l
    Dim udtMail As IDU_LDU_LGCHGREQ_CMD     '��ʕ\���v��
    Dim lngErrCode As Long                  '�G���[�R�[�h
    Dim bFlag As Boolean                    '���[����M�����̖߂�l
    Dim lId As Long                         '���[��ID
   
    On Error Resume Next
    
    LstFile.Clear

    '�uLD���[�e�B���e�B���O�Ǘ���ʁF���O�֖ؑt�����v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_CHANGE_BUTTOM, 0)

' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_LOG_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��

    '���O�ؑ֗v�����[����ID���ɑ��M����B
    udtMail.udtlHeader.dwId = ML_ID_IDU_LDU_LGCHGREQ_CMD
    udtMail.udtlHeader.dwSize = MlSize.IDU_LDU_LGCHGREQ_CMD
    udtMail.udtlHeader.dwProid = RHOSHU_ID
    udtMail.udtlHeader.dwSubArea = 0
    If iNowChk1 = 0 Then
        udtMail.dwLgch_Type = ML_DT_APL_LOG           ' �A�v�����O
    ElseIf iNowChk1 = 1 Then
        udtMail.dwLgch_Type = ML_DT_APL_LOG           ' �A�v�����O
    Else
        udtMail.dwLgch_Type = ML_DT_HOSHU_LOG         ' �ێ烍�O
    End If
    bRet = DssSendMail(MAIL_SLOT_LDSEI, Len(udtMail), udtMail.udtlHeader)
    If bRet = False Then
       '�u���O�ؑ֗v��CMD���M�ُ�v���O�o��
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, LOG_CHANGE_CMD_SEND, lngErrCode)
       
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
       '�v���O���X�o�[����������
       Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
       
       Exit Sub
    Else
       '�u���O�ؑ֗v��CMD���M�ُ�v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, LOG_CHANGE_CMD_SEND, 0)
    End If
  
    '���O�ؑ֗v��RES��M
    bFlag = False
    Do Until bFlag = True
        '���[����M�������s��
        lId = fMailRecieve()
        Select Case lId         '���[���h�c
        '�u�v���Z�X�I���w���v�̏ꍇ
        Case ML_ID_PROEND_ORD
             '�u�v���Z�X�I���w����M����v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
            
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
            '�v���O���X�o�[����������
            Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
            '�������I������
            Exit Sub
        '�u���O�ؑ֗v��RES�v�̏ꍇ
        Case ML_ID_IDU_LDU_LGCHGREQ_RES
            '�u���O�ؑ֗v��RES��M����v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, LOG_CHANGE_RES_RECV, 0)
            '���[�v�𔲂���
            Exit Do
        Case Else
        End Select
        Sleep (MN_MAIL_INTERVAL)
    Loop
    If sSetListBox = False Then
        '���X�g�{�b�N�X�̏�����
        LstFile.Clear
        
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
        MsgBox "���O�ꗗ�̎擾�Ɏ��s���܂����B", vbCritical, "�\���ُ�"
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    Else
        '�v���O���X�o�[����������
        Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    End If
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
   '�uLD���[�e�B���e�B���O�Ǘ���ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LDU_LOG_KANRI_GAMEN_END, 0)
    frmLogMenu.ZOrder
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : optApp_Click
'//  �@�\����  : ���W�I�t�F�A�v���P�[�V�������O�I��������
'//  �@�\�T�v  : �\�����X�V����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(3.5.0.1) 2012-02-17   REVISED BY [TCC] T.Furuya
'//                 EG20 �t�F�C�Y2 �T�C�N��3 �c�����C
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub optApp_Click()

    On Error GoTo Err_mgs

   '�uLD���[�e�B���e�B���O�Ǘ���ʁF�A�v�����O�v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_CHANGE_APLLOG, 0)
    
    '�I������Ă����̂��A�ێ�v���O�������O�������ꍇ
    If iNowChk1 <> 1 Then
        
        '�I������Ă���`�F�b�N��ێ�����
        iNowChk1 = 1
                              
        '��\���ɂȂ��Ă������ڂ�\��������
        'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
'        frmMod.Visible = True
'        frmEGXGokiSentaku.Visible = True
'        cmdEGXZSentaku.Visible = True
'        cmdEGXZHisentaku.Visible = True
        'EG20 V2.1.0.1 DEL END
        
        'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
        frmMod.Visible = True                           'V3.5.0.1 ADD
        tabCorner.Visible = True
        cmdZSentaku.Visible = True
        cmdZHisentaku.Visible = True
        cmdHSentaku.Visible = True
        cmdHHisentaku.Visible = True
        'EG20 V2.1.0.1 ADD END
        
        '�\�����ēǂݍ��݂���
        If sSetListBox = False Then
            '�uLD���[�e�B���e�B���O�Ǘ���ʁF�A�v�����O�\���ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
            '���X�g�{�b�N�X�̏�����
            LstFile.Clear
            MsgBox "���O�ꗗ�̎擾�Ɏ��s���܂����B", vbCritical, "�\���ُ�"
        End If
    End If
    
   '�uLD���[�e�B���e�B���O��ʁF�A�v�����O�\������v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_APLLOG_OK, 0)
 
   Exit Sub
    
Err_mgs:
    '�uLD���[�e�B���e�B���O�Ǘ���ʁF�A�v�����O�\���ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
    
    '���X�g�{�b�N�X�̏�����
    LstFile.Clear
    MsgBox "���O�ꗗ�̎擾�Ɏ��s���܂����B", vbCritical, "�\���ُ�"
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : optHoshu_Click
'//  �@�\����  : ���W�I�t�F�ێ�v���O�������O�I��������
'//  �@�\�T�v  : �\�����X�V����B
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
Private Sub optHoshu_Click()
    
    On Error GoTo Err_mgs
    
   '�uLD���[�e�B���e�B���O�Ǘ���ʁF�ێ烍�O�v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_CHANGE_HOSHULOG, 0)
    
    '�I������Ă����̂��A�A�v���P�[�V�������O�������ꍇ
    If iNowChk1 <> 2 Then
            
        '�I������Ă���`�F�b�N��ێ�����
        iNowChk1 = 2
        
        '�\���ɂȂ��Ă��鍀�ڂ��\���ɂ���
        frmMod.Visible = False
        'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
'        cmdEGXZSentaku.Visible = False
'        cmdEGXZHisentaku.Visible = False
'        frmEGXGokiSentaku.Visible = False
        'EG20 V2.1.0.1 DEL END
        
        'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
        tabCorner.Visible = False
        cmdZSentaku.Visible = False
        cmdZHisentaku.Visible = False
        cmdHSentaku.Visible = False
        cmdHHisentaku.Visible = False
        'EG20 V2.1.0.1 ADD END
        
        '�\�����ēǂݍ��݂���
        If sSetListBox = False Then
            '�uLD���[�e�B���e�B���O�Ǘ���ʁF�ێ烍�O�\���ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_HODHULOG_ERROR, 0)
            '���X�g�{�b�N�X�̏�����
            LstFile.Clear
            MsgBox "���O�ꗗ�̎擾�Ɏ��s���܂����B", vbCritical, "�\���ُ�"
        End If
    End If
    
    '�uLD���[�e�B���e�B���O�Ǘ���ʁF�ێ烍�O�\������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_HODHULOG_OK, 0)
    
    Exit Sub

Err_mgs:
    '�uLD���[�e�B���e�B���O�Ǘ���ʁF�ێ烍�O�\���ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_HODHULOG_ERROR, 0)
    '���X�g�{�b�N�X�̏�����
    LstFile.Clear
    MsgBox "���O�ꗗ�̎擾�Ɏ��s���܂����B", vbCritical, "�\���ُ�"
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : optHaniari_Click
'//  �@�\����  : ���W�I�t�F�\���͈͎w��L�I��������
'//  �@�\�T�v  : �\�����X�V����B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
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
Private Sub optHaniari_Click()
   
    '�I������Ă����̂��A�\���͈͎w�薳�������ꍇ
    If iNowChk2 = 2 Then
    
        '�J�n�ƏI������͉\�ɂ���
        lblSt.Enabled = True
        lblStNen.Enabled = True
        lblStTuki.Enabled = True
        lblStHi.Enabled = True
        lblStZi.Enabled = True
        lblStFun.Enabled = True
        
        lblEd.Enabled = True
        lblEdNen.Enabled = True
        lblEdTuki.Enabled = True
        lblEdHi.Enabled = True
        lblEdZi.Enabled = True
        lblEdFun.Enabled = True
        
        txtStNen.Enabled = True
        txtStTuki.Enabled = True
        txtStHi.Enabled = True
        txtStZi.Enabled = True
        txtStFun.Enabled = True
        
        txtEdNen.Enabled = True
        txtEdTuki.Enabled = True
        txtEdHi.Enabled = True
        txtEdZi.Enabled = True
        txtEdFun.Enabled = True
        
        '�I������Ă���`�F�b�N��ێ�����
        iNowChk2 = 1
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : optHaninasi_Click
'//  �@�\����  : ���W�I�t�F�\���͈͎w�薳�I��������
'//  �@�\�T�v  : �\�����X�V����B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
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
Private Sub optHaninasi_Click()
    
    '�I������Ă����̂��A�\���͈͎w��L�������ꍇ
    If iNowChk2 = 1 Then
        
        '�J�n�ƏI������͉\�ɂ���
        lblSt.Enabled = False
        lblStNen.Enabled = False
        lblStTuki.Enabled = False
        lblStHi.Enabled = False
        lblStZi.Enabled = False
        lblStFun.Enabled = False
        
        lblEd.Enabled = False
        lblEdNen.Enabled = False
        lblEdTuki.Enabled = False
        lblEdHi.Enabled = False
        lblEdZi.Enabled = False
        lblEdFun.Enabled = False
        
        txtStNen.Enabled = False
        txtStTuki.Enabled = False
        txtStHi.Enabled = False
        txtStZi.Enabled = False
        txtStFun.Enabled = False
        
        txtEdNen.Enabled = False
        txtEdTuki.Enabled = False
        txtEdHi.Enabled = False
        txtEdZi.Enabled = False
        txtEdFun.Enabled = False
        
        '�I������Ă���`�F�b�N��ێ�����
        iNowChk2 = 2
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdModSen_Click
'//  �@�\����  : �u�S�đI���v�t����������
'//  �@�\�T�v  : �\�����X�V����B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u���W���[���w��v
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
Private Sub cmdModSen_Click()
    Dim iCnt As Integer
        
    For iCnt = CNT_MIN To iModCnt
        chkMod(iCnt).Value = 1
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdModHi_Click
'//  �@�\����  : �u�S�Ĕ�I���v�t����������
'//  �@�\�T�v  : �\�����X�V����B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u���W���[���w��v
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
Private Sub cmdModHi_Click()
    Dim iCnt As Integer
    
    For iCnt = CNT_MIN To iModCnt
        chkMod(iCnt).Value = 0
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : chkMod_Click
'//  �@�\����  : �e�`�F�b�N�{�b�N�X����������
'//  �@�\�T�v  : �\�����X�V����B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u���W���[���w��v
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
Private Sub chkMod_Click(Index As Integer)
    Dim iCnt As Integer
    Dim sDai As String
    Dim iChkType As Integer
        
   '�����ڒl�`�F�b�N���s���B
    If Int(uModFileData(Index).sShou) = 0 Then
        '�����ڒl���A�`�F�b�N�{�b�N�X�ő�l�̏ꍇ�͏����I��
        If Index = iModCnt Then
            Exit Sub
        End If
        
       '�����l�ݒ�
        '�Ώە��ނɘA�Ȃ���̂̃C���f�b�N�X���擾����B
        iCnt = Index + 1
        '�區�ڔԍ����擾����B
        sDai = uModFileData(Index).sDai
        '�區�ڂ̃`�F�b�N�{�b�N�X�l���擾����B
        iChkType = chkMod(Index).Value
        Do
           '�Ώە��ނ̑區�ڔԍ��ƁA�������ނ̑區�ڔԍ���v���邩�`�F�b�N����B
            If sDai = uModFileData(iCnt).sDai Then
               '��v�����ꍇ�A�������ނ̃`�F�b�N�{�b�N�X�l���A���f����B
                chkMod(iCnt).Value = iChkType
            Else
                Exit Do
            End If
            '���̕��ނɐi�ށB
            iCnt = iCnt + 1
            If iCnt > iModCnt Then
              '�`�F�b�N�{�b�N�X�̍ő�l�ɂȂ����ꍇ�͏����I��
                Exit Sub
            End If
        Loop
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtStNen_LostFocus
'//  �@�\����  : �J�n�N���͎�����
'//  �@�\�T�v  : ���͊J�n�N�������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
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
Private Sub txtStNen_LostFocus()
    Dim iRet
    '�������`�F�b�N
    iRet = TextTime_Check("Year", txtStNen.Text)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtStTuki_LostFocus
'//  �@�\����  : �J�n�����͎�����
'//  �@�\�T�v  : ���͊J�n���������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
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
Private Sub txtStTuki_LostFocus()
    Dim iRet
    '�������`�F�b�N
    iRet = TextTime_Check("Month", txtStTuki.Text)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtStHi_LostFocus
'//  �@�\����  : �J�n�����͎�����
'//  �@�\�T�v  : ���͊J�n���������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
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
Private Sub txtStHi_LostFocus()
    Dim iRet
    '�������`�F�b�N
    iRet = TextTime_Check("Day", txtStHi.Text)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtStZi_LostFocus
'//  �@�\����  : �J�n�����͎�����
'//  �@�\�T�v  : ���͊J�n���������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-06  REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub txtStZi_LostFocus()
    Dim iRet
    '�������`�F�b�N
    iRet = TextTime_Check("Hour", txtStZi.Text)

'EG20 V2.0.1.1 ADD START
    If Len(Trim(txtStNen.Text)) <> 0 And _
       Len(Trim(txtStTuki.Text)) <> 0 And _
       Len(Trim(txtStHi.Text)) <> 0 And _
       Len(Trim(txtStZi.Text)) = 0 Then
    
        iRet = MsgBox("�\���͈͂̊J�n�ɖ����͂̍��ڂ�����܂��B", vbExclamation, "���ُ͈�")
     End If
'EG20 V2.0.1.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtStFun_LostFocus
'//  �@�\����  : �J�n�����͎�����
'//  �@�\�T�v  : ���͊J�n���������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-06  REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub txtStFun_LostFocus()
    Dim iRet
    '�������`�F�b�N
    iRet = TextTime_Check("Minutes", txtStFun.Text)

'EG20 V2.0.1.1 ADD START
    If Len(Trim(txtStNen.Text)) <> 0 And _
       Len(Trim(txtStTuki.Text)) <> 0 And _
       Len(Trim(txtStHi.Text)) <> 0 And _
       Len(Trim(txtStZi.Text)) <> 0 And _
       Len(Trim(txtStFun.Text)) = 0 Then
       
        iRet = MsgBox("�\���͈͂̊J�n�ɖ����͂̍��ڂ�����܂��B", vbExclamation, "���ُ͈�")
     End If
'EG20 V2.0.1.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtEdNen_LostFocus
'//  �@�\����  : �I���N���͎�����
'//  �@�\�T�v  : ���͏I���N�������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
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
Private Sub txtEdNen_LostFocus()
    Dim iRet
    '�������`�F�b�N
    iRet = TextTime_Check("Year", txtEdNen.Text)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtEdTuki_LostFocus
'//  �@�\����  : �I�������͎�����
'//  �@�\�T�v  : ���͏I�����������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
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
Private Sub txtEdTuki_LostFocus()
    Dim iRet
    '�������`�F�b�N
    iRet = TextTime_Check("Month", txtEdTuki.Text)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtEdHi_LostFocus
'//  �@�\����  : �I�������͎�����
'//  �@�\�T�v  : ���͏I�����������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
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
Private Sub txtEdHi_LostFocus()
    Dim iRet
    '�������`�F�b�N
    iRet = TextTime_Check("Day", txtEdHi.Text)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtedZi_LostFocus
'//  �@�\����  : �I�������͎�����
'//  �@�\�T�v  : ���͏I�����������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-06  REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub txtedZi_LostFocus()
    Dim iRet
    '�������`�F�b�N
    iRet = TextTime_Check("Hour", txtEdZi.Text)

'EG20 V2.0.1.1 ADD START
    If Len(Trim(txtEdNen.Text)) <> 0 And _
       Len(Trim(txtEdTuki.Text)) <> 0 And _
       Len(Trim(txtEdHi.Text)) <> 0 And _
       Len(Trim(txtEdZi.Text)) = 0 Then
    
        iRet = MsgBox("�\���͈͂̏I���ɖ����͂̍��ڂ�����܂��B", vbExclamation, "���ُ͈�")
        '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
     End If
'EG20 V2.0.1.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtedFun_LostFocus
'//  �@�\����  : �I�������͎�����
'//  �@�\�T�v  : ���͏I�����������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-06  REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή�
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub txtedFun_LostFocus()
    Dim iRet
    '�������`�F�b�N
    iRet = TextTime_Check("Minutes", txtEdFun.Text)

'EG20 V2.0.1.1 ADD START
    If Len(Trim(txtEdNen.Text)) <> 0 And _
       Len(Trim(txtEdTuki.Text)) <> 0 And _
       Len(Trim(txtEdHi.Text)) <> 0 And _
       Len(Trim(txtEdZi.Text)) <> 0 And _
       Len(Trim(txtEdFun.Text)) = 0 Then
    
        iRet = MsgBox("�\���͈͂̏I���ɖ����͂̍��ڂ�����܂��B", vbExclamation, "���ُ͈�")
        '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
     End If
'EG20 V2.0.1.1 ADD END

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : TextTime_Check
'//  �@�\����  : �J�n/�I���N�����������͐������`�F�b�N������
'//  �@�\�T�v  : ���͂��ꂽ�l�̐������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\�����ڎw�蕔�F�u�\���͈͎w��v
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-05 REVISED BY [TCC] T.Koyama
'//                �d�f�Q�O�t�F�[�Y�Q�Ή�
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-14 REVISED BY [TCC] M.Matsumoto
'//                �y��-341�Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function TextTime_Check(sType As String, sTxt As String)
    Dim iChk As Integer
    Dim iRet As Integer
    Dim sChk As String
    
    Dim k As Integer                                'EG20 V2.0.1.1 ADD
    
    '�߂�l�Ɉُ���Z�b�g
    TextTime_Check = False
        
    If Trim(sTxt) <> "" Then
    
        'EG20 V2.0.1.1 ADD START �y��-341�Ή��z
        '���͂��ꂽ���ɐ��l�ȊO�̕��������݂���ꍇ�́A�G���[
        For k = 1 To Len(sTxt)
            If Not Mid(sTxt, k, 1) Like "[0-9]" Then
                iRet = MsgBox("���͂��ꂽ�����͐����ł͂���܂���B", vbExclamation, "���ُ͈�")
                Exit Function
            End If
        Next k
        'EG20 V2.0.1.1 ADD END
            
        iChk = Val(sTxt)
        'EG20 V2.0.1.1 DEL START �y��-341�Ή��z
'        If iChk = 0 And sType <> "Hour" And sType <> "Minutes" Then
'            iRet = MsgBox("���͂��ꂽ�����͐����ł͂���܂���B", vbExclamation, "���ُ͈�")
'            Exit Function
'        Else
        'EG20 V2.0.1.1 DEL END
            '�����O�̎��̃`�F�b�N�i�N�ȊO�j
            If sType <> "Year" Then
                sChk = Left(sTxt, 1)
                If Len(sTxt) = 2 And sChk = "0" Then
                    sTxt = Right(sTxt, 1)
                End If
            End If
            
            'EG20 V2.0.1.1 DEL START �y��-341�Ή��z
            '����
'            If Len(Trim(str(iChk))) <> Len(sTxt) Then
'                iRet = MsgBox("���͂��ꂽ�����ɐ����ȊO�̂��̂��܂܂�Ă��܂��B", vbExclamation, "���ُ͈�")
'                Exit Function
'            End If
            'EG20 V2.0.1.1 DEL END
            
            '�͈̓`�F�b�N
            Select Case sType
                Case "Year"
                    '�N
'                    If iChk < 1980 Or iChk > 2079 Then         'EG20 V2.0.1.1 DEL
                    If iChk < 2000 Or iChk > 2037 Then          'EG20 V2.0.1.1 ADD
                        iRet = MsgBox("�N�w��͈̔͂𒴂��Ă��܂��B", vbExclamation, "���ُ͈�")
                        Exit Function
                    End If
                Case "Month"
                    '��
                    If iChk < 1 Or iChk > 12 Then
                        iRet = MsgBox("���w��͈̔͂𒴂��Ă��܂��B", vbExclamation, "���ُ͈�")
                        Exit Function
                    End If
                Case "Day"
                    '��
                    If iChk < 1 Or iChk > 31 Then
                        iRet = MsgBox("���w��͈̔͂𒴂��Ă��܂��B", vbExclamation, "���ُ͈�")
                        Exit Function
                    End If
                Case "Hour"
                    '��
                    If iChk < 0 Or iChk > 23 Then
                        iRet = MsgBox("���Ԏw��͈̔͂𒴂��Ă��܂��B", vbExclamation, "���ُ͈�")
                        Exit Function
                    End If
                Case "Minutes"
                    '��
                    If iChk < 0 Or iChk > 59 Then
                        iRet = MsgBox("���Ԏw��͈̔͂𒴂��Ă��܂��B", vbExclamation, "���ُ͈�")
                        Exit Function
                    End If
            End Select
        End If
'    End If         'EG20 V2.0.1.1 DEL �y��-341�Ή��z
    
    '�߂�l�ɐ����Ԃ�
    TextTime_Check = True
    Exit Function
End Function

'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : cmdEGXZSentaku_Click
''//  �@�\����  : �u���D�@�S���@�I���v�t����������
''//  �@�\�T�v  : �\�����X�V����B
''//�@�@�@�@�@�@�@�\�����@�w�蕔�F
''//
''//              �^        ����      �Ӗ�
''//  ����      : �Ȃ�
''//
''//              �^        �l        �Ӗ�
''//  �߂�l    : �Ȃ�
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Private Sub cmdEGXZSentaku_Click()
'    Dim iCnt As Integer
'
'    For iCnt = 1 To 18
'        chkEGXGoki(iCnt).Value = 1
'    Next
'End Sub
'
''///////////////////////////////////////////////////////////////////
''//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
''//
''//  �֐�����  : cmdEGXZHisentaku_Click
''//  �@�\����  : �u���D�@�S���@��I���v�t����������
''//  �@�\�T�v  : �\�����X�V����B
''//�@�@�@�@�@�@�@�\�����@�w�蕔�F
''//
''//              �^        ����      �Ӗ�
''//  ����      : �Ȃ�
''//
''//              �^        �l        �Ӗ�
''//  �߂�l    : �Ȃ�
''//
''//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
''//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
''//  ���l�F
''///////////////////////////////////////////////////////////////////
'Private Sub cmdEGXZHisentaku_Click()
'    Dim iCnt As Integer
'
'    For iCnt = 1 To 18
'        chkEGXGoki(iCnt).Value = 0
'    Next
'End Sub
'EG20 V2.1.0.1 DEL END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSetListBox
'//  �@�\����  : ���O�t�@�C���o�^����
'//  �@�\�T�v  : ���O�t�@�C�������X�g�{�b�N�X�ɓo�^����B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F��������
'//�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�u���O�ؑցv�t����������
'//                                  ���W�I�t�F�A�v��/�ێ烍�O�I����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y����� ����No.36�֘A�z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sSetListBox()
    Dim i As Integer
    Dim iCnt As Integer
    Dim strSQL As String
    Dim datWork As Date
    Dim sEntry As String            '�ҏW������
    Dim bRet As Integer
    
    Dim strLogFilePath As String    '�A�v��or�ێ烍�O�t�@�C���p�X
    Dim strLogListPath As String    '�\���p���O���X�g�쐬�p�t�@�C���p�X
    Dim strLogFileName As String    '�A�v��or�ێ烍�O�t�@�C����
    Dim strConnectString As String  '�\���p���O���X�gCSV�t�@�C���ڑ��p
    Dim intLogFileNo As Integer     '���O���X�g�t�@�C���ԍ��w��p
    Dim strLogData() As String      'CSV����ǂݍ��񂾃��O�f�[�^�i�[�p�z��
    Dim strLineCount As String      'CSV�t�@�C���s���Ǎ�
    Dim j As Integer                'CSV�t�@�C���s���J�E���g�p
    Dim lErrSts As Long
    Dim lngErrCode As Long          '�G���[�R�[�h
    
    On Error GoTo Err_mgs
    
    sSetListBox = False
    
    On Error Resume Next            ' �G���[�̃g���b�v�𗯕ۂ��܂��B
    Err.Clear

    
    '���O�ꗗ�̃��X�g���쐬����
    '�A�v���A�ێ�`�F�b�N
    Select Case iNowChk1
        Case 1 '�A�v�����������̏���
            strLogFilePath = PATH_LDU_LOG & PATH_LDU_LOG_APP '�����F�A�v�����O�̃t�@�C���p�X
            strLogListPath = PATH_LDU_APP & PATH_LDU_WORK & LDU_APL_LOG_LIST '�A�v�����O�ꗗCSV�̃t�@�C���p�X
        Case 2 '�ێ烍�O���������̏���
            strLogFilePath = PATH_LDU_LOG & PATH_LDU_LOG_HOSHU '�����F�ێ烍�O�̃t�@�C���p�X
            strLogListPath = PATH_LDU_APP & PATH_LDU_WORK & LDU_HOSHU_LOG_LIST '�ێ烍�O�ꗗCSV�̃t�@�C���p�X
        Case Else
            Exit Function
    End Select

    bRet = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    '// �ێ��p�֐�:�\���p���O�ꗗ���X�g�쐬
    '///////////////////////////////////////////////////////////////////////////////////////////
    bRet = dllCreateLDULogInfo(lErrSts, strLogFilePath, strLogListPath)

    '���O���X�g(CSV)�쐬����
    If bRet Then
       '�uLD���[�e�B���e�B���O�Ǘ���ʁF���O���X�g�t�@�C���쐬����v
        Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LDU_LOG_KANRI_CREATE_LOGLISTFILE_OK, 0)
    '���O���X�g(CSV)�쐬���s
    Else
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
       '�uLD���[�e�B���e�B���O�Ǘ���ʁF���O���X�g�t�@�C���쐬�ُ�v
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LDU_LOG_KANRI_CREATE_LOGLISTFILE_ERROR, 0)
        Exit Function
    End If

    'CSV�t�@�C���̗L���m�F
    If Len(Trim(Dir(strLogListPath))) = 0 Then
        Exit Function
    End If

    'CSV�t�@�C���ԍ����擾����B
    intLogFileNo = FreeFile

    'CSV�t�@�C���I�[�v��
    Open strLogListPath For Input As #intLogFileNo


    'CSV�t�@�C���s���J�E���g�i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1�폜
        Do While Not EOF(intLogFileNo)                      ' EG20 V3.3.0.1�ǉ�
            Line Input #intLogFileNo, strLineCount
            j = j + 1
        Loop

    'CSV�t�@�C���N���[�Y
    Close #intLogFileNo

    'CSV�t�@�C���Ǎ��p�z��i�s�����j
    ReDim strLogData(j)

    'CSV�t�@�C���ԍ����擾����B
    intLogFileNo = FreeFile

    'CSV�t�@�C���I�[�v��
    Open strLogListPath For Input As #intLogFileNo


    '���X�g�\�����ǂݍ��݁i�t�@�C���I�[�܂Ń��[�v���J��Ԃ��j
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1�폜
        Do While Not EOF(intLogFileNo)                      ' EG20 V3.3.0.1�ǉ�
            Line Input #intLogFileNo, strLogData(i)
            i = i + 1
        Loop

    'CSV�t�@�C���N���[�Y
    Close #intLogFileNo

    iCnt = i


    '�G���[����
    On Error GoTo Err_mgs

    '�u���O�t�@�C���v���X�g�{�b�N�X���N���A����
    LstFile.Clear

    '���O�t�@�C������ҏW����
    For i = 0 To iCnt - 1
        sEntry = Left(strLogData(i), 12)  '�t�@�C����
            sEntry = sEntry & "  " & Mid(strLogData(i), 14, 19) '���O�J�n�N������\��

            sEntry = sEntry & "  " & Mid(strLogData(i), 34, 19) '���O�I���N������\��

            sEntry = sEntry & "  " & Format(Mid(strLogData(i), 54), "@@@@@@@@") '���O�T�C�Y��\��
        LstFile.AddItem sEntry
    Next

    If iCnt > 0 Then                '���O�t�@�C�������݂���
        LstFile.ListIndex = 0        '��s�ڂɃC���f�b�N�X���Z�b�g
    End If

    sSetListBox = True

    Exit Function

Err_mgs:
    '�A�v���A�ێ�`�F�b�N
    Select Case iNowChk1
        Case 1
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            '�uLD���[�e�B���e�B���O�Ǘ���ʁF�t�@�C���A�N�Z�X�ُ�v
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, 0)
        Case 2
            lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
            '�uLD���[�e�B���e�B���O�Ǘ���ʁF�t�@�C���A�N�Z�X�ُ�v
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, 0)
    End Select
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fLogSearchCheck
'//  �@�\����  : ���O�����f�[�^�`�F�b�N����
'//  �@�\�T�v  : ���O�����f�[�^�̐������`�F�b�N���s���B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F�u���O�\��(�e�L�X�g�\��)�v�t����������
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 3.6.0.1) 2012-02-23   REVISED BY [TCC] T.Furuya
'//                 EG20 �t�F�[�Y�Q �c�����
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function fLogSearchCheck() As Boolean
    Dim i As Integer            '�J�E���^
    Dim j As Integer            '�R���g���[���z��
    Dim bFlag As Boolean        '�t���O
    Dim iSelectedLines As Integer '���X�g�{�b�N�X�őI�𒆂̍s��
    Dim iChk As Integer
    Dim iChk2 As Integer
    Dim sChk As String
    Dim sFileName As String
    Dim sStAll As String
    Dim sEdAll As String
    Dim dStAll As Double
    Dim dEdAll As Double
    Dim sChkDate As String
    Dim sTxt As String
    Dim iRet

    fLogSearchCheck = False     '�߂�l�ɏ����l�Ƃ��ăG���[���Z�b�g
   
    '�t�@�C���I�𐔃`�F�b�N
    iChk = 0
    For i = 0 To LstFile.ListCount - 1
        If LstFile.Selected(i) Then
            iChk = iChk + 1
        End If
    Next
    
    If iChk = 0 Then
        '�\���t�@�C�����I������Ă��Ȃ���΁A�G���[���b�Z�[�W��\������
        MsgBox "���ڎw��Ɉُ킪����܂��B�w�肵���\�����ڂ��m�F���Ă��������B", _
               vbCritical, _
               "���ڎw��ُ�"
               Exit Function
    ElseIf iChk > 1 Then
        '�����t�@�C�����I������Ă��Ă��A�G���[���b�Z�[�W��\������
        MsgBox "�����t�@�C���w��A�L���r�l�b�g�t�@�C���ȊO�̃t�@�C���w��͂ł��܂���B", _
               vbCritical, _
               "�t�@�C���w��ُ�"
        Exit Function
    End If
    
    '�t�@�C�����̎擾
    sFileName = Trim(Left(LstFile.List(LstFile.ListIndex), 12))
    
    '�g���q�`�F�b�N
    If LCase(Right(sFileName, 3)) <> "idu" Then
        '�L���r�l�b�g�t�@�C���ȊO���w�肳�ꂽ�ꍇ�A�G���[���b�Z�[�W��\������
        MsgBox "�����t�@�C���w��A�L���r�l�b�g�t�@�C���ȊO�̃t�@�C���w��͂ł��܂���B", _
               vbCritical, _
               "�t�@�C���w��ُ�"
        Exit Function
    End If
    
    '�������ʎw��
    '����
    If chkSeijou.Value = 0 Then
        '�ُ�
        If chkIjou.Value = 0 Then
            '��O
            If chkReigai.Value = 0 Then
                '�x��
                If chkKeikoku.Value = 0 Then
                    MsgBox "���ڎw��Ɉُ킪����܂��B�w�肵���\�����ڂ��m�F���Ă��������B", vbCritical, "���ڎw��ُ�"
                    Exit Function
                End If
            End If
        End If
    End If
    
    '���ڎ�ʎw��
    '�L�[����
    If chkKey.Value = 0 Then
        '�f�o�b�O����
        If chkDeb.Value = 0 Then
            MsgBox "���ڎw��Ɉُ킪����܂��B�w�肵���\�����ڂ��m�F���Ă��������B", vbCritical, "���ڎw��ُ�"
            Exit Function
        End If
    End If
    
    
    '�Ώێ����̐������`�F�b�N
    '�͈͎w�肪����̏ꍇ�̂݃`�F�b�N����
    If optHaniari.Value = True Then
        '�J�n�`�F�b�N
        '���O�f�[�^�Ώێ����̐������`�F�b�N
        If Len(Trim(txtStNen.Text)) = 0 And _
           Len(Trim(txtStTuki.Text)) = 0 And _
           Len(Trim(txtStHi.Text)) = 0 And _
           Len(Trim(txtStZi.Text)) = 0 And _
           Len(Trim(txtStFun.Text)) = 0 Then
           
           '�S�Ė����͂Ȃ�0���Z�b�g
            sStAll = "0"
    
        ElseIf Len(Trim(txtStNen.Text)) = 0 Or _
           Len(Trim(txtStTuki.Text)) = 0 Or _
           Len(Trim(txtStHi.Text)) = 0 Then
           
           '�J�n�����ɖ����͂̍��ڂ�����Ȃ�
            MsgBox "�\���͈͂̊J�n�ɖ����͂̍��ڂ�����܂��B", _
                   vbExclamation, _
                   "���ُ͈�"               ' EG20 V3.6.0.1 ADD
'                   "�����w��ُ�"          ' EG20 V3.6.0.1 DEL
            Exit Function
        
        Else
            '�J�n�N�`�F�b�N
            iRet = TextTime_Check("Year", txtStNen.Text)
            If iRet = False Then
                Exit Function
            End If
            
            '�J�n���`�F�b�N
            iRet = TextTime_Check("Month", txtStTuki.Text)
            If iRet = False Then
                Exit Function
            End If
            
            '�J�n���`�F�b�N
            iRet = TextTime_Check("Day", txtStHi.Text)
            If iRet = False Then
                Exit Function
            End If
            
            '�J�n���`�F�b�N
            If Len(Trim(txtStZi.Text)) = 0 Then
'EG20 V2.0.1.1 DEL START
'                txtStZi.Text = "00"
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
                If Len(Trim(txtStNen.Text)) = 0 And _
                   Len(Trim(txtStTuki.Text)) = 0 And _
                   Len(Trim(txtStHi.Text)) = 0 Then
                    
                    txtStZi.Text = "00"
                Else
                    iRet = MsgBox("�\���͈͂̊J�n�ɖ����͂̍��ڂ�����܂��B", vbExclamation, "���ُ͈�")
                    Exit Function
                End If
'EG20 V2.0.1.1 ADD END
            Else
                iRet = TextTime_Check("Hour", txtStZi.Text)
                If iRet = False Then
                    Exit Function
                End If
            End If
            
            '�J�n���`�F�b�N
            If Len(Trim(txtStFun.Text)) = 0 Then
'EG20 V2.0.1.1 DEL START
'                txtStFun.Text = "00"
'EG20 V2.0.1.1 DEL START
'EG20 V2.0.1.1 ADD START
                If Len(Trim(txtStNen.Text)) = 0 And _
                   Len(Trim(txtStTuki.Text)) = 0 And _
                   Len(Trim(txtStHi.Text)) = 0 And _
                   Len(Trim(txtStZi.Text)) = 0 Then
            
                    txtStFun.Text = "00"
            
                Else
                    iRet = MsgBox("�\���͈͂̊J�n�ɖ����͂̍��ڂ�����܂��B", vbExclamation, "���ُ͈�")
                    Exit Function
                End If
'EG20 V2.0.1.1 ADD END
            Else
                iRet = TextTime_Check("Minutes", txtStFun.Text)
                If iRet = False Then
                    Exit Function
                End If
            End If
            
            '���t�������`�F�b�N
            sChkDate = Format(txtStNen.Text, "0000") & "/" & _
                     Format(txtStTuki.Text, "00") & "/" & _
                     Format(txtStHi.Text, "00") & " " & _
                     Format(txtStZi.Text, "00") & ":" & _
                     Format(txtStFun.Text, "00")
            If IsDate(sChkDate) = False Then
                '���t�w�肪����������܂���
                MsgBox "���t�̎w�肪�ُ�ł��B", vbExclamation, "���ُ͈�"
                Exit Function
            End If
            
            
            '�I���N���̃Z�b�g
            sStAll = Format(txtStNen.Text, "0000") & _
                     Format(txtStTuki.Text, "00") & _
                     Format(txtStHi.Text, "00") & _
                     Format(txtStZi.Text, "00") & _
                     Format(txtStFun.Text, "00")
        End If
         
        '�I���`�F�b�N
        '���O�f�[�^�Ώێ����̐������`�F�b�N
        If Len(Trim(txtEdNen.Text)) = 0 And _
           Len(Trim(txtEdTuki.Text)) = 0 And _
           Len(Trim(txtEdHi.Text)) = 0 And _
           Len(Trim(txtEdZi.Text)) = 0 And _
           Len(Trim(txtEdFun.Text)) = 0 Then
           
           '�S�Ė����͂Ȃ�Max���Z�b�g
            sEdAll = "999999999999"
    
        ElseIf Len(Trim(txtEdNen.Text)) = 0 Or _
           Len(Trim(txtEdTuki.Text)) = 0 Or _
           Len(Trim(txtEdHi.Text)) = 0 Then
           
           '�I�������ɖ����͂̍��ڂ�����Ȃ�
            MsgBox "�\���͈͂̏I���ɖ����͂̍��ڂ�����܂��B", _
                   vbExclamation, _
                   "���ُ͈�"               ' EG20 V3.6.0.1 ADD
'                   "�����w��ُ�"          ' EG20 V3.6.0.1 DEL
            Exit Function
        
        Else
            '�I���N�`�F�b�N
            iRet = TextTime_Check("Year", txtEdNen.Text)
            If iRet = False Then
                Exit Function
            End If
            
            '�I�����`�F�b�N
            iRet = TextTime_Check("Month", txtEdTuki.Text)
            If iRet = False Then
                Exit Function
            End If
            
            '�I�����`�F�b�N
            iRet = TextTime_Check("Day", txtEdHi.Text)
            If iRet = False Then
                Exit Function
            End If
            
            '�I�����`�F�b�N
            If Len(Trim(txtEdZi.Text)) = 0 Then
'EG20 V2.0.1.1 DEL START
'                txtEdZi.Text = "00"
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
                If Len(Trim(txtEdNen.Text)) = 0 And _
                   Len(Trim(txtEdTuki.Text)) = 0 And _
                   Len(Trim(txtEdHi.Text)) = 0 Then

                    txtEdZi.Text = "00"
                Else
                    iRet = MsgBox("�\���͈͂̏I���ɖ����͂̍��ڂ�����܂��B", vbExclamation, "���ُ͈�")
                    Exit Function
                End If
'EG20 V2.0.1.1 ADD END
            Else
                iRet = TextTime_Check("Hour", txtEdZi.Text)
                If iRet = False Then
                    Exit Function
                End If
            End If
            
            '�I�����`�F�b�N
            If Len(Trim(txtEdFun.Text)) = 0 Then
'EG20 V2.0.1.1 DEL START
'                txtEdFun.Text = "00"
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
                If Len(Trim(txtEdNen.Text)) = 0 And _
                   Len(Trim(txtEdTuki.Text)) = 0 And _
                   Len(Trim(txtEdHi.Text)) = 0 And _
                   Len(Trim(txtEdZi.Text)) = 0 Then
                   
                    txtEdFun.Text = "00"
                
                Else
                    iRet = MsgBox("�\���͈͂̏I���ɖ����͂̍��ڂ�����܂��B", vbExclamation, "���ُ͈�")
                    Exit Function
                End If
'EG20 V2.0.1.1 ADD END            Else
                iRet = TextTime_Check("Minutes", txtEdFun.Text)
                If iRet = False Then
                    Exit Function
                End If
            End If
            
            '���t�������`�F�b�N
            sChkDate = Format(txtEdNen.Text, "0000") & "/" & _
                       Format(txtEdTuki.Text, "00") & "/" & _
                       Format(txtEdHi.Text, "00") & " " & _
                       Format(txtEdZi.Text, "00") & ":" & _
                       Format(txtEdFun.Text, "00")
            If IsDate(sChkDate) = False Then
                '���t�w�肪����������܂���
                MsgBox "���t�̎w�肪�ُ�ł��B", vbExclamation, "���ُ͈�"
                Exit Function
            End If
            
            '�I���N���̃Z�b�g
            sEdAll = Format(txtEdNen.Text, "0000") & _
                     Format(txtEdTuki.Text, "00") & _
                     Format(txtEdHi.Text, "00") & _
                     Format(txtEdZi.Text, "00") & _
                     Format(txtEdFun.Text, "00")
        End If

        '�J�n�A�I���O��`�F�b�N
        dStAll = Val(sStAll)
        dEdAll = Val(sEdAll)
        If dStAll > dEdAll Then
            MsgBox "�͈͎w��̊J�n�������I����������ɐݒ肳��Ă��܂��B", vbExclamation, "���ُ͈�"
            Exit Function
        End If

    End If
    
    
    '�A�v���P�[�V�������O�̏ꍇ�̂݃`�F�b�N����
    Dim bFlg As Boolean
    bFlg = False
    If optApp.Value = True Then
        '���W���[������
        For i = 0 To iModCnt
            '�`�F�b�N���n�m�Ȃ珈������
            If chkMod(i).Value = 1 And uModFileData(i).sType <> "" Then
                '�t���O�𗧂Ă�
                bFlg = True
            End If
        Next

        '����I������Ă��Ȃ����A�G���[�Ƃ���
        If bFlg = False Then
            MsgBox "���ڎw��Ɉُ킪����܂��B�w�肵���\�����ڂ��m�F���Ă��������B", vbCritical, "���ڎw��ُ�"
            Exit Function
        End If
    
        
        '���@����
        bFlg = False
        'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
        'EGX�`�F�b�N
'        For i = 0 To iEgxCnt
'            '�`�F�b�N���n�m�Ȃ珈������
'            If chkEGXGoki(uEgxFileData(i).iIndex).Value = 1 Then
'                '�t���O�𗧂Ă�
'                bFlg = True
'            End If
'        Next
        'EG20 V2.1.0.1 DEL END
        
        'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
        'EGX�`�F�b�N
        For i = 0 To chkEGXGoki.UBound
            '�`�F�b�N���n�m�Ȃ珈������
            If chkEGXGoki(i).Visible = True Then
                If chkEGXGoki(i).Value = 1 Then
                    '�t���O�𗧂Ă�
                    bFlg = True
                End If
            End If
        Next
        'EG20 V2.1.0.1 ADD END
        
        '����I������Ă��Ȃ����A�G���[�Ƃ���
        If bFlg = False Then
            MsgBox "���ڎw��Ɉُ킪����܂��B�w�肵���\�����ڂ��m�F���Ă��������B", vbCritical, "���ڎw��ُ�"
            Exit Function
        End If
    End If
    
    fLogSearchCheck = True              '�߂�l�ɐ�����Z�b�g
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : fWriteLogtxt
'//  �@�\����  : ���O�e�L�X�g�t�@�C�������ݏ���
'//  �@�\�T�v  : ���O�e�L�X�g�t�@�C���������݂��s���B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F�u���O�\��(�e�L�X�g�\��)�v�t����������
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
Private Function fWriteLogtxt() As Boolean

    Dim uLogConv As VB_LOG_DISP_SETTING '���O�����f�[�^
    Dim bRet As Boolean                 '�߂�l
    Dim sFileName As String
    Dim lId As Long                     '���[���h�c
    Dim bFlag As Boolean                '�t���O
    Dim iResponse As Integer            'MsgBox�{�^���R�[�h
    Dim iStatus As Long
    Dim MyPath As String
    Dim MyName As String
    Dim lErr As Long

    On Error Resume Next
    
    fWriteLogtxt = False

    '���O�ϊ������쐬����
    If sGetSearchData(uLogConv) = False Then
        Exit Function
    End If

    '���O�e�L�X�g�̍쐬
    If iNowChk1 = 1 Then
        MyPath = PATH_LDU_LOG & PATH_LDU_LOG_APP
    Else
        MyPath = PATH_LDU_LOG & PATH_LDU_LOG_HOSHU
    End If

    sObjectTopFile = Trim(Left(LstFile.List(LstFile.ListIndex), 12))
    sFileName = MyPath & sObjectTopFile
    sObjectTopFile = Left(sObjectTopFile, Len(sObjectTopFile) - 4) & ".txt"

    '�t�@�C���L���`�F�b�N
    MyName = Dir(sFileName, vbNormal)
    If MyName = "" Then ' ���[�v���J�n���܂��B
        MsgBox "�}�̏o�͂���f�[�^������܂���B", vbExclamation, "�f�[�^���x��"
        Exit Function
    End If
    
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[��\������
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_LOG_KANRI)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    '////////////////////////////////////////////////
    '�ێ��p�֐��F�\�����O�t�@�C���쐬����
    '////////////////////////////////////////////////
    iStatus = dllCreateDispLogFile(lErr, sFileName, uLogConv, sObjectTopFile, PATH_LDU_APP)
    If iStatus = 1 Then    '����̂Ƃ�
        fWriteLogtxt = True
    Else                    '�G���[�̂Ƃ�
        fWriteLogtxt = False
    End If

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
'//  �@�\�T�v  : ���O�g���[�X��ʂ��烍�O�ϊ������쐬����B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F�u���O�\��(�e�L�X�g�\��)�v�t����������
'//
'//              �^                  ����      �Ӗ�
'//  ����      : VB_LOG_DISP_SETTING uLogConv  [OUT]���O�ϊ����
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function sGetSearchData(uLogConv As VB_LOG_DISP_SETTING)
    
    Dim i As Integer
    Dim ii As Integer
    Dim iBitCnt As Double
    Dim bModFlg1 As Boolean
    Dim bModFlg2 As Boolean
    Dim bEGXGokiFlg As Boolean
    Dim bCPUGokiFlg As Boolean

    On Error Resume Next
    
    sGetSearchData = False

    '���O���
    If iNowChk1 = 1 Then
        '�A�v���I��
        uLogConv.LogType = 0
    Else
        '�ێ�I��
        uLogConv.LogType = 1
    End If


    '�͈͎w��
    If iNowChk2 = 1 Then
        '�͈͎w��A��
        uLogConv.TermType = 1
        uLogConv.StartTime = Format(txtStNen.Text, "0000") & _
                             Format(txtStTuki.Text, "00") & _
                             Format(txtStHi.Text, "00") & _
                             Format(txtStZi.Text, "00") & _
                             Format(txtStFun.Text, "00")
        uLogConv.EndTime = Format(txtEdNen.Text, "0000") & _
                           Format(txtEdTuki.Text, "00") & _
                           Format(txtEdHi.Text, "00") & _
                           Format(txtEdZi.Text, "00") & _
                           Format(txtEdFun.Text, "00")
        '�J�n�������͂̎��A�ŏ��l���Z�b�g����
        If Len(Trim(uLogConv.StartTime)) = 0 Then
            uLogConv.StartTime = "198001010000"
        End If
        '�I���������͂̎��A�ő�l���Z�b�g����
        If Len(Trim(uLogConv.EndTime)) = 0 Then
            uLogConv.EndTime = "207912312359"
        End If
    Else
        '�͈͎w�薳��
        uLogConv.TermType = 0
        uLogConv.StartTime = ""
        uLogConv.EndTime = ""
    End If


    '�\���I�v�V����
    If optSam.Value = True Then
        '�T�}���[�\��
        uLogConv.DispType = 0
    Else
        '�ڍו\��
        uLogConv.DispType = 1
    End If


    '�������ʎw��
    uLogConv.ResultType = 0
    '����
    If chkSeijou.Value = 1 Then
        uLogConv.ResultType = uLogConv.ResultType + 1
    End If
    '�ُ�
    If chkIjou.Value = 1 Then
        uLogConv.ResultType = uLogConv.ResultType + 2
    End If
    '��O
    If chkReigai.Value = 1 Then
        uLogConv.ResultType = uLogConv.ResultType + 8
    End If
    '�x��
    If chkKeikoku.Value = 1 Then
        uLogConv.ResultType = uLogConv.ResultType + 4
    End If


    '���ڎ�ʎw��
    uLogConv.ItemType = 0
    '�L�[����
    If chkKey.Value = 1 Then
        uLogConv.ItemType = uLogConv.ItemType + 1
    End If
    '�f�o�b�O����
    If chkDeb.Value = 1 Then
        uLogConv.ItemType = uLogConv.ItemType + 2
    End If


    '���W���[���w��
    uLogConv.ModuleType1 = 0
    uLogConv.ModuleType2 = 0
    uLogConv.ModuleType3 = 0
    '�\�����@�w��
    uLogConv.Goki = 0

    '�A�v���P�[�V�������O�̏ꍇ�̂݃`�F�b�N����
    If optApp.Value = True Then

        bModFlg1 = False
        bModFlg2 = False

        '�S���`�F�b�N
        For i = 0 To iModCnt
            '�`�F�b�N���n�m�Ȃ珈������
            If chkMod(i).Value = 1 And uModFileData(i).sType <> "" Then

                If uModFileData(i).iBit = 31 Then
                    '�Ή����郂�W���[���^�C�v�̃t���O�𗧂Ă�
                    If uModFileData(i).sType = 1 Then
                        bModFlg1 = True
                    Else
                        bModFlg2 = True
                    End If
                Else

                    '�r�b�g�J�E���g�v�Z
                    iBitCnt = 1
                    If uModFileData(i).iBit <> 0 Then
                        For ii = 1 To uModFileData(i).iBit
                            iBitCnt = iBitCnt * 2
                        Next
                    End If

                    '�Ή����郂�W���[���^�C�v�ɒǉ�����
                    If uModFileData(i).sType = 1 Then
                        uLogConv.ModuleType1 = uLogConv.ModuleType1 + iBitCnt
                    ElseIf uModFileData(i).sType = 2 Then
                        uLogConv.ModuleType2 = uLogConv.ModuleType2 + iBitCnt
                    ElseIf uModFileData(i).sType = 3 Then
                        uLogConv.ModuleType3 = uLogConv.ModuleType3 + iBitCnt
                    End If
                End If
            End If
        Next

        If bModFlg1 = True Then
            uLogConv.ModuleType1 = -2147483648# + uLogConv.ModuleType1
        End If

        If bModFlg2 = True Then
            uLogConv.ModuleType2 = -2147483648# + uLogConv.ModuleType2
        End If

        '�S���`�F�b�N(EGX)
        bEGXGokiFlg = False
        'EG20 V2.1.0.1 DEL START �y�t�F�[�Y�Q�Ή��z
'        For i = 0 To iEgxCnt
'            '�`�F�b�N���n�m�Ȃ珈������
'            If chkEGXGoki(uEgxFileData(i).iIndex).Value = 1 Then
'                If uEgxFileData(i).iRonri = 16 Then
'                    '�t���O�𗧂Ă�
'                    bEGXGokiFlg = True
'                Else
'
'                    '�r�b�g�J�E���g�v�Z
'                    iBitCnt = 1
'                    If uEgxFileData(i).iRonri <> 1 Then
'                        For ii = 1 To uEgxFileData(i).iRonri - 1
'                            iBitCnt = iBitCnt * 2
'                        Next
'                    End If
'
'                    '�ϐ��ɒǉ�����
'                    uLogConv.Goki = uLogConv.Goki + iBitCnt
'                End If
'            End If
'        Next
        'EG20 V2.1.0.1 DEL END
        
        'EG20 V2.1.0.1 ADD START �y�t�F�[�Y�Q�Ή��z
        For i = 0 To UBound(mintStatus)
            '�`�F�b�N���n�m�Ȃ珈������
            If mintStatus(i) = 1 Then
                If i = MAX_GATE_NO - 1 Then
                    '�t���O�𗧂Ă�
                    bEGXGokiFlg = True
                Else

                    '�r�b�g�J�E���g�v�Z
                    iBitCnt = 1
                    If i <> 0 Then
                        For ii = 1 To i
                            iBitCnt = iBitCnt * 2
                        Next
                    End If

                    '�ϐ��ɒǉ�����
                    uLogConv.Goki = uLogConv.Goki + iBitCnt
                End If
            End If
        Next
        'EG20 V2.1.0.1 DEL END

        If bEGXGokiFlg = True Then
'            uLogConv.Goki = -32768# + uLogConv.Goki            'EG20 V2.1.0.1 DEL �y�t�F�[�Y�Q�Ή��z
            uLogConv.Goki = -2147483648# + uLogConv.Goki        'EG20 V2.1.0.1 ADD �y�t�F�[�Y�Q�Ή��z
        End If
    End If

    sGetSearchData = True
End Function

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
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  CODED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
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

        Case ML_ID_HOSHU_ACTIVE_REQ
             '�ێ��ʃA�N�e�B�u�\���̏ꍇ
             '�u�ێ��ʃA�N�e�B�u�\���v����M����v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
'             AppActivate frmKansiLogKanri.Caption, False   ' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL
'             pfFormActive (frmIDULogkanri.hwnd)            ' EG20 V8.1.0.1�yEG20_KANSI05_01�zDEL
             AppActivate frmLDULogkanri.Caption, False      ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
             pfFormActive (frmLDULogkanri.hwnd)             ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
             fMailRecieve = ML_ID_HOSHU_ACTIVE_REQ

        Case ML_ID_IDU_LDU_LGCHGREQ_RES
             '���O�ؑ֗v��RES�̏ꍇ
             '�u���O�ؑ֗v��RES��M����v���O�o��
             Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, LOG_CHANGE_RES_RECV, 0)
             fMailRecieve = ML_ID_IDU_LDU_LGCHGREQ_RES

        Case Else
        '���[���h�c�s��
          '�u���[��ID�s���v���O�o��
          Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
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
        AppActivate frmLDULogkanri.Caption, False
        pfFormActive (frmLDULogkanri.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-07   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdZHisentaku_Click()

    Dim intLoop As Integer
    
    On Error Resume Next
    
    For intLoop = 0 To chkEGXGoki.UBound
        chkEGXGoki(intLoop).Value = CHECKBOX_OFF
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-07   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdZSentaku_Click()

    Dim intLoop As Integer
    
    On Error Resume Next
    
    For intLoop = 0 To chkEGXGoki.UBound
        chkEGXGoki(intLoop).Value = CHECKBOX_ON
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-07   CODED   BY [TCC] M.Matsumoto
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
        chkEGXGoki(intLoop).Value = CHECKBOX_OFF
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
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-07   CODED   BY [TCC] M.Matsumoto
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
        chkEGXGoki(intLoop).Value = CHECKBOX_ON
    Next intLoop
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : chkEGXGoki_Click
'//  �@�\����  : �w�荆�@�I�v�V�����{�^���N���b�N������
'//  �@�\�T�v  : �����ϐ���ON/OFF��؂�ւ���
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@�@Index  �@[IN]�I�v�V�����{�^���C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-12-08   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Private Sub chkEGXGoki_Click(Index As Integer)

    Dim intGoki As Integer
    
    On Error Resume Next
    
    intGoki = CInt(chkEGXGoki(Index).Tag) - 1
    
    mintStatus(intGoki) = chkEGXGoki(Index).Value
    
End Sub

'EG20 V2.1.0.1 ADD END

