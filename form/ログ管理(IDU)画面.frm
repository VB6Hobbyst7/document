VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIDULogkanri 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "                                                                  �h�c���p���j�b�g���O�Ǘ�"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   9360
      Top             =   7440
   End
   Begin VB.CommandButton cmdInstall 
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
      TabIndex        =   238
      Top             =   6480
      Width           =   2600
   End
   Begin VB.TextBox txtDummy 
      Height          =   375
      Left            =   11400
      TabIndex        =   0
      Text            =   "Text11"
      Top             =   15000
      Width           =   2175
   End
   Begin VB.CommandButton cmdLogHyouzi 
      Caption         =   "   ���O�\��    (�e�L�X�g�\��)"
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
      TabIndex        =   176
      Top             =   480
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
      TabIndex        =   177
      Top             =   1680
      Width           =   2600
   End
   Begin VB.CommandButton cmdSyslog 
      Caption         =   " �V�X�e�����O   �}�̏o��"
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
      TabIndex        =   178
      Top             =   2880
      Visible         =   0   'False
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
      TabIndex        =   179
      Top             =   4080
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.CommandButton cmdMemoridump 
      Caption         =   " �������_���v   �}�̏o��"
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
      TabIndex        =   180
      Top             =   5280
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
      TabIndex        =   181
      Top             =   7920
      Width           =   2600
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   8535
      Left            =   120
      TabIndex        =   182
      Top             =   375
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
      TabPicture(0)   =   "���O�Ǘ�(IDU)���.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblSize"
      Tab(0).Control(1)=   "lblEnd"
      Tab(0).Control(2)=   "lblStart"
      Tab(0).Control(3)=   "lblFile"
      Tab(0).Control(4)=   "cmdRefresh"
      Tab(0).Control(5)=   "optHoshu"
      Tab(0).Control(6)=   "optApp"
      Tab(0).Control(7)=   "LstFile"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "�\�����ڎw��"
      TabPicture(1)   =   "���O�Ǘ�(IDU)���.frx":001C
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
      TabPicture(2)   =   "���O�Ǘ�(IDU)���.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabCorner"
      Tab(2).Control(1)=   "cmdHHisentaku"
      Tab(2).Control(2)=   "cmdHSentaku"
      Tab(2).Control(3)=   "cmdZHisentaku"
      Tab(2).Control(4)=   "cmdZSentaku"
      Tab(2).ControlCount=   5
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
         TabIndex        =   188
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
            Top             =   630
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
            Top             =   330
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
            TabIndex        =   200
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
            TabIndex        =   199
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
            TabIndex        =   198
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
            TabIndex        =   197
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
            TabIndex        =   196
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
            TabIndex        =   195
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
            TabIndex        =   194
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
            TabIndex        =   193
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
            TabIndex        =   192
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
            TabIndex        =   191
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
            TabIndex        =   190
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
            TabIndex        =   189
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
         TabIndex        =   187
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
            Top             =   540
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
         TabIndex        =   186
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
         TabIndex        =   185
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
         TabIndex        =   183
         Top             =   2520
         Width           =   8415
         Begin VB.CommandButton cmdModSen 
            Caption         =   "�S�đI��"
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
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdModHi 
            Caption         =   "�S�Ĕ�I��"
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
            Left            =   1680
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
         Begin VB.Frame frmModMeisai 
            Height          =   5055
            Left            =   120
            TabIndex        =   184
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
               TabIndex        =   236
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
               TabIndex        =   235
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
               TabIndex        =   234
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
               TabIndex        =   233
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
               TabIndex        =   232
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
               TabIndex        =   231
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
               TabIndex        =   230
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
               TabIndex        =   229
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
               TabIndex        =   228
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
               TabIndex        =   227
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
               TabIndex        =   226
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
               TabIndex        =   225
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
               TabIndex        =   224
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
               TabIndex        =   223
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
               TabIndex        =   222
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
               TabIndex        =   221
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
               TabIndex        =   220
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
               TabIndex        =   219
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
               TabIndex        =   218
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
               TabIndex        =   217
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
               TabIndex        =   216
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
               TabIndex        =   215
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
               TabIndex        =   214
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
               TabIndex        =   213
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
               TabIndex        =   212
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
               TabIndex        =   211
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
               TabIndex        =   210
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
               TabIndex        =   209
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
               TabIndex        =   208
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
               TabIndex        =   207
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
               TabIndex        =   206
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
               TabIndex        =   205
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
               Width           =   1890
            End
         End
      End
      Begin VB.CommandButton cmdZSentaku 
         Caption         =   "  �S�R�[�i    �S���@   �I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   -74760
         TabIndex        =   75
         Top             =   840
         Width           =   2000
      End
      Begin VB.CommandButton cmdZHisentaku 
         Caption         =   "   �S�R�[�i     �S���@ ��I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   -72600
         TabIndex        =   76
         Top             =   840
         Width           =   2000
      End
      Begin VB.CommandButton cmdHSentaku 
         Caption         =   "  �\���R�[�i   �S���@ �I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   -70440
         TabIndex        =   77
         Top             =   840
         Width           =   2000
      End
      Begin VB.CommandButton cmdHHisentaku 
         Caption         =   "  �\���R�[�i    �S���@ ��I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   -68280
         TabIndex        =   78
         Top             =   840
         Width           =   2000
      End
      Begin TabDlg.SSTab tabCorner 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   79
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
         TabCaption(0)   =   " "
         TabPicture(0)   =   "���O�Ǘ�(IDU)���.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "chkCorner(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "chkCorner(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkCorner(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkCorner(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "chkCorner(4)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkCorner(5)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "chkCorner(6)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "chkCorner(7)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "chkCorner(8)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "chkCorner(9)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "chkCorner(10)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "chkCorner(11)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "chkCorner(12)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "chkCorner(13)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "chkCorner(14)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "chkCorner(15)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "  "
         TabPicture(1)   =   "���O�Ǘ�(IDU)���.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkCorner(16)"
         Tab(1).Control(1)=   "chkCorner(17)"
         Tab(1).Control(2)=   "chkCorner(18)"
         Tab(1).Control(3)=   "chkCorner(19)"
         Tab(1).Control(4)=   "chkCorner(20)"
         Tab(1).Control(5)=   "chkCorner(21)"
         Tab(1).Control(6)=   "chkCorner(22)"
         Tab(1).Control(7)=   "chkCorner(23)"
         Tab(1).Control(8)=   "chkCorner(24)"
         Tab(1).Control(9)=   "chkCorner(25)"
         Tab(1).Control(10)=   "chkCorner(26)"
         Tab(1).Control(11)=   "chkCorner(27)"
         Tab(1).Control(12)=   "chkCorner(28)"
         Tab(1).Control(13)=   "chkCorner(29)"
         Tab(1).Control(14)=   "chkCorner(30)"
         Tab(1).Control(15)=   "chkCorner(31)"
         Tab(1).ControlCount=   16
         TabCaption(2)   =   "  "
         TabPicture(2)   =   "���O�Ǘ�(IDU)���.frx":008C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkCorner(32)"
         Tab(2).Control(1)=   "chkCorner(33)"
         Tab(2).Control(2)=   "chkCorner(34)"
         Tab(2).Control(3)=   "chkCorner(35)"
         Tab(2).Control(4)=   "chkCorner(36)"
         Tab(2).Control(5)=   "chkCorner(37)"
         Tab(2).Control(6)=   "chkCorner(38)"
         Tab(2).Control(7)=   "chkCorner(39)"
         Tab(2).Control(8)=   "chkCorner(40)"
         Tab(2).Control(9)=   "chkCorner(41)"
         Tab(2).Control(10)=   "chkCorner(42)"
         Tab(2).Control(11)=   "chkCorner(43)"
         Tab(2).Control(12)=   "chkCorner(44)"
         Tab(2).Control(13)=   "chkCorner(45)"
         Tab(2).Control(14)=   "chkCorner(46)"
         Tab(2).Control(15)=   "chkCorner(47)"
         Tab(2).ControlCount=   16
         TabCaption(3)   =   "  "
         TabPicture(3)   =   "���O�Ǘ�(IDU)���.frx":00A8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "chkCorner(48)"
         Tab(3).Control(1)=   "chkCorner(49)"
         Tab(3).Control(2)=   "chkCorner(50)"
         Tab(3).Control(3)=   "chkCorner(51)"
         Tab(3).Control(4)=   "chkCorner(52)"
         Tab(3).Control(5)=   "chkCorner(53)"
         Tab(3).Control(6)=   "chkCorner(54)"
         Tab(3).Control(7)=   "chkCorner(55)"
         Tab(3).Control(8)=   "chkCorner(56)"
         Tab(3).Control(9)=   "chkCorner(57)"
         Tab(3).Control(10)=   "chkCorner(58)"
         Tab(3).Control(11)=   "chkCorner(59)"
         Tab(3).Control(12)=   "chkCorner(60)"
         Tab(3).Control(13)=   "chkCorner(61)"
         Tab(3).Control(14)=   "chkCorner(62)"
         Tab(3).Control(15)=   "chkCorner(63)"
         Tab(3).ControlCount=   16
         TabCaption(4)   =   "  "
         TabPicture(4)   =   "���O�Ǘ�(IDU)���.frx":00C4
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "chkCorner(64)"
         Tab(4).Control(1)=   "chkCorner(65)"
         Tab(4).Control(2)=   "chkCorner(66)"
         Tab(4).Control(3)=   "chkCorner(67)"
         Tab(4).Control(4)=   "chkCorner(68)"
         Tab(4).Control(5)=   "chkCorner(69)"
         Tab(4).Control(6)=   "chkCorner(70)"
         Tab(4).Control(7)=   "chkCorner(71)"
         Tab(4).Control(8)=   "chkCorner(72)"
         Tab(4).Control(9)=   "chkCorner(73)"
         Tab(4).Control(10)=   "chkCorner(74)"
         Tab(4).Control(11)=   "chkCorner(75)"
         Tab(4).Control(12)=   "chkCorner(76)"
         Tab(4).Control(13)=   "chkCorner(77)"
         Tab(4).Control(14)=   "chkCorner(78)"
         Tab(4).Control(15)=   "chkCorner(79)"
         Tab(4).ControlCount=   16
         TabCaption(5)   =   "  "
         TabPicture(5)   =   "���O�Ǘ�(IDU)���.frx":00E0
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "chkCorner(80)"
         Tab(5).Control(1)=   "chkCorner(81)"
         Tab(5).Control(2)=   "chkCorner(82)"
         Tab(5).Control(3)=   "chkCorner(83)"
         Tab(5).Control(4)=   "chkCorner(84)"
         Tab(5).Control(5)=   "chkCorner(85)"
         Tab(5).Control(6)=   "chkCorner(86)"
         Tab(5).Control(7)=   "chkCorner(87)"
         Tab(5).Control(8)=   "chkCorner(88)"
         Tab(5).Control(9)=   "chkCorner(89)"
         Tab(5).Control(10)=   "chkCorner(90)"
         Tab(5).Control(11)=   "chkCorner(91)"
         Tab(5).Control(12)=   "chkCorner(92)"
         Tab(5).Control(13)=   "chkCorner(93)"
         Tab(5).Control(14)=   "chkCorner(94)"
         Tab(5).Control(15)=   "chkCorner(95)"
         Tab(5).ControlCount=   16
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   175
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   174
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   173
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   172
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   171
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   170
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   169
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   168
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   167
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   166
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   165
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   164
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   163
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   162
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   161
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   160
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   159
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   158
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   157
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   155
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   154
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   153
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   152
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   151
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   150
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   149
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   148
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   147
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   146
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   145
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   144
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   143
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   142
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   141
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   140
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   139
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   138
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   137
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   136
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   135
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   134
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   133
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   132
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   131
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   130
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   129
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   128
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   127
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   126
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   125
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   124
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   123
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   122
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   121
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   120
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   119
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   118
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   117
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   116
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   115
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   114
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   113
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   112
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   111
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   110
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   109
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   108
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   107
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   106
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   105
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   104
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   103
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   102
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   101
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   100
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   99
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   98
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   97
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            TabIndex        =   96
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   15
            Left            =   6360
            TabIndex        =   95
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   14
            Left            =   4200
            TabIndex        =   94
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   13
            Left            =   2160
            TabIndex        =   93
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   12
            Left            =   120
            TabIndex        =   92
            Top             =   2040
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   11
            Left            =   6360
            TabIndex        =   91
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   10
            Left            =   4200
            TabIndex        =   90
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   9
            Left            =   2160
            TabIndex        =   89
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   8
            Left            =   120
            TabIndex        =   88
            Top             =   1560
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   7
            Left            =   6360
            TabIndex        =   87
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   6
            Left            =   4200
            TabIndex        =   86
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   5
            Left            =   2160
            TabIndex        =   85
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   4
            Left            =   120
            TabIndex        =   84
            Top             =   1080
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   3
            Left            =   6360
            TabIndex        =   83
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   2
            Left            =   4200
            TabIndex        =   82
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   1
            Left            =   2160
            TabIndex        =   81
            Top             =   600
            Value           =   1  '����
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkCorner 
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
            Index           =   0
            Left            =   120
            TabIndex        =   80
            Top             =   600
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
         TabIndex        =   204
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
         TabIndex        =   203
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
         TabIndex        =   202
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
         TabIndex        =   201
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackColor       =   &H00C0C000&
      Caption         =   "IDU�A�v���P�[�V�������O�Ǘ�"
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
      TabIndex        =   237
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmIDULogkanri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmIDULogKanri.frm
'//  �p�b�P�[�W���FID���p���j�b�g���O�Ǘ����
'//
'//  �T�v�FID���p���j�b�g���O�Ǘ����
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 �EID���p���j�b�g�A���O�Ǘ����(frmLogKanri.frm)�𗬗p
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-06-12   REVISED BY [TCC] S.Terao
'//                 �t�F�[�Y�R�Ή�
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�H���_�I���|�b�v�A�b�v��ʂ̏����t�H���_�ύX
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02   REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή��y�Ď�D-115�z
'//                 �@�E�������ʃ��b�Z�[�W�{�b�N�X�̕����ύX
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

Public sYobidasi As String
Public iNowChk1 As Integer
Public iNowChk2 As Integer

'DB�ڑ��p
'�A�v�����O�p
Private cnConn              As New ADODB.Connection     'Connection �I�u�W�F�N�g�̒�`
Private rsRecordSet         As New ADODB.Recordset      'RecordSet �I�u�W�F�N�g�̒�`
'�A�v�����O�e�[�u���擾�p
Private gLogData() As typLogDataTable
'�A�v�����O�ۑ��Ǘ�DB
Private Type typLogDataTable
    sName As String
    sStTime As String
    sEdTime As String
    iSize As Long
End Type

'�ێ烍�O�p
Private cnConn2              As New ADODB.Connection     'Connection �I�u�W�F�N�g�̒�`

'///////////////////////////////////////////////////////////////////
'�Ώۃt�@�C���t���p�X�i����̧�ق̎��A��߰�1�����ŋ�؂�B�j
'///////////////////////////////////////////////////////////////////
Private sObjectFiles As String      '۸�̧��ؽ��ޯ���őI�𒆂�̧�ق����߽������
Private sObjectTopFile As String    '����A�I�𒆂̐擪�i�ŋ��j̧�ٖ��B(12����)�B

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
    sName As String                '���W���[����
    sDai  As String                '�區��
    sShou As String                '������
    sType As String                '���W���[���^�C�v
    iBit  As Integer               '�r�b�g�ԍ�
End Type

Private uModFileData(79) As ModFileData
Private iModCnt As Integer

'///////////////////////////////////////////////////////////////////
'ICM���i�[�G���A
'///////////////////////////////////////////////////////////////////
Private Type IcmFileData
    iRonri As Integer               '�_�����@
    iHyozi As Integer               '�\�����@
    iConer As Integer               '�R�[�i�[�ԍ�
    iIndex As Integer               'chkCorner��INDEX
End Type

Private uIcmFileData(31) As IcmFileData
Private iIcmCnt As Integer

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

'V1.3.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : ID���p���j�b�g���O�Ǘ����(�A�N�e�B�u��)
'//  �@�\�T�v  : �őO�ʕ\�����s���B
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
    pfFormActive (hwnd)
    '���[����M�^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : ID���p���j�b�g���O�Ǘ����(�f�B�A�N�e�B�u��)
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
'//  �@�\����  : ID���p���j�b�g���O�Ǘ����(���[�h��)
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
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-25  CODED BY  [TCC] T.Koyama
'//                 EG20�t�F�[�Y�Q�Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim intModulesFileNo As Integer
    Dim sModules As String * IDU_LOG_SIZE    '�P�s���t�@�C�����e�擾�p
    Dim Cnt As Integer
    Dim iMozi As Integer
    Dim iKbn As Integer
    Dim iRet As Integer
'    Dim sConer As String * IDU_LOG_CONER_SIZE '�t�@�C���`�F���W�c�[���̎��s�t�@�C�����i�t���p�X)   ' EG20 V3.6.0.1 DEL
    Dim sConer As String * 30                  '�t�@�C���`�F���W�c�[���̎��s�t�@�C�����i�t���p�X)   ' EG20 V3.6.0.1 ADD
    Dim sType As String * IDU_LOG_TYPE        '�ݒu�^�C�v
    Dim sIcmData As String * IDU_LOG_SIZE     '�P�s���t�@�C�����e�擾�p
    Dim i As Integer                          '���[�v�p
    Dim sKeyName As String
    Dim str As String
    Dim iLoop As Integer
    Dim MyName As String
    Dim iErr As Integer
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
' EG20 V3.6.0.1 ADD START
    Dim myLen As Long
    Dim strCodeTxt As String
    Dim strCorner As String
' EG20 V3.6.0.1 ADD END
    
    '�p�X�w��
    IDU_PROFILE_NAME = PATH_IDU_APP & IDU_STATION_FILE
    IDU_PROFILE_NAME_ICM = PATH_IDU_APP & IDU_ICM_FILE
    
    gStrCurrentForm = sFormName_IDULog
     
    cmdCancel.Caption = "���O�Ǘ�" & Chr(13) & "��ʂ֖߂�"
    cmdLogHyouzi.Caption = "���O�\��" & Chr(13) & "(�e�L�X�g�\���j"
    cmdZSentaku.Caption = "�S�R�[�i" & Chr(13) & "�S���@�@�I��"
    cmdZHisentaku.Caption = "�S�R�[�i" & Chr(13) & "�S���@�@��I��"
    cmdHSentaku.Caption = "�\���R�[�i" & Chr(13) & "�S���@�@�I��"
    cmdHHisentaku.Caption = "�\���R�[�i" & Chr(13) & "�S���@�@��I��"
     
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
    
    For i = 0 To 5
        tabCorner.Tab = 5 - i
        tabCorner.Caption = ""
    Next

    'INI�t�@�C�����̐ݒ�擾
    '���W���[���w��擾
    On Error GoTo FileError
    iErr = 1
    
    '�t�@�C���L���`�F�b�N
    MyName = Dir(PATH_IDU_APP & IDU_MODULES_FILE_FULLPASS, vbNormal)
    If MyName = "" Then
        GoTo FileError
    End If
    
    Cnt = 0
    
    For Cnt = 0 To 79
        sKeyName = "ID" & Format(Cnt, "000")
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_ID, _
                                       sKeyName, _
                                       DEFAILT, sModules, Len(sModules), _
                                       PATH_IDU_APP & IDU_MODULES_FILE_FULLPASS)
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
    
    '�t�@�C���L���`�F�b�N
    MyName = Dir(IDU_PROFILE_NAME, vbNormal)
    If MyName = "" Then
        GoTo FileError
    End If
    
    MyName = Dir(IDU_PROFILE_NAME_ICM, vbNormal)
    If MyName = "" Then
        iErr = 3
        GoTo FileError
    End If
    
    
    '�R�[�i�[���̎擾
    '�U�R�[�i�[�ڂ̖��̎擾
    iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER6, _
                                   IDU_PROFILE_KEY_NAME_TYPE, _
                                   DEFAILT, sType, Len(sType), _
                                   IDU_PROFILE_NAME)
    If Int(sType) <> 0 Then
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER6, _
                                       IDU_PROFILE_KEY_NAME_CONER, _
                                       DEFAILT, sConer, Len(sConer), _
                                       IDU_PROFILE_NAME)
        If iRet = 0 Then
            GoTo FileError
        End If
' EG20 V3.6.0.1 DEL START
'        tabCorner.Tab = 5
'        tabCorner.Caption = sConer
' EG20 V3.6.0.1 DEL END
' EG20 V3.6.0.1 ADD START
        strCodeTxt = StrConv(sConer, vbFromUnicode)     '�������ϊ�
        myLen = LenB(strCodeTxt)                        '���p���Z�̃o�C�g�����擾
    
        If myLen <= 24 Then                             '�w��̒������Z���ꍇ
            strCorner = strCodeTxt

        Else
            '�Y���̕�����̕��������ꍇ�A�w��̃o�C�g�ŃJ�b�g����
            strCorner = StrConv(LeftB$(strCodeTxt, 24), vbUnicode)

            If InStr(strCorner, vbNullChar) > 0 Then
                '�����P�o�C�g�ڂŕ��f���ꂽ�ꍇ�̏���
                strCorner = Left$(strCorner, InStr(strCorner, vbNullChar) - 1) & " "
            End If
        End If
        
        tabCorner.Tab = 5
        tabCorner.Caption = strCorner
        tabCorner.Font.Size = 10
' EG20 V3.6.0.1 ADD END
    Else
        tabCorner.Tab = 5
        tabCorner.Caption = ""
    End If
    
    '�T�R�[�i�[�ڂ̖��̎擾
    iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER5, _
                                   IDU_PROFILE_KEY_NAME_TYPE, _
                                   DEFAILT, sType, Len(sType), _
                                   IDU_PROFILE_NAME)
    If Int(sType) <> 0 Then
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER5, _
                                       IDU_PROFILE_KEY_NAME_CONER, _
                                       DEFAILT, sConer, Len(sConer), _
                                       IDU_PROFILE_NAME)
        If iRet = 0 Then
            GoTo FileError
        End If
' EG20 V3.6.0.1 DEL START
'        tabCorner.Tab = 4
'        tabCorner.Caption = sConer
' EG20 V3.6.0.1 DEL END
' EG20 V3.6.0.1 ADD START
        strCodeTxt = StrConv(sConer, vbFromUnicode)     '�������ϊ�
        myLen = LenB(strCodeTxt)                        '���p���Z�̃o�C�g�����擾
    
        If myLen <= 24 Then                             '�w��̒������Z���ꍇ
            strCorner = strCodeTxt

        Else
            '�Y���̕�����̕��������ꍇ�A�w��̃o�C�g�ŃJ�b�g����
            strCorner = StrConv(LeftB$(strCodeTxt, 24), vbUnicode)

            If InStr(strCorner, vbNullChar) > 0 Then
                '�����P�o�C�g�ڂŕ��f���ꂽ�ꍇ�̏���
                strCorner = Left$(strCorner, InStr(strCorner, vbNullChar) - 1) & " "
            End If
        End If
        
        tabCorner.Tab = 4
        tabCorner.Caption = strCorner
        tabCorner.Font.Size = 10
' EG20 V3.6.0.1 ADD END
    Else
        tabCorner.Tab = 4
        tabCorner.Caption = ""
    End If
    
    '�S�R�[�i�[�ڂ̖��̎擾
    iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER4, _
                                   IDU_PROFILE_KEY_NAME_TYPE, _
                                   DEFAILT, sType, Len(sType), _
                                   IDU_PROFILE_NAME)
    If Int(sType) <> 0 Then
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER4, _
                                       IDU_PROFILE_KEY_NAME_CONER, _
                                       DEFAILT, sConer, Len(sConer), _
                                       IDU_PROFILE_NAME)
        If iRet = 0 Then
            GoTo FileError
        End If
' EG20 V3.6.0.1 DEL START
'        tabCorner.Tab = 3
'        tabCorner.Caption = sConer
' EG20 V3.6.0.1 DEL END
' EG20 V3.6.0.1 ADD START
        strCodeTxt = StrConv(sConer, vbFromUnicode)     '�������ϊ�
        myLen = LenB(strCodeTxt)                        '���p���Z�̃o�C�g�����擾
    
        If myLen <= 24 Then                             '�w��̒������Z���ꍇ
            strCorner = strCodeTxt

        Else
            '�Y���̕�����̕��������ꍇ�A�w��̃o�C�g�ŃJ�b�g����
            strCorner = StrConv(LeftB$(strCodeTxt, 24), vbUnicode)

            If InStr(strCorner, vbNullChar) > 0 Then
                '�����P�o�C�g�ڂŕ��f���ꂽ�ꍇ�̏���
                strCorner = Left$(strCorner, InStr(strCorner, vbNullChar) - 1) & " "
            End If
        End If
        
        tabCorner.Tab = 3
        tabCorner.Caption = strCorner
        tabCorner.Font.Size = 10
' EG20 V3.6.0.1 ADD END
    Else
        tabCorner.Tab = 3
        tabCorner.Caption = ""
    End If
     
    '�R�R�[�i�[�ڂ̖��̎擾
    iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER3, _
                                   IDU_PROFILE_KEY_NAME_TYPE, _
                                   DEFAILT, sType, Len(sType), _
                                   IDU_PROFILE_NAME)
    If Int(sType) <> 0 Then
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER3, _
                                       IDU_PROFILE_KEY_NAME_CONER, _
                                       DEFAILT, sConer, Len(sConer), _
                                       IDU_PROFILE_NAME)
        If iRet = 0 Then
            GoTo FileError
        End If
' EG20 V3.6.0.1 DEL START
'        tabCorner.Tab = 2
'        tabCorner.Caption = sConer
' EG20 V3.6.0.1 DEL END
' EG20 V3.6.0.1 ADD START
        strCodeTxt = StrConv(sConer, vbFromUnicode)     '�������ϊ�
        myLen = LenB(strCodeTxt)                        '���p���Z�̃o�C�g�����擾
    
        If myLen <= 24 Then                             '�w��̒������Z���ꍇ
            strCorner = strCodeTxt

        Else
            '�Y���̕�����̕��������ꍇ�A�w��̃o�C�g�ŃJ�b�g����
            strCorner = StrConv(LeftB$(strCodeTxt, 24), vbUnicode)

            If InStr(strCorner, vbNullChar) > 0 Then
                '�����P�o�C�g�ڂŕ��f���ꂽ�ꍇ�̏���
                strCorner = Left$(strCorner, InStr(strCorner, vbNullChar) - 1) & " "
            End If
        End If
        
        tabCorner.Tab = 2
        tabCorner.Caption = strCorner
        tabCorner.Font.Size = 10
' EG20 V3.6.0.1 ADD END
    Else
        tabCorner.Tab = 2
        tabCorner.Caption = ""
    End If
    
    '�Q�R�[�i�[�ڂ̖��̎擾
    iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER2, _
                                   IDU_PROFILE_KEY_NAME_TYPE, _
                                   DEFAILT, sType, Len(sType), _
                                   IDU_PROFILE_NAME)
    If Int(sType) <> 0 Then
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER2, _
                                       IDU_PROFILE_KEY_NAME_CONER, _
                                       DEFAILT, sConer, Len(sConer), _
                                       IDU_PROFILE_NAME)
        If iRet = 0 Then
            GoTo FileError
        End If
' EG20 V3.6.0.1 DEL START
'        tabCorner.Tab = 1
'        tabCorner.Caption = sConer
' EG20 V3.6.0.1 DEL END
' EG20 V3.6.0.1 ADD START
        strCodeTxt = StrConv(sConer, vbFromUnicode)     '�������ϊ�
        myLen = LenB(strCodeTxt)                        '���p���Z�̃o�C�g�����擾
    
        If myLen <= 24 Then                             '�w��̒������Z���ꍇ
            strCorner = strCodeTxt

        Else
            '�Y���̕�����̕��������ꍇ�A�w��̃o�C�g�ŃJ�b�g����
            strCorner = StrConv(LeftB$(strCodeTxt, 24), vbUnicode)

            If InStr(strCorner, vbNullChar) > 0 Then
                '�����P�o�C�g�ڂŕ��f���ꂽ�ꍇ�̏���
                strCorner = Left$(strCorner, InStr(strCorner, vbNullChar) - 1) & " "
            End If
        End If
        
        tabCorner.Tab = 1
        tabCorner.Caption = strCorner
        tabCorner.Font.Size = 10
' EG20 V3.6.0.1 ADD END
    Else
        tabCorner.Tab = 1
        tabCorner.Caption = ""
    End If

    '�P�R�[�i�[�ڂ̖��̎擾
    iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER1, _
                                   IDU_PROFILE_KEY_NAME_TYPE, _
                                   DEFAILT, sType, Len(sType), _
                                   IDU_PROFILE_NAME)
    If Int(sType) <> 0 Then
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_CONER1, _
                                       IDU_PROFILE_KEY_NAME_CONER, _
                                       DEFAILT, sConer, Len(sConer), _
                                       IDU_PROFILE_NAME)
        If iRet = 0 Then
            GoTo FileError
        End If

' EG20 V3.6.0.1 DEL START
'        tabCorner.Tab = 0
'        tabCorner.Caption = sConer
' EG20 V3.6.0.1 DEL END
' EG20 V3.6.0.1 ADD START
        strCodeTxt = StrConv(sConer, vbFromUnicode)     '�������ϊ�
        myLen = LenB(strCodeTxt)                        '���p���Z�̃o�C�g�����擾
    
        If myLen <= 24 Then                             '�w��̒������Z���ꍇ
            strCorner = strCodeTxt

        Else
            '�Y���̕�����̕��������ꍇ�A�w��̃o�C�g�ŃJ�b�g����
            strCorner = StrConv(LeftB$(strCodeTxt, 24), vbUnicode)

            If InStr(strCorner, vbNullChar) > 0 Then
                '�����P�o�C�g�ڂŕ��f���ꂽ�ꍇ�̏���
                strCorner = Left$(strCorner, InStr(strCorner, vbNullChar) - 1) & " "
            End If
        End If
        
        tabCorner.Tab = 0
        tabCorner.Caption = strCorner
        tabCorner.Font.Size = 10
' EG20 V3.6.0.1 ADD END
    Else
        tabCorner.Tab = 0
        tabCorner.Caption = ""
    End If

    iErr = 3

    iIcmCnt = -1
    'ICM���擾
    For i = 1 To 32
        sKeyName = "icm" & Format(i, "00")
        iRet = GetPrivateProfileString(IDU_PROFILE_SECTION_NAME_ICM, _
                                       sKeyName, _
                                       DEFAILT, sIcmData, Len(sIcmData), _
                                       IDU_PROFILE_NAME_ICM)
        If iRet = 0 Then
            GoTo FileError
        End If
                        
        '�f�[�^�̎擾
        ReDim sFData(14)
        iFCnt = 1
        
        For iFLoop = 1 To Len(sIcmData)
            If Mid(sIcmData, iFLoop, 1) <> " " Or Mid(sIcmData, iFLoop, 1) <> "," Then
                iFLoop2 = iFLoop
                Do
                    iFLoop2 = iFLoop2 + 1
                    If iFLoop2 > Len(sIcmData) Then
                        sFData(iFCnt) = Mid(sIcmData, iFLoop, iFLoop2 - iFLoop)
                        iFCnt = iFCnt + 1
                        If iFCnt >= 15 Then
                            Exit For
                        End If
                        iFLoop = iFLoop2
                        Exit Do
                    End If
                    
                    If Mid(sIcmData, iFLoop2, 1) = " " Or Mid(sIcmData, iFLoop2, 1) = "," Then
                        sFData(iFCnt) = Mid(sIcmData, iFLoop, iFLoop2 - iFLoop)
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
        
        '�ʘH��ʂ����ݒu�̎��͏�������
        If Trim(sFData(5)) <> "��" Then
            iIcmCnt = iIcmCnt + 1
            uIcmFileData(iIcmCnt).iRonri = i
            uIcmFileData(iIcmCnt).iHyozi = Trim(sFData(1))
            uIcmFileData(iIcmCnt).iConer = Trim(sFData(3))
            uIcmFileData(iIcmCnt).iIndex = uIcmFileData(iIcmCnt).iConer * 16 - 16 + Int(Trim(sFData(4))) - 1
            chkCorner(uIcmFileData(iIcmCnt).iIndex).Visible = True
            chkCorner(uIcmFileData(iIcmCnt).iIndex).Caption = uIcmFileData(iIcmCnt).iHyozi & "���@"
        End If
     Next
    
    On Error GoTo OtherError
       
    'DB�ڑ�
    '�A�v�����O�p
    cnConn.ConnectionString = "File Name=" & PATH_IDU_APP & PATH_IDU_APPLOG
    cnConn.Open
        
    On Error GoTo 0
    
    iNowChk1 = 1
    iNowChk2 = 2
    
    '���X�g�̏����\��
    If sSetListBox = False Then
        '�uID���p���j�b�g���O�Ǘ��F�A�v�����O�\���ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
        '���X�g�{�b�N�X�̏�����
        LstFile.Clear
        MsgBox "���O�ꗗ�̎擾�Ɏ��s���܂����B", vbCritical, "�\���ُ�"
   End If
   
   '�uID���p���j�b�g���O�Ǘ���ʁF�\���v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, IDU_LOG_KANRI_GAMEN_START, 0)
   
   
 Exit Sub
    
FileError:
    Select Case iErr
    Case 1:
       '�uID���p���j�b�g���O�Ǘ��FINI�t�@�C���ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
     Case 2:
       '�uID���p���j�b�g���O�Ǘ��FINI�t�@�C���ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
     Case 3:
       '�uID���p���j�b�g���O�Ǘ��FINI�t�@�C���ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_INIFILE_ERROR, 0)
     End Select
   MsgBox "INI�t�@�C���̎擾�Ɏ��s���܂����", vbCritical, "�t�@�C���ُ�"
   
   Exit Sub
OtherError:
   '�uID���p���j�b�g���O�Ǘ��F�A�v�����O�\���ُ�v���O�o��
   Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
   LstFile.Clear
   MsgBox "���O�ꗗ�̎擾�Ɏ��s���܂����B", vbCritical, "�\���ُ�"
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdLogHyouzi_Click
'//  �@�\����  : �u���O�\��(�e�L�X�g�\���j�v�t����������
'//  �@�\�T�v  : �I���t�@�C�����A�e�L�X�g�ɂĕ\������B
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
    Dim lngErrCode As Long   '�G���[�R�[�h

   '�uID���p���j�b�g���O�Ǘ��F���O�\���t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_TEXT_HYOUJI_BUTTOM, 0)

    '���O�����f�[�^�������`�F�b�N
    bRet = fLogSearchCheck
    If bRet = False Then                                '���O�����f�[�^�ɃG���[������ꍇ�A�����I��
        Exit Sub
    End If

    '���O�e�L�X�g�t�@�C������������
    bRet = fWriteLogtxt
    If bRet = True Then                                 '���O�e�L�X�g�t�@�C��������ɍ쐬���ꂽ�ꍇ
        '�uID���p���j�b�g���O�Ǘ��F���O�e�L�X�g�t�@�C���쐬����v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_CREATE_TEXT_HYOUJI, 0)
        '�t�@�C���R�s�[
        sFileName = Trim(Left(LstFile.List(LstFile.ListIndex), 12))
        sFileName = PATH_IDU_APP & PATH_IDU_WORK & "\\" & Left(sFileName, Len(sFileName) - 4) & ".txt"
        '�t�@�C���I�[�v��
        On Error GoTo FileError
        sCommand = MN_EXE_MEMO & sFileName              '���s�R�}���h���쐬����
        lRetVal = Shell(sCommand, vbMaximizedFocus)     '�m�[�g�p�b�h���N������
        AppActivate lRetVal, True                       '�A�N�e�B�u�i�O�ʕ\���j�ɂ���
        SendKeys "{LEFT}", True
        On Error GoTo 0
        '�uID���p���j�b�g���O�Ǘ��F���O�e�L�X�g�\������v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LOG_KANRI_TEXT_HYOUJI_OK, 0)
    Else
        '�uID���p���j�b�g���O�Ǘ��F�o�̓f�[�^�쐬�ُ�v���O�o��
        lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_KANRI_CREATE_TEXT_ERROR, lngErrCode)
       '�u�f�[�^�o�͎��s�v�|�b�v�A�b�v�\��
       MsgBox "�}�̏o�͂���f�[�^�̍쐬�Ɏ��s���܂����B", vbCritical, "�f�[�^�o�͎��s"
    End If
    Exit Sub

FileError:
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   '�uID���p���j�b�g���O�Ǘ��F���O�e�L�X�g�\�������ُ�v���O�o��
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, LOG_KANRI_TEXT_HYOUJI_ERROR, lngErrCode)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdLog_Click
'//  �@�\����  : �u���O�}�̏o�́v�t����������
'//  �@�\�T�v  : �I���t�@�C�����A�w��t�H���_�֏o�͂���B
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
'//     REVISIONS :(EG20 V2.0.1.1) 2011-12-02   REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή��y�Ď�D-115�z
'//                 �@�E�������ʃ��b�Z�[�W�{�b�N�X�̕����ύX
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

    '�uID���p���j�b�g���O�Ǘ���ʁF���O�}�̏o�͖t�����v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_OUTPUT_BUTTOM, 0)

    Dim bFrmShow As Boolean
    bFrmShow = False

    txtDummy.SetFocus

    On Error GoTo EVENTLOG_ERROR
    If iNowChk1 = 1 Then
        MyPath = PATH_IDU_LOG & PATH_IDU_LOG_APP
    Else
        MyPath = PATH_IDU_LOG & PATH_IDU_LOG_HOSHU
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
            frmIDULogkanri.Refresh
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
'   '�����ԍ��i�[�i�������j
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

'V1.6.0.1 ADD START

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
    
    '�R�s�[��t�H���_�p�X�쐬(�w��t�H���_��IDULOG)
    sWriteDir = sWriteDir & "\" & IDU_LOGKANRI_IDULOG
    
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
'EG20 V2.0.1.1�y�Ď�D-115�zDEL START
'        MsgBox "�g�c�c���ꎞ�t�H���_�ւ̏o�͂͐���I�����܂����B", vbInformation + vbOKOnly, "�o�͌���"
'EG20 V2.0.1.1�y�Ď�D-115�zDEL END
'EG20 V2.0.1.1�y�Ď�D-115�zADDL START
        MsgBox "����I�����܂����B", vbInformation + vbOKOnly, "�o�͌���"
'EG20 V2.0.1.1�y�Ď�D-115�zADDL END
    End If
    
    '�uID���p���j�b�g���O�Ǘ���ʁF���O�}�̏o�͏�������v
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, LOG_OUTPUT_OK, 0)

    Exit Sub

EVENTLOG_ERROR:
   'V1.6.0.1 ADD START
       '�t�@�C���V�X�e���I�u�W�F�N�g���
      Set objFso = Nothing
      '�uID���p���j�b�g���O�Ǘ���ʁF�t�H���_�쐬�ُ�v
      Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_CREATE_LOGFOLDER_ERROR, 0)
   'V1.6.0.1 ADD END
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��J�n
    '�v���O���X�o�[����������
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1�y�v���O���X�o�[�\���@�\�������Ή��z�ǉ��I��
    
    If UCase(Left(sWriteDir, 1)) = "A" Then
        MsgBox "�e�c�o�ُ͈͂�I�����܂����B", vbCritical, "�o�͌���"
    Else
'EG20 V2.0.1.1�y�Ď�D-115�zDEL START
'        MsgBox "�g�c�c���ꎞ�t�H���_�ւ̏o�ُ͈͂�I�����܂����B", vbCritical, "�o�͌���"
'EG20 V2.0.1.1�y�Ď�D-115�zDEL END
'EG20 V2.0.1.1�y�Ď�D-115�zADDL START
        MsgBox "�ُ�I�����܂����B", vbCritical, "�o�͌���"
'EG20 V2.0.1.1�y�Ď�D-115�zADDL END
    End If
    
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    '�uID���p���j�b�g���O�Ǘ���ʁF���O�}�̏o�͏����ُ�v
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
    Dim udtMail As IDU_LDU_LGCHGREQ_CMD     '���O�ؑ֗v��
    Dim lngErrCode As Long                  '�G���[�R�[�h
    Dim bFlag As Boolean                    '���[����M�t���O
    Dim lId As Long                         '���[��ID

    On Error Resume Next

    LstFile.Clear

    '�uID���p���j�b�g���O�Ǘ���ʁF���O�֖ؑt�����v
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
    bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.udtlHeader)
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
'//  �֐�����  : cmdInstall_Click
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
Private Sub cmdInstall_Click()
   On Error Resume Next
  
   '�u�}�̎�O�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '�}�̎�O����
    Call pfRemove(Me)
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
   
   '�uID���p���j�b�g���O�Ǘ���ʁF�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, IDU_LOG_KANRI_GAMEN_END, 0)
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub optApp_Click()

    On Error GoTo Err_mgs
    
   '�uID���p���j�b�g���O�Ǘ���ʁF�A�v�����O�v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_CHANGE_APLLOG, 0)
    
    '�I������Ă����̂��A�ێ�v���O�������O�������ꍇ
    If iNowChk1 <> 1 Then
        'DB�̐ڑ���؂�ւ���
        If iNowChk1 <> 0 Then
            If Not cnConn2 Is Nothing Then
                cnConn2.Close
            End If
        End If
        '�ڑ����i�V�ɂ���
        iNowChk1 = 0

        cnConn.ConnectionString = "File Name=" & PATH_IDU_APP & PATH_IDU_APPLOG
        cnConn.Open
        
        '�I������Ă���`�F�b�N��ێ�����
        iNowChk1 = 1
        
        '��\���ɂȂ��Ă������ڂ�\��������
        frmMod.Visible = True
        cmdZSentaku.Visible = True
        cmdZHisentaku.Visible = True
        cmdHSentaku.Visible = True
        cmdHHisentaku.Visible = True
        tabCorner.Visible = True
        
        '�\�����ēǂݍ��݂���
        If sSetListBox = False Then
            '�uID���p���j�b�g���O�Ǘ��F�A�v�����O�\���ُ�v���O�o��
             Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_APLLOG_ERROR, 0)
            '���X�g�{�b�N�X�̏�����
            LstFile.Clear
            MsgBox "���O�ꗗ�̎擾�Ɏ��s���܂����B", vbCritical, "�\���ُ�"
        End If
    End If
   
   '�uID���p���j�b�g���O�Ǘ��F�A�v�����O�\������v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_APLLOG_OK, 0)

   Exit Sub
    
Err_mgs:
   '�uID���p���j�b�g���O�Ǘ��F�A�v�����O�\���ُ�v���O�o��
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
    
   '�uID���p���j�b�g���O�Ǘ���ʁF�ێ烍�O�v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_CHANGE_HOSHULOG, 0)
    
    '�I������Ă����̂��A�A�v���P�[�V�������O�������ꍇ
    If iNowChk1 <> 2 Then
        'DB�̐ڑ���؂�ւ���
        If iNowChk1 <> 0 Then
            If Not cnConn Is Nothing Then
                cnConn.Close
            End If
        End If
        '�ڑ����i�V�ɂ���
        iNowChk1 = 0

        cnConn2.ConnectionString = "File Name=" & PATH_IDU_APP & PATH_IDU_HOSHULOG
        cnConn2.Open
            
        '�I������Ă���`�F�b�N��ێ�����
        iNowChk1 = 2
        
        '�\���ɂȂ��Ă��鍀�ڂ��\���ɂ���
        frmMod.Visible = False
        cmdZSentaku.Visible = False
        cmdZHisentaku.Visible = False
        cmdHSentaku.Visible = False
        cmdHHisentaku.Visible = False
        tabCorner.Visible = False
        
        '�\�����ēǂݍ��݂���
        If sSetListBox = False Then
            '�uID���p���j�b�g���O�Ǘ��F�ێ烍�O�\���ُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_KANRI_HODHULOG_ERROR, 0)
            '���X�g�{�b�N�X�̏�����
            LstFile.Clear
            MsgBox "���O�ꗗ�̎擾�Ɏ��s���܂����B", vbCritical, "�\���ُ�"
        End If
    End If
     
    '�uID���p���j�b�g���O�Ǘ��F�ێ烍�O�\������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, LOG_KANRI_HODHULOG_OK, 0)
    
    Exit Sub

Err_mgs:
    '�uID���p���j�b�g���O�Ǘ��F�ێ烍�O�\���ُ�v���O�o��
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
    
    For iCnt = 0 To iModCnt
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
        
    For iCnt = 0 To iModCnt
        chkMod(iCnt).Value = 0
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdModHi_Click
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
        '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
        '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
'//     REVISIONS :(EG20 V2.1.0.1) 2011-12-14 REVISED BY [TCC] M.Matsumoto
'//                �y��-336�Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function TextTime_Check(sType As String, sTxt As String)
    Dim iChk As Integer
    Dim iRet As Integer
    Dim sChk As String
    
    Dim k As Integer                            'EG20 V2.0.1.1 ADD
    
    '�߂�l�Ɉُ���Z�b�g
    TextTime_Check = False
        
    If Trim(sTxt) <> "" Then
        iChk = Val(sTxt)
        'EG20 V2.0.1.1 ADD START �y��-336�Ή��z
        '���͂��ꂽ���ɐ��l�ȊO�̕��������݂���ꍇ�́A�G���[
        For k = 1 To Len(sTxt)
            If Not Mid(sTxt, k, 1) Like "[0-9]" Then
                iRet = MsgBox("���͂��ꂽ�����͐����ł͂���܂���B", vbExclamation, "���ُ͈�")
                '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
                Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                Exit Function
            End If
        Next k
        'EG20 V2.0.1.1 ADD END
                    
        'EG20 V2.0.1.1 DEL START �y��-336�Ή��z
'        If iChk = 0 And sType <> "Hour" And sType <> "Minutes" Then
'            iRet = MsgBox("���͂��ꂽ�����͐����ł͂���܂���B", vbExclamation, "���ُ͈�")
'            '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
'            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
            
            'EG20 V2.0.1.1 DEL START �y��-336�Ή��z
            '����
'            If Len(Trim(str(iChk))) <> Len(sTxt) Then
'                iRet = MsgBox("���͂��ꂽ�����ɐ����ȊO�̂��̂��܂܂�Ă��܂��B", vbExclamation, "���ُ͈�")
'                '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
'                Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
                        '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
                        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                        Exit Function
                    End If
                Case "Month"
                    '��
                    If iChk < 1 Or iChk > 12 Then
                        iRet = MsgBox("���w��͈̔͂𒴂��Ă��܂��B", vbExclamation, "���ُ͈�")
                        '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
                        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                        Exit Function
                    End If
                Case "Day"
                    '��
                    If iChk < 1 Or iChk > 31 Then
                        iRet = MsgBox("���w��͈̔͂𒴂��Ă��܂��B", vbExclamation, "���ُ͈�")
                        '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
                        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                        Exit Function
                    End If
                Case "Hour"
                    '��
                    If iChk < 0 Or iChk > 23 Then
                        iRet = MsgBox("���Ԏw��͈̔͂𒴂��Ă��܂��B", vbExclamation, "���ُ͈�")
                        '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
                        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                        Exit Function
                    End If
                Case "Minutes"
                    '��
                    If iChk < 0 Or iChk > 59 Then
                        iRet = MsgBox("���Ԏw��͈̔͂𒴂��Ă��܂��B", vbExclamation, "���ُ͈�")
                        '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
                        Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                        Exit Function
                    End If
            End Select
'        End If             'EG20 V2.0.1.1 DEL �y��-336�Ή��z
    End If
    
    '�߂�l�ɐ����Ԃ�
    TextTime_Check = True
    Exit Function
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdZSentaku_Click
'//  �@�\����  : �u�S�R�[�i�[�@�S���@�I���v�t����������
'//  �@�\�T�v  : �\�����X�V����B
'//�@�@�@�@�@�@�@�\�����@�w�蕔�F
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
Private Sub cmdZSentaku_Click()
    Dim iCnt As Integer
        
    For iCnt = 0 To 95
        chkCorner(iCnt).Value = 1
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdZHisentaku_Click
'//  �@�\����  : �u�S�R�[�i�[�@�S���@��I���v�t����������
'//  �@�\�T�v  : �\�����X�V����B
'//�@�@�@�@�@�@�@�\�����@�w�蕔�F
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
Private Sub cmdZHisentaku_Click()
    Dim iCnt As Integer
        
    For iCnt = 0 To 95
        chkCorner(iCnt).Value = 0
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdHSentaku_Click
'//  �@�\����  : �u�\���R�[�i�[�@�S���@�I���v�t����������
'//  �@�\�T�v  : �\�����X�V����B
'//�@�@�@�@�@�@�@�\�����@�w�蕔�F
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
Private Sub cmdHSentaku_Click()
    Dim iCnt As Integer
    Dim iMin As Integer
    Dim iMax As Integer
        
    '�ŏ��l�A�ő�l�擾
    iMin = tabCorner.Tab * 16
    iMax = tabCorner.Tab * 16 + 15
        For iCnt = iMin To iMax
            chkCorner(iCnt).Value = 1
        Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdHHisentaku_Click
'//  �@�\����  : �u�\���R�[�i�[�@�S���@��I���v�t����������
'//  �@�\�T�v  : �\�����X�V����B
'//�@�@�@�@�@�@�@�\�����@�w�蕔�F
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
Private Sub cmdHHisentaku_Click()
    Dim iCnt As Integer
    Dim iMin As Integer
    Dim iMax As Integer
        
    '�ŏ��l�A�ő�l�擾
    iMin = tabCorner.Tab * 16
    iMax = tabCorner.Tab * 16 + 15
        For iCnt = iMin To iMax
            chkCorner(iCnt).Value = 0
        Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Unload
'//  �@�\����  : �u���j���[��ʂ֖߂�v�t����������
'//  �@�\�T�v  : �@DB�ڑ��̉�����s���B
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
Private Sub Form_Unload(Cancel As Integer)
    If iNowChk1 = 1 Then
        If Not cnConn Is Nothing Then
            cnConn.Close
        End If
    End If
    If iNowChk1 = 2 Then
        If Not cnConn2 Is Nothing Then
            cnConn2.Close
        End If
    End If

    'RecordSet��`������������̍폜����
    Set rsRecordSet = Nothing
    'Connection��`������������폜����
    Set cnConn = Nothing
    'Connection2��`������������폜����
    Set cnConn2 = Nothing
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSetListBox
'//  �@�\����  : ���O�t�@�C����o�^
'//  �@�\�T�v  : ���O�t�@�C�������X�g�{�b�N�X�ɓo�^����B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F��������
'//�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@  �\�����O���W�I�t����������
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
Private Function sSetListBox()
    Dim i As Integer
    Dim iCnt As Integer
    Dim strSQL As String
    Dim datWork As Date
    Dim sEntry As String        '�ҏW������

    On Error GoTo Err_mgs

    sSetListBox = False
    '�_�����@�̃��O���擾�̂r�p�k��
    strSQL = "Select LOG_NAME,LOG_START_TIME,LOG_END_TIME,LOG_SIZE" _
            & " from T_LOG"

    On Error Resume Next            ' �G���[�̃g���b�v�𗯕ۂ��܂��B
    Err.Clear

    '�A�v���A�ێ�`�F�b�N
    Select Case iNowChk1
        Case 1
            rsRecordSet.Open strSQL, cnConn
        Case 2
            rsRecordSet.Open strSQL, cnConn2
        Case Else
            Exit Function
    End Select

    '�r�p�k���s�G���[�������ꍇ
    If Err.Number <> 0 Then
        '���R�[�h�Z�b�g�̃N���[�Y
        rsRecordSet.Close

        GoTo Err_mgs
    End If
    i = 0
   '���O�����\���̔z��(gtypLogData)�Ɋi�[����
    Do While Not rsRecordSet.EOF
        ReDim Preserve gLogData(i)
        gLogData(i).sName = rsRecordSet!LOG_NAME
        gLogData(i).sStTime = Format(rsRecordSet!LOG_START_TIME, "yyyy/mm/dd hh:mm:ss")
        gLogData(i).sEdTime = Format(rsRecordSet!LOG_END_TIME, "yyyy/mm/dd hh:mm:ss")
        gLogData(i).iSize = rsRecordSet!LOG_SIZE

        rsRecordSet.MoveNext
        i = i + 1
    Loop
    iCnt = i

    '�r�p�k���s�G���[�������ꍇ
    If Err.Number <> 0 Then
        '���R�[�h�Z�b�g�̃N���[�Y
        rsRecordSet.Close

        GoTo Err_mgs
    End If

    '���R�[�h�Z�b�g�̃N���[�Y
    rsRecordSet.Close


    On Error GoTo Err_mgs
    '�u���O�t�@�C���v���X�g�{�b�N�X���N���A����
    LstFile.Clear

    '���O�t�@�C������ҏW����
    For i = 0 To iCnt - 1
        sEntry = Left(gLogData(i).sName, 12)
        If Len(gLogData(i).sStTime) = 19 Then
            sEntry = sEntry & "  " & gLogData(i).sStTime
        Else
            sEntry = sEntry & "                     "
        End If

        If Len(gLogData(i).sEdTime) = 19 Then
            sEntry = sEntry & "  " & gLogData(i).sEdTime
        Else
            sEntry = sEntry & "                     "
        End If

        sEntry = sEntry & "  " & Format(gLogData(i).iSize, "@@@@@@@@")
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
            '�uID���p���j�b�g���O�Ǘ���ʁFDB�A�N�Z�X�ُ�v
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, IDU_LOG_KANRI_DB_ACCESS_ERROR, 0)
        Case 2
            '�uID���p���j�b�g���O�Ǘ���ʁFDB�A�N�Z�X�ُ�v
            Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, IDU_LOG_KANRI_DB_ACCESS_ERROR, 0)
    End Select
End Function

'///////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2005 All Right Reserved
'/
'/  �֐�����  : fLogSearchCheck
'/  �T�v     : ���O�����f�[�^�`�F�b�N
'/  ����     : ���O�����f�[�^�̐��������`�F�b�N����
'/  ���Ұ�   :
'/           :
'/
'/  ORIGINAL  �F(1.0.0.1) 2005-01-27  CODED BY  [TCC] T.Yashiro
'//     REVISIONS :(EG20 3.6.0.1) 2012-02-23   REVISED BY [TCC] T.Furuya
'//                 EG20 �t�F�[�Y�Q �c�����
'/  REVISIONS �F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'/  ���l�F
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
   
    On Error Resume Next

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
                    '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
                    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
'                    txtStFun.Text = "00"
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
                If Len(Trim(txtStNen.Text)) = 0 And _
                   Len(Trim(txtStTuki.Text)) = 0 And _
                   Len(Trim(txtStHi.Text)) = 0 And _
                   Len(Trim(txtStZi.Text)) = 0 Then
            
                    txtStFun.Text = "00"
            
                Else
                    iRet = MsgBox("�\���͈͂̊J�n�ɖ����͂̍��ڂ�����܂��B", vbExclamation, "���ُ͈�")
                    '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
                    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
'                    txtEdZi.Text = "00"
'EG20 V2.0.1.1 DEL END
'EG20 V2.0.1.1 ADD START
                If Len(Trim(txtEdNen.Text)) = 0 And _
                   Len(Trim(txtEdTuki.Text)) = 0 And _
                   Len(Trim(txtEdHi.Text)) = 0 Then

                    txtEdZi.Text = "00"
                Else
                    iRet = MsgBox("�\���͈͂̏I���ɖ����͂̍��ڂ�����܂��B", vbExclamation, "���ُ͈�")
                    '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
                    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
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
                    '�uID���p���j�b�g���O�Ǘ��F���͎����ُ�v���O�o��
                    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, LOG_JIKOKU_ERROR, 0)
                    Exit Function
                End If
'EG20 V2.0.1.1 ADD END
            Else
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
        For i = 0 To iIcmCnt
            '�`�F�b�N���n�m�Ȃ珈������
            If chkCorner(uIcmFileData(i).iIndex).Value = 1 Then
                '�t���O�𗧂Ă�
                bFlg = True
            End If
        Next
        
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
'//  �@�\����  : ���O�e�L�X�g�t�@�C���������ݏ���
'//  �@�\�T�v  : ���O�t�@�C�����e�L�X�g�t�@�C���ɏ�������
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
        MyPath = PATH_IDU_LOG & PATH_IDU_LOG_APP
    Else
        MyPath = PATH_IDU_LOG & PATH_IDU_LOG_HOSHU
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
    iStatus = dllCreateDispLogFile(lErr, sFileName, uLogConv, sObjectTopFile, PATH_IDU_APP)
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
'//  �@�\�T�v  : ���O�g���[�X��ʂ���A���O�ϊ������쐬����B
'//�@�@�@�@�@�@�@�\���t�@�C���w�蕔�F�u���O�\��(�e�L�X�g�\��)�v�t����������
'//
'//              �^        �@�@�@�@�@�@����      �Ӗ�
'//  ����      : VB_LOG_DISP_SETTING�@uLogConv�@[OUT]���O�ϊ����
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
    Dim bGokiFlg As Boolean
   
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

        '�S���`�F�b�N
         bGokiFlg = False
        For i = 0 To iIcmCnt
            '�`�F�b�N���n�m�Ȃ珈������
            If chkCorner(uIcmFileData(i).iIndex).Value = 1 Then
                If uIcmFileData(i).iRonri = 32 Then
                    '�t���O�𗧂Ă�
                    bGokiFlg = True
                Else
                    '�r�b�g�J�E���g�v�Z
                    iBitCnt = 1
                    If uIcmFileData(i).iRonri <> 1 Then
                        For ii = 1 To uIcmFileData(i).iRonri - 1
                            iBitCnt = iBitCnt * 2
                        Next
                    End If
                    '�ϐ��ɒǉ�����
                    uLogConv.Goki = uLogConv.Goki + iBitCnt
                End If
            End If
        Next

        If bGokiFlg = True Then
            uLogConv.Goki = -2147483648# + uLogConv.Goki
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
             AppActivate frmIDULogkanri.Caption, False      ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
             pfFormActive (frmIDULogkanri.hwnd)
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
        AppActivate frmIDULogkanri.Caption, False
        pfFormActive (frmIDULogkanri.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END
