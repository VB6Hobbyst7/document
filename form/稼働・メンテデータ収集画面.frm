VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSyusyu 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�ғ��E�����e�f�[�^���W�i�����㎩�����D�@�j"
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
      Height          =   1095
      Left            =   360
      TabIndex        =   11
      Top             =   840
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
      Height          =   1095
      Left            =   2640
      TabIndex        =   10
      Top             =   840
      Width           =   2000
   End
   Begin VB.CommandButton cmdHSentaku 
      Caption         =   " �\���R�[�i   �S���@ �I��"
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
      Left            =   4920
      TabIndex        =   9
      Top             =   840
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
      Height          =   1095
      Left            =   7200
      TabIndex        =   8
      Top             =   840
      Width           =   2000
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   8400
      Top             =   7920
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "  �f�[�^���W�E�o��    ��ʂ֖߂�"
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
   Begin VB.CommandButton cmdSyusyu 
      Caption         =   " ���W "
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
      Left            =   360
      TabIndex        =   0
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton cmdFDWrite 
      Caption         =   "�}�̏o��"
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
      Left            =   2880
      TabIndex        =   1
      Top             =   7680
      Width           =   2175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8281
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   970
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " ������������ ������������"
      TabPicture(0)   =   "�ғ��E�����e�f�[�^���W���.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " ������������ ������������"
      TabPicture(1)   =   "�ғ��E�����e�f�[�^���W���.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   " ������������ ������������"
      TabPicture(2)   =   "�ғ��E�����e�f�[�^���W���.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   " ������������ ������������"
      TabPicture(3)   =   "�ғ��E�����e�f�[�^���W���.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   " ������������ ������������"
      TabPicture(4)   =   "�ғ��E�����e�f�[�^���W���.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame2(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   " ������������ ������������"
      TabPicture(5)   =   "�ғ��E�����e�f�[�^���W���.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame2(5)"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "�w�荆�@�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Index           =   5
         Left            =   -74880
         TabIndex        =   174
         Top             =   120
         Width           =   10935
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   95
            Left            =   9480
            Style           =   1  '���̨���
            TabIndex        =   190
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   94
            Left            =   8160
            Style           =   1  '���̨���
            TabIndex        =   189
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   93
            Left            =   6840
            Style           =   1  '���̨���
            TabIndex        =   188
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   92
            Left            =   5520
            Style           =   1  '���̨���
            TabIndex        =   187
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   91
            Left            =   4200
            Style           =   1  '���̨���
            TabIndex        =   186
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   90
            Left            =   2880
            Style           =   1  '���̨���
            TabIndex        =   185
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   89
            Left            =   1560
            Style           =   1  '���̨���
            TabIndex        =   184
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   88
            Left            =   240
            Style           =   1  '���̨���
            TabIndex        =   183
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   87
            Left            =   9480
            Style           =   1  '���̨���
            TabIndex        =   182
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   86
            Left            =   8160
            Style           =   1  '���̨���
            TabIndex        =   181
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   85
            Left            =   6840
            Style           =   1  '���̨���
            TabIndex        =   180
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   84
            Left            =   5520
            Style           =   1  '���̨���
            TabIndex        =   179
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   83
            Left            =   4200
            Style           =   1  '���̨���
            TabIndex        =   178
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   82
            Left            =   2880
            Style           =   1  '���̨���
            TabIndex        =   177
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   81
            Left            =   1560
            Style           =   1  '���̨���
            TabIndex        =   176
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   80
            Left            =   240
            Style           =   1  '���̨���
            TabIndex        =   175
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   95
            Left            =   9480
            TabIndex        =   206
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   94
            Left            =   8160
            TabIndex        =   205
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   93
            Left            =   6840
            TabIndex        =   204
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   92
            Left            =   5520
            TabIndex        =   203
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   91
            Left            =   4200
            TabIndex        =   202
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   90
            Left            =   2880
            TabIndex        =   201
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   89
            Left            =   1560
            TabIndex        =   200
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   88
            Left            =   240
            TabIndex        =   199
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   87
            Left            =   9480
            TabIndex        =   198
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   86
            Left            =   8160
            TabIndex        =   197
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   85
            Left            =   6840
            TabIndex        =   196
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   84
            Left            =   5520
            TabIndex        =   195
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   83
            Left            =   4200
            TabIndex        =   194
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   82
            Left            =   2880
            TabIndex        =   193
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   81
            Left            =   1560
            TabIndex        =   192
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   80
            Left            =   240
            TabIndex        =   191
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "�w�荆�@�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Index           =   4
         Left            =   -74880
         TabIndex        =   141
         Top             =   120
         Width           =   10935
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   79
            Left            =   9480
            Style           =   1  '���̨���
            TabIndex        =   157
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   78
            Left            =   8160
            Style           =   1  '���̨���
            TabIndex        =   156
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   77
            Left            =   6840
            Style           =   1  '���̨���
            TabIndex        =   155
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   76
            Left            =   5520
            Style           =   1  '���̨���
            TabIndex        =   154
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   75
            Left            =   4200
            Style           =   1  '���̨���
            TabIndex        =   153
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   74
            Left            =   2880
            Style           =   1  '���̨���
            TabIndex        =   152
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   73
            Left            =   1560
            Style           =   1  '���̨���
            TabIndex        =   151
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   72
            Left            =   240
            Style           =   1  '���̨���
            TabIndex        =   150
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   71
            Left            =   9480
            Style           =   1  '���̨���
            TabIndex        =   149
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   70
            Left            =   8160
            Style           =   1  '���̨���
            TabIndex        =   148
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   69
            Left            =   6840
            Style           =   1  '���̨���
            TabIndex        =   147
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   68
            Left            =   5520
            Style           =   1  '���̨���
            TabIndex        =   146
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   67
            Left            =   4200
            Style           =   1  '���̨���
            TabIndex        =   145
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   66
            Left            =   2880
            Style           =   1  '���̨���
            TabIndex        =   144
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   65
            Left            =   1560
            Style           =   1  '���̨���
            TabIndex        =   143
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   64
            Left            =   240
            Style           =   1  '���̨���
            TabIndex        =   142
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   79
            Left            =   9480
            TabIndex        =   173
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   78
            Left            =   8160
            TabIndex        =   172
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   77
            Left            =   6840
            TabIndex        =   171
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   76
            Left            =   5520
            TabIndex        =   170
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   75
            Left            =   4200
            TabIndex        =   169
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   74
            Left            =   2880
            TabIndex        =   168
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   73
            Left            =   1560
            TabIndex        =   167
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   72
            Left            =   240
            TabIndex        =   166
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   71
            Left            =   9480
            TabIndex        =   165
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   70
            Left            =   8160
            TabIndex        =   164
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   69
            Left            =   6840
            TabIndex        =   163
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   68
            Left            =   5520
            TabIndex        =   162
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   67
            Left            =   4200
            TabIndex        =   161
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   66
            Left            =   2880
            TabIndex        =   160
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   65
            Left            =   1560
            TabIndex        =   159
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   64
            Left            =   240
            TabIndex        =   158
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "�w�荆�@�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Index           =   3
         Left            =   -74880
         TabIndex        =   108
         Top             =   120
         Width           =   10935
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   63
            Left            =   9480
            Style           =   1  '���̨���
            TabIndex        =   124
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   62
            Left            =   8160
            Style           =   1  '���̨���
            TabIndex        =   123
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   61
            Left            =   6840
            Style           =   1  '���̨���
            TabIndex        =   122
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   60
            Left            =   5520
            Style           =   1  '���̨���
            TabIndex        =   121
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   59
            Left            =   4200
            Style           =   1  '���̨���
            TabIndex        =   120
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   58
            Left            =   2880
            Style           =   1  '���̨���
            TabIndex        =   119
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   57
            Left            =   1560
            Style           =   1  '���̨���
            TabIndex        =   118
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   56
            Left            =   240
            Style           =   1  '���̨���
            TabIndex        =   117
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   55
            Left            =   9480
            Style           =   1  '���̨���
            TabIndex        =   116
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   54
            Left            =   8160
            Style           =   1  '���̨���
            TabIndex        =   115
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   53
            Left            =   6840
            Style           =   1  '���̨���
            TabIndex        =   114
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   52
            Left            =   5520
            Style           =   1  '���̨���
            TabIndex        =   113
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   51
            Left            =   4200
            Style           =   1  '���̨���
            TabIndex        =   112
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   50
            Left            =   2880
            Style           =   1  '���̨���
            TabIndex        =   111
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   49
            Left            =   1560
            Style           =   1  '���̨���
            TabIndex        =   110
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   48
            Left            =   240
            Style           =   1  '���̨���
            TabIndex        =   109
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   63
            Left            =   9480
            TabIndex        =   140
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   62
            Left            =   8160
            TabIndex        =   139
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   61
            Left            =   6840
            TabIndex        =   138
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   60
            Left            =   5520
            TabIndex        =   137
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   59
            Left            =   4200
            TabIndex        =   136
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   58
            Left            =   2880
            TabIndex        =   135
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   57
            Left            =   1560
            TabIndex        =   134
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   56
            Left            =   240
            TabIndex        =   133
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   55
            Left            =   9480
            TabIndex        =   132
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   54
            Left            =   8160
            TabIndex        =   131
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   53
            Left            =   6840
            TabIndex        =   130
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   52
            Left            =   5520
            TabIndex        =   129
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   51
            Left            =   4200
            TabIndex        =   128
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   50
            Left            =   2880
            TabIndex        =   127
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   49
            Left            =   1560
            TabIndex        =   126
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   48
            Left            =   240
            TabIndex        =   125
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "�w�荆�@�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Index           =   2
         Left            =   -74880
         TabIndex        =   75
         Top             =   120
         Width           =   10935
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   47
            Left            =   9480
            Style           =   1  '���̨���
            TabIndex        =   91
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   46
            Left            =   8160
            Style           =   1  '���̨���
            TabIndex        =   90
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   45
            Left            =   6840
            Style           =   1  '���̨���
            TabIndex        =   89
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   44
            Left            =   5520
            Style           =   1  '���̨���
            TabIndex        =   88
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   43
            Left            =   4200
            Style           =   1  '���̨���
            TabIndex        =   87
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   42
            Left            =   2880
            Style           =   1  '���̨���
            TabIndex        =   86
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   41
            Left            =   1560
            Style           =   1  '���̨���
            TabIndex        =   85
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   40
            Left            =   240
            Style           =   1  '���̨���
            TabIndex        =   84
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   39
            Left            =   9480
            Style           =   1  '���̨���
            TabIndex        =   83
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   38
            Left            =   8160
            Style           =   1  '���̨���
            TabIndex        =   82
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   37
            Left            =   6840
            Style           =   1  '���̨���
            TabIndex        =   81
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   36
            Left            =   5520
            Style           =   1  '���̨���
            TabIndex        =   80
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   35
            Left            =   4200
            Style           =   1  '���̨���
            TabIndex        =   79
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   34
            Left            =   2880
            Style           =   1  '���̨���
            TabIndex        =   78
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   33
            Left            =   1560
            Style           =   1  '���̨���
            TabIndex        =   77
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   32
            Left            =   240
            Style           =   1  '���̨���
            TabIndex        =   76
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   47
            Left            =   9480
            TabIndex        =   107
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   46
            Left            =   8160
            TabIndex        =   106
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   45
            Left            =   6840
            TabIndex        =   105
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   44
            Left            =   5520
            TabIndex        =   104
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   43
            Left            =   4200
            TabIndex        =   103
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   42
            Left            =   2880
            TabIndex        =   102
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   41
            Left            =   1560
            TabIndex        =   101
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   40
            Left            =   240
            TabIndex        =   100
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   39
            Left            =   9480
            TabIndex        =   99
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   38
            Left            =   8160
            TabIndex        =   98
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   37
            Left            =   6840
            TabIndex        =   97
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   36
            Left            =   5520
            TabIndex        =   96
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   35
            Left            =   4200
            TabIndex        =   95
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   34
            Left            =   2880
            TabIndex        =   94
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   33
            Left            =   1560
            TabIndex        =   93
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   32
            Left            =   240
            TabIndex        =   92
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "�w�荆�@�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Index           =   1
         Left            =   -74880
         TabIndex        =   42
         Top             =   120
         Width           =   10935
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   31
            Left            =   9480
            Style           =   1  '���̨���
            TabIndex        =   58
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   30
            Left            =   8160
            Style           =   1  '���̨���
            TabIndex        =   57
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   29
            Left            =   6840
            Style           =   1  '���̨���
            TabIndex        =   56
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   28
            Left            =   5520
            Style           =   1  '���̨���
            TabIndex        =   55
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   27
            Left            =   4200
            Style           =   1  '���̨���
            TabIndex        =   54
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   26
            Left            =   2880
            Style           =   1  '���̨���
            TabIndex        =   53
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   25
            Left            =   1560
            Style           =   1  '���̨���
            TabIndex        =   52
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   24
            Left            =   240
            Style           =   1  '���̨���
            TabIndex        =   51
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   23
            Left            =   9480
            Style           =   1  '���̨���
            TabIndex        =   50
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   22
            Left            =   8160
            Style           =   1  '���̨���
            TabIndex        =   49
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   21
            Left            =   6840
            Style           =   1  '���̨���
            TabIndex        =   48
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   20
            Left            =   5520
            Style           =   1  '���̨���
            TabIndex        =   47
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   19
            Left            =   4200
            Style           =   1  '���̨���
            TabIndex        =   46
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   18
            Left            =   2880
            Style           =   1  '���̨���
            TabIndex        =   45
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   17
            Left            =   1560
            Style           =   1  '���̨���
            TabIndex        =   44
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   16
            Left            =   240
            Style           =   1  '���̨���
            TabIndex        =   43
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   31
            Left            =   9480
            TabIndex        =   74
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   30
            Left            =   8160
            TabIndex        =   73
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   29
            Left            =   6840
            TabIndex        =   72
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   28
            Left            =   5520
            TabIndex        =   71
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   27
            Left            =   4200
            TabIndex        =   70
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   26
            Left            =   2880
            TabIndex        =   69
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   25
            Left            =   1560
            TabIndex        =   68
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   24
            Left            =   240
            TabIndex        =   67
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   23
            Left            =   9480
            TabIndex        =   66
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   22
            Left            =   8160
            TabIndex        =   65
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   21
            Left            =   6840
            TabIndex        =   64
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   20
            Left            =   5520
            TabIndex        =   63
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   19
            Left            =   4200
            TabIndex        =   62
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   18
            Left            =   2880
            TabIndex        =   61
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   17
            Left            =   1560
            TabIndex        =   60
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   16
            Left            =   240
            TabIndex        =   59
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "�w�荆�@�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   10935
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   15
            Left            =   9480
            Style           =   1  '���̨���
            TabIndex        =   40
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   14
            Left            =   8160
            Style           =   1  '���̨���
            TabIndex        =   38
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   13
            Left            =   6840
            Style           =   1  '���̨���
            TabIndex        =   36
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   12
            Left            =   5520
            Style           =   1  '���̨���
            TabIndex        =   34
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   11
            Left            =   4200
            Style           =   1  '���̨���
            TabIndex        =   32
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   10
            Left            =   2880
            Style           =   1  '���̨���
            TabIndex        =   30
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   9
            Left            =   1560
            Style           =   1  '���̨���
            TabIndex        =   28
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   8
            Left            =   240
            Style           =   1  '���̨���
            TabIndex        =   26
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   7
            Left            =   9480
            Style           =   1  '���̨���
            TabIndex        =   24
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   6
            Left            =   8160
            Style           =   1  '���̨���
            TabIndex        =   22
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   5
            Left            =   6840
            Style           =   1  '���̨���
            TabIndex        =   20
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   4
            Left            =   5520
            Style           =   1  '���̨���
            TabIndex        =   18
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   3
            Left            =   4200
            Style           =   1  '���̨���
            TabIndex        =   16
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   2
            Left            =   2880
            Style           =   1  '���̨���
            TabIndex        =   14
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   1560
            Style           =   1  '���̨���
            TabIndex        =   12
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chkSiteiGoki 
            BackColor       =   &H0080FFFF&
            Caption         =   "���I��"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   240
            Style           =   1  '���̨���
            TabIndex        =   6
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   15
            Left            =   9480
            TabIndex        =   41
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   8160
            TabIndex        =   39
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   6840
            TabIndex        =   37
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   5520
            TabIndex        =   35
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   4200
            TabIndex        =   33
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   2880
            TabIndex        =   31
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   1560
            TabIndex        =   29
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   27
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   9480
            TabIndex        =   25
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   8160
            TabIndex        =   23
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   6840
            TabIndex        =   21
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   5520
            TabIndex        =   19
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   4200
            TabIndex        =   17
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2880
            TabIndex        =   15
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   13
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblGokiNo 
            Alignment       =   2  '��������
            BackStyle       =   0  '����
            Caption         =   "Z9���@"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1215
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�ғ��E�����e�f�[�^���W"
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
      TabIndex        =   3
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmSyusyu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmSyusyu.frm
'//  �p�b�P�[�W���F�����ێ�f�[�^���W���
'//
'//  �T�v�F�����ێ�f�[�^���W���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//                 �E�������A�����ێ�f�[�^���W���(frmSyusyu.frm)�𗬗p
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 �ێ瑍�_���C��
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 �L�����Z���{�^��������������ǉ��i�����I���j
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014�N�x�{�� �yEG20_KANSI05_01�z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000     '���[���^�C�}�̃C���^�[�o���l
Public glbFilePath  As String             '�t�@�C���p�X     'V1.12.0.1 ADD
'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
Private mintMaxIndex As Integer
Private Type SHUSHU_STATUS
    intStatus As Integer    '�X�e�[�^�X
    strCaption As String    '�{�^������
    strColor As String      '�{�^���F
    IntValue As Integer     '�������
End Type
Private mudtBtn_Status() As SHUSHU_STATUS

Private mblnZenBtnOuka As Boolean
'EG20 V2.1.0.1 ADD END


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �����ێ�f�[�^���W���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
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
    '���[����M�p�^�C�}���N������
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �����ێ�f�[�^���W���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���~
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
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �����ێ�f�[�^���W���(���[�h��)
'//  �@�\�T�v  : �����������s���B
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
Private Sub Form_Load()

    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    Dim bySyoAssort As Byte             '���O�p������
    Dim intFileNumber As Integer        '�t�@�C���ԍ�
    Dim strFileName As String           '�t�@�C����
    Dim intX() As Integer
    Dim intY() As Integer
    Dim strItmNum As String
    Dim strTemp As String
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intCornerNo As Integer
    Dim intIndex As Integer
    Dim strCorner1 As String
    Dim strCorner2 As String
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '���[����M�p�̃^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    '���@���擾
    Call gsGetGateInfo
    Call gsGetCornerName
    
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    
    '�^�u����ݒu�R�[�i���Ƃ���
    SSTab1.Tab = 0
    SSTab1.Tabs = gintCornerNum

    '���W��ԏ�����
    Erase gintStatus
    
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
            
        End If
    
    Next intCount
    
    '���g�p�̃t�@�C���ԍ����擾���܂��B
    intFileNumber = FreeFile

    '�ݒ���t�@�C������ݒ肷��B
    strFileName = SHUSHU_STATUS_FILE

    '�ݒ���t�@�C�����I�[�v������B
    If strFileName <> "" Then
        Open strFileName For Input As #intFileNumber
    End If

    For intCount = 0 To 1

        '�ݒ���t�@�C�����ɐݒ肳��Ă���t�ݒ�t�@�C����ǂށB
        Input #intFileNumber, strItmNum, strTemp, strTemp, strTemp

        '�ő�R���g���[������ϐ��ɐݒ肷��B
        If intCount = 1 Then
            mintMaxIndex = CInt(strItmNum) - 1
        End If
    Next

    ReDim mudtBtn_Status(mintMaxIndex)

    For intCount = 0 To mintMaxIndex
        '�ݒ���t�@�C�����ɐݒ肳��Ă���t�ݒ�t�@�C����ǂށB
        With mudtBtn_Status(intCount)
            Input #intFileNumber, .intStatus, .strCaption, .strColor, .IntValue
        End With
    Next

    Close #intFileNumber


    intIndex = 0

    '�ݒu�R�[�i�������[�v
    For intCount = 0 To UBound(gblnCornerSet)

        If gblnCornerSet(intCount) = False Then
            SSTab1.TabVisible(intCount) = False
            Frame2(intCount).Visible = False
        End If

        '�ő卆�@�������[�v
        For intCount2 = 0 To 15
            intIndex = (intCount * 16) + intCount2
            lblGokiNo(intIndex).Visible = False
            chkSiteiGoki(intIndex).Visible = False
            chkSiteiGoki(intIndex).Tag = "0"
        Next
        
        For intCount2 = 0 To 15
            intIndex = (intCount * 16) + (gudtSettiCorner(intCount).intGokiNo(intCount2) - 1)
            If gudtSettiCorner(intCount).intGokiNo(intCount2) > 0 Then
                lblGokiNo(intIndex).Caption = gudtSettiCorner(intCount).strDispGoki(intCount2) + "���@"
                'Tag�ɑΉ����鍆�@�ԍ����L�^�i1�`32���@�j
                chkSiteiGoki(intIndex).Tag = CStr(gudtSettiCorner(intCount).intGateNo(intCount2))
                gintStatus(CInt(chkSiteiGoki(intIndex).Tag) - 1) = TAG_STATUS.STS_SENTAKU
                lblGokiNo(intIndex).Visible = True
                chkSiteiGoki(intIndex).Visible = True
            End If
        Next intCount2
        
    Next intCount
    
    Call sSet_GokiStatus(SSTab1.Tab)
    
Exit Sub

'�G���[����
Err_LOG:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If

    '�G���[���O�̏o��
     Call sLogTraceReq(LTYP_ERROR, bySyoAssort, KADO_MENTE_SYUSYU_GAMEN_START, 0)
     
    'EG20 V2.1.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdSyusyu_Click
'//  �@�\����  : �u���W�v�t����������
'//  �@�\�T�v  : �ێ�f�[�^(�ғ��A�����e�A�G���[���O)���W���s���B
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
Private Sub cmdSyusyu_Click()
    Dim iResponse As Integer   'MsgBox�߂�l
    Dim iSendRet As Integer        'V1.7.0.1 ADD
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intIndex As Integer
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next
    
    '�u�ғ��E�����e�f�[�^���W��ʁF���W�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KADO_MENTE_SYUSYU_GAMEN_SYUSYU_BUTTOM, 0)
   
    'V1.7.0.1 ADD START
    iSendRet = CheckAppStart(PROC_KANRI)
    If iSendRet = 0 Then
      '�Ď��Ֆ��N��=�����I��
      '�u�ғ��E�����e�f�[�^���W��ʁF�Ď��Ֆ��N���ُ�v���O�o��
      Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, KADO_MENTE_SYUSYU_GAMEN_NOTAPL_ERROR, 0)
      iResponse = MsgBox("�Ď��Ղ��N�����Ă��܂���B", _
              vbOKOnly, "�m�F")
      Exit Sub
    End If
     'V1.7.0.1 ADD END
    
    '�u�ێ�f�[�^���W�v�|�b�v�A�b�v��\��
    'EG20 V2.1.0.1 DEL START �yMainte_03_01�z
'    iResponse = MsgBox(vbCrLf & _
'              "�S�Ă̎������A�d���n�m�E�ʐM����łȂ��ƁA" & _
'              "���W�Ɏ��s���܂��B" & _
'              vbCrLf & vbCrLf & vbCrLf & _
'              "�m�F���Ă���u�n�j�v�{�^���������ĉ������B" & _
'              vbCrLf & vbCrLf, _
'              vbOKCancel, "�m�F")
    'EG20 V2.1.0.1 DEL END
              
    'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
    iResponse = MsgBox(vbCrLf & _
              "�S�Ẳ��D�@���A�d���n�m�E�ʐM����łȂ��ƁA" & _
              "���W�Ɏ��s���܂��B" & _
              vbCrLf & vbCrLf & vbCrLf & _
              "�m�F���Ă���u�n�j�v�{�^���������ĉ������B" & _
              vbCrLf & vbCrLf, _
              vbOKCancel, "�m�F")
    'EG20 V2.1.0.1 ADD END
              
    If iResponse = vbOK Then
    '�n�j�t�������ꂽ��A
        '�ێ�f�[�^���W���t�H�[�����A���[�_���E�B���h�E�ŕ\������B
        frmSyusyuCyu.Show vbModal
        
        'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
        '�������ʂ�\��
        Call sSet_GokiStatus(SSTab1.Tab)
        'EG20 V2.1.0.1 ADD END
    
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdFDWrite_Click
'//  �@�\����  : �u�}�̏o�́v�t����������
'//  �@�\�T�v  : �ێ�f�[�^(�ғ��A�����e�A�G���[���O)�̏o�͂��s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-10   REVISED BY [TCC] C.Terui
'//                 �t�H���_�I���|�b�v�A�b�v��ǉ�
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 �t�H���_�I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFDWrite_Click()
    Dim iResponse As Integer   'MsgBox�߂�l
    Dim strWriteDir As String  '�I���t�H���_    'V1.12.0.1 ADD
    'EG20 V2.1.0.1 ADD START�yMainte_03_01�z
    Dim intStatusIdx As Integer
    'EG20 V2.1.0.1 ADD END
    
    On Error Resume Next

'V1.12.0.1 ADD START
    '������
    glbFilePath = ""
'V1.12.0.1 ADD END
    
    '�u�ғ��E�����e�f�[�^���W��ʁF�}�̏o�͖t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KADO_MENTE_SYUSYU_GAMEN_OUTPUT_BUTTOM, 0)
    
    '�u�}�̏o�͊m�F�v�|�b�v�A�b�v��\��
     iResponse = MsgBox("�ێ�f�[�^��}�̂ɏo�͂��܂����A��낵���ł����H", _
                         vbOKCancel, _
                         "�ғ��E�����e�f�[�^���W")
    If iResponse = vbOK Then
'V1.12.0.1 DEL START
'    '�n�j�t�������ꂽ��A
'        '�ێ�f�[�^�o�͒��t�H�[�����A���[�_���E�B���h�E�ŕ\������B
'        frmSyusyuOutPut.Show vbModal
'V1.12.0.1 DEL END
'V1.12.0.1 ADD START
        '�t�H���_�I���|�b�v�A�b�v��ʕ\��
        'strWriteDir = pfDirSelection("H:", "�ғ��E�����e�f�[�^�����ݐ�f�B���N�g���I��")   'V1.12.0.1 ADD  'V1.20.0.1 DEL
        strWriteDir = ShowFolders(Me.hwnd, "�t�H���_���w�肵�Ă�������", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD

        '�w��t�H���_�Ȃ�
        If Len(strWriteDir) = 0 Then
             Exit Sub
        End If
        
        glbFilePath = strWriteDir
        
        '�ێ�f�[�^�o�͒��t�H�[�����A���[�_���E�B���h�E�ŕ\������B
        frmSyusyuOutPut.Show vbModal
'V1.12.0.1 ADD END

        'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
        Call sSet_GokiStatus(SSTab1.Tab)
        'EG20 V2.1.0.1 ADD END
    End If
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
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
On Error Resume Next
   '�u�ғ��E�����e�f�[�^���W��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KADO_MENTE_SYUSYU_GAMEN_END, 0)
 
    '����ʂ������B
    Unload Me
End Sub

'EG20 V2.1.0.1 ADD START �yMainte_03_01�z
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : SSTab1_Click
'//  �@�\����  : �^�u�N���b�N����
'//  �@�\�T�v  : �\���^�u��ύX����
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-12   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub SSTab1_Click(PreviousTab As Integer)

    On Error Resume Next
        
    '���ݕ\���^�u�̍X�V
    Call sSet_GokiStatus(SSTab1.Tab)
    
End Sub



'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : chkSiteiGoki_Click
'//  �@�\����  : ���@�t��������
'//  �@�\�T�v  : ���@�t�̃X�e�[�^�X��ύX����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-12   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub chkSiteiGoki_Click(Index As Integer)

    Dim intGokiIdx As Integer
    
    On Error Resume Next
    
    DoEvents
    '�X�V�����Œl���ς�����ꍇ�͔�����
    If mblnZenBtnOuka = True Then
        Exit Sub
    End If
    
    '�Y���R�[�i�ƍ��@�ԍ������߂�
    intGokiIdx = CInt(chkSiteiGoki(Index).Tag) - 1

    '�I�������I��
    If gintStatus(intGokiIdx) = TAG_STATUS.STS_SENTAKU Then
        gintStatus(intGokiIdx) = CStr(TAG_STATUS.STS_MISENTAKU)
    '����ȊO���I��
    Else
        gintStatus(intGokiIdx) = CStr(TAG_STATUS.STS_SENTAKU)
    End If
    
    '���ݕ\���^�u�̍X�V
    Call sSet_GokiStatus(SSTab1.Tab)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdZSentaku_Click
'//  �@�\����  : �S�R�[�i�S���@�I���{�^����������
'//  �@�\�T�v  : �S�R�[�i�S���@�̃{�^����I����Ԃɂ���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-12   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdZSentaku_Click()

    Dim intCount As Integer
    Dim intGokiIndex As Integer
    
    On Error Resume Next
        
    For intCount = 0 To chkSiteiGoki.UBound
        intGokiIndex = CInt(chkSiteiGoki(intCount).Tag) - 1
        If intGokiIndex >= 0 Then
            gintStatus(intGokiIndex) = TAG_STATUS.STS_SENTAKU
        End If
    Next intCount
    
    '���ݕ\���^�u�̍X�V
    Call sSet_GokiStatus(SSTab1.Tab)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdZHisentaku_Click
'//  �@�\����  : �S�R�[�i�S���@���I���{�^����������
'//  �@�\�T�v  : �S�R�[�i�S���@�̃{�^���𖢑I����Ԃɂ���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-12   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdZHisentaku_Click()

    Dim intCount As Integer
    Dim intGokiIndex As Integer
    
    On Error Resume Next

    For intCount = 0 To chkSiteiGoki.UBound
        intGokiIndex = CInt(chkSiteiGoki(intCount).Tag) - 1
        If intGokiIndex >= 0 Then
            gintStatus(intGokiIndex) = TAG_STATUS.STS_MISENTAKU
        End If
    Next intCount
    
    '���ݕ\���^�u�̍X�V
    Call sSet_GokiStatus(SSTab1.Tab)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdHHisentaku_Click
'//  �@�\����  : �\���R�[�i�S���@���I���{�^����������
'//  �@�\�T�v  : �S�R�[�i�S���@�̃{�^���𖢑I����Ԃɂ���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-12   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdHHisentaku_Click()

    Dim intCount As Integer
    Dim intStIndex As Integer
    Dim intEdIndex As Integer
    Dim intGokiIndex As Integer
    
    On Error Resume Next

    intStIndex = SSTab1.Tab * 16
    intEdIndex = intStIndex + 15
    
    For intCount = intStIndex To intEdIndex
        intGokiIndex = CInt(chkSiteiGoki(intCount).Tag) - 1
        If intGokiIndex >= 0 Then
            gintStatus(intGokiIndex) = TAG_STATUS.STS_MISENTAKU
        End If
    Next intCount
    
    '���ݕ\���^�u�̍X�V
    Call sSet_GokiStatus(SSTab1.Tab)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : cmdHSentaku_Click
'//  �@�\����  : �\���R�[�i�S���@�I���{�^����������
'//  �@�\�T�v  : �S�R�[�i�S���@�̃{�^����I����Ԃɂ���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-12   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdHSentaku_Click()

    Dim intCount As Integer
    Dim intStIndex As Integer
    Dim intEdIndex As Integer
    Dim intGokiIndex As Integer
    
    On Error Resume Next

    intStIndex = SSTab1.Tab * 16
    intEdIndex = intStIndex + 15
    
    For intCount = intStIndex To intEdIndex
        intGokiIndex = CInt(chkSiteiGoki(intCount).Tag) - 1
        If intGokiIndex >= 0 Then
            gintStatus(intGokiIndex) = TAG_STATUS.STS_SENTAKU
        End If
    Next intCount
    
    '���ݕ\���^�u�̍X�V
    Call sSet_GokiStatus(SSTab1.Tab)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : sSet_GokiStatus
'//  �@�\����  : ���@�t�ݒ菈��
'//  �@�\�T�v  : �e���@�t�̓��e���ATag�̒l�ɏ]���čX�V����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer   intTab    �X�V�^�u
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-12   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub sSet_GokiStatus(ByVal intTab As Integer)

    Dim intIndex As Integer
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intStIndex As Integer
    Dim intEdIndex As Integer
    Dim intStatusIdx As Integer
    Dim intStatus As Integer
    
    On Error Resume Next

    mblnZenBtnOuka = True
    
    '�Ώۃ^�u�̐擪���@�tIndex���Z�o
    intStIndex = intTab * 16
    intEdIndex = intStIndex + 15
    
    For intCount = intStIndex To intEdIndex
        '�L���ȃ{�^���̂�
        If chkSiteiGoki(intCount).Tag <> "0" Then
            intStatusIdx = CInt(chkSiteiGoki(intCount).Tag) - 1
            intStatus = gintStatus(intStatusIdx)
            'Tag�l�ƈ�v���镶���A�F�A������Ԃɂ���
            For intCount2 = 0 To UBound(mudtBtn_Status)
                If mudtBtn_Status(intCount2).intStatus = intStatus Then
                    chkSiteiGoki(intCount).Caption = mudtBtn_Status(intCount2).strCaption
                    chkSiteiGoki(intCount).BackColor = mudtBtn_Status(intCount2).strColor
                    chkSiteiGoki(intCount).Value = mudtBtn_Status(intCount2).IntValue
                End If
            Next intCount2
        End If
    Next intCount

    mblnZenBtnOuka = False
    
End Sub

'*****************************************************************************
'    (C) Copyright TOSHIBA Corporation 2011 All Right Reserved
'
'     �T�v      : �u���[����M�p�^�C�}�v���^�C���A�b�v�������̃C�x���g�v���V�[�W��
'     ����      : ���[����M�������s���B
'
'     ORIGINAL  :(EG20 V2.1.0.1) 2011-09-15   CODED   BY [TCC] M.Matsumoto
'     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'                 2014�N�x�{�� �yEG20_KANSI05_01�z
'     REVISIONS :(00.00) '00-00-00   REVISED BY [  ]
'*****************************************************************************
Private Sub tmrMail_Timer()

    On Error Resume Next
    
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmSyusyu.Caption, False
        pfFormActive (frmSyusyu.hwnd)           ' EG20 V8.1.0.1�yEG20_KANSI05_01�zADD
    End If

End Sub

'EG20 V2.1.0.1 ADD END
