VERSION 5.00
Begin VB.Form frmToriatukaiKenshuModeSettei 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�����[�g�����e�i���X"
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
      Left            =   6960
      Top             =   7680
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "�s��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   31
      Left            =   10320
      Style           =   1  '���̨���
      TabIndex        =   72
      Top             =   6240
      Value           =   1  '����
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   30
      Left            =   8850
      Style           =   1  '���̨���
      TabIndex        =   71
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "�s��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   29
      Left            =   7410
      Style           =   1  '���̨���
      TabIndex        =   70
      Top             =   6240
      Value           =   1  '����
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   28
      Left            =   5970
      Style           =   1  '���̨���
      TabIndex        =   69
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "�s��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   27
      Left            =   4530
      Style           =   1  '���̨���
      TabIndex        =   68
      Top             =   6240
      Value           =   1  '����
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   26
      Left            =   3090
      Style           =   1  '���̨���
      TabIndex        =   67
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "�s��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   25
      Left            =   1650
      Style           =   1  '���̨���
      TabIndex        =   66
      Top             =   6240
      Value           =   1  '����
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   24
      Left            =   240
      Style           =   1  '���̨���
      TabIndex        =   65
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "�s��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   15
      Left            =   10290
      Style           =   1  '���̨���
      TabIndex        =   56
      Top             =   2760
      Value           =   1  '����
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   14
      Left            =   8850
      Style           =   1  '���̨���
      TabIndex        =   55
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "�s��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   13
      Left            =   7410
      Style           =   1  '���̨���
      TabIndex        =   54
      Top             =   2760
      Value           =   1  '����
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   12
      Left            =   5970
      Style           =   1  '���̨���
      TabIndex        =   53
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "�s��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   4530
      Style           =   1  '���̨���
      TabIndex        =   52
      Top             =   2760
      Value           =   1  '����
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   3090
      Style           =   1  '���̨���
      TabIndex        =   51
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "�s��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   1650
      Style           =   1  '���̨���
      TabIndex        =   50
      Top             =   2760
      Value           =   1  '����
      Width           =   1215
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   240
      Style           =   1  '���̨���
      TabIndex        =   49
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdJIKISelect_All 
      Caption         =   " ���C�戵 �S���@��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   3600
      Style           =   1  '���̨���
      TabIndex        =   40
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdJIKISelect_All 
      Caption         =   " ���C�戵 �S���@�s��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   5280
      Style           =   1  '���̨���
      TabIndex        =   39
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdICSelect_All 
      Caption         =   " �h�b�戵 �S���@��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   240
      Style           =   1  '���̨���
      TabIndex        =   38
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdICSelect_All 
      Caption         =   "�h�b�戵 �S���@�s��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   1920
      Style           =   1  '���̨���
      TabIndex        =   37
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "���C�戵"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   11655
      Begin VB.CheckBox chkMode 
         Caption         =   "�s��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   23
         Left            =   10200
         Style           =   1  '���̨���
         TabIndex        =   64
         Top             =   720
         Value           =   1  '����
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   22
         Left            =   8730
         Style           =   1  '���̨���
         TabIndex        =   63
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "�s��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   21
         Left            =   7290
         Style           =   1  '���̨���
         TabIndex        =   62
         Top             =   720
         Value           =   1  '����
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   20
         Left            =   5850
         Style           =   1  '���̨���
         TabIndex        =   61
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "�s��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   19
         Left            =   4410
         Style           =   1  '���̨���
         TabIndex        =   60
         Top             =   720
         Value           =   1  '����
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   18
         Left            =   2970
         Style           =   1  '���̨���
         TabIndex        =   59
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "�s��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   17
         Left            =   1530
         Style           =   1  '���̨���
         TabIndex        =   58
         Top             =   720
         Value           =   1  '����
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   16
         Left            =   120
         Style           =   1  '���̨���
         TabIndex        =   57
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   31
         Left            =   10200
         TabIndex        =   36
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   8760
         TabIndex        =   35
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   7320
         TabIndex        =   34
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   5880
         TabIndex        =   33
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   4440
         TabIndex        =   32
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   3000
         TabIndex        =   31
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   1560
         TabIndex        =   30
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   10200
         TabIndex        =   28
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   8760
         TabIndex        =   27
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   7320
         TabIndex        =   26
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   5880
         TabIndex        =   25
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   4440
         TabIndex        =   24
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   3000
         TabIndex        =   23
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   1560
         TabIndex        =   22
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�h�b�戵"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   11655
      Begin VB.CheckBox chkMode 
         Caption         =   "�s��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   7
         Left            =   10200
         Style           =   1  '���̨���
         TabIndex        =   48
         Top             =   720
         Value           =   1  '����
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   8760
         Style           =   1  '���̨���
         TabIndex        =   47
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "�s��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   7320
         Style           =   1  '���̨���
         TabIndex        =   46
         Top             =   720
         Value           =   1  '����
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   5880
         Style           =   1  '���̨���
         TabIndex        =   45
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "�s��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   4440
         Style           =   1  '���̨���
         TabIndex        =   44
         Top             =   720
         Value           =   1  '����
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   3000
         Style           =   1  '���̨���
         TabIndex        =   43
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "�s��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   1560
         Style           =   1  '���̨���
         TabIndex        =   42
         Top             =   720
         Value           =   1  '����
         Width           =   1215
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   150
         Style           =   1  '���̨���
         TabIndex        =   41
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   18
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   16
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5880
         TabIndex        =   15
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   7320
         TabIndex        =   14
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   8760
         TabIndex        =   13
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   10200
         TabIndex        =   12
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   1560
         TabIndex        =   10
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   3000
         TabIndex        =   9
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   4440
         TabIndex        =   8
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   5880
         TabIndex        =   7
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   7320
         TabIndex        =   6
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   8760
         TabIndex        =   5
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblGokiBetsuNumber 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "Z9"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   10200
         TabIndex        =   4
         Top             =   1680
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdKakutei 
      Caption         =   "�m��"
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
      Left            =   7440
      TabIndex        =   2
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " �@�@���j���[ �@�@  ��ʂ֖߂�"
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
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "�戵���탂�[�h�ݒ�"
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
      TabIndex        =   1
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmToriatukaiKenshuModeSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmToriatukaiKenshuModeSettei.frm
'//  �p�b�P�[�W���F�戵���탂�[�h�ݒ���
'//
'//  �T�v�F�戵���탂�[�h�ݒ���
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//                 �E�t�F�[�Y�R�Ή��@�V�K�ǉ����
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000             '���[���^�C�}�̃C���^�[�o���l
Private iIC_KenshuMode_Sts(0 To 15) As Integer    'IC�戵���l�擾�G���A
Private iJIKI_KenshuMode_Sts(0 To 15) As Integer  '���C�戵���l�擾�G���A
Private iICGOUKI_SETTEI(0 To 15) As Integer       'IC�ݒ�ύX���@�t���O
Private iJIKIGOUKI_SETTEI(0 To 15) As Integer     '���C�ݒ�ύX���@�t���O
Private Const MAX_GOUKI = 15                      '�ő卆�@�l

Private Const MOVE_JIKI_INDEX = 16                '���C�戵���J�n�C���f�b�N�X�l�܂ł̈ړ�
Private Const SETTEI_ARI = 1                      '�ݒ�L
Private Const SETTEI_NASI = 0                     '�ݒ薳
Private Const HUTEI = -1                          '�l�s��
Private Const HUKA_STS = 1                        '�s�l
Private Const KA_STS = 0                          '�l
Private Const HUKA = "�s��"                       '�\�������F�s��
Private Const KA = "��"                           '�\�������F��
Private Const IC_KENSHU = 0                       'IC�戵����
Private Const JIKI_KENSHU = 1                     '���C�戵����
Dim bBUTTOM_STS As Boolean                        '�t������ԁFTRUE=�������@FALSE=�񉟉�
Dim bUpData_Flag As Boolean                       '�ݒ�X�V�����L���t���O�@TRUE�F�X�V�����L�@FALSE=�X�V������

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : �戵���탂�[�h�ݒ���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    
    On Error Resume Next
 
    '���C����M�p�̃^�C�}���N������B
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : �戵���탂�[�h�ݒ���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�̃^�C�}�N��
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    
    On Error Resume Next
    
    '���C����M�p�̃^�C�}���~�߂�B
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : �戵���탂�[�h�ݒ���(���[�h��)
'//  �@�\�T�v  : �戵���탂�[�h�ݒ��ʂ̏����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    Dim iCnt As Integer     '�J�E���^�[
    
    On Error Resume Next

    '�u�戵���탂�[�h�ݒ��� �\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_GAMEN_START, 0)

    '���C����M�p�̃C���^�o���^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '�e�G���A������
    For iCnt = 0 To MAX_GOUKI
      iIC_KenshuMode_Sts(iCnt) = HUTEI
      iJIKI_KenshuMode_Sts(iCnt) = HUTEI
      iICGOUKI_SETTEI(iCnt) = SETTEI_NASI
      iJIKIGOUKI_SETTEI(iCnt) = SETTEI_NASI
    Next
    
    bUpData_Flag = False
    
    bBUTTOM_STS = True
    
    '��ʕ\������
    pfDispSettei
    
    bBUTTOM_STS = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u���j���[��ʂ֖߂�v�t������
'//  �@�\�T�v  : ����ʂ���������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    
    On Error Resume Next

    '�u�戵���탂�[�h�ݒ��� �����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_GAMEN_END, 0)
   
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdICSelect_All_Click
'//  �@�\����  : IC�戵�S���@�t����������
'//  �@�\�T�v  : IC�戵�S���@�t��������(��/�s��)���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdICSelect_All_Click(Index As Integer)
  
  Dim iCnt As Integer '���@�J�E���^�[
  
  On Error Resume Next
  
  bBUTTOM_STS = True

  If Index = 0 Then
    '�S���@�F�ݒ�
    '�u�戵���탂�[�h�ݒ���:IC�戵�S���@�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_IC_ALLGOUKI_KA_BUTTOM, 0)

    For iCnt = 0 To MAX_GOUKI
        If chkMode(iCnt).Visible = True Then
           iIC_KenshuMode_Sts(iCnt) = KA_STS
           chkMode(iCnt).Caption = KA
           chkMode(iCnt).Value = 0
        End If
    Next
  Else
     '�S���@�F�s�ݒ�
     '�u�戵���탂�[�h�ݒ���:IC�戵�S���@�s�t�����v���O�o��
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_IC_ALLGOUKI_HUKA_BUTTOM, 0)
     For iCnt = 0 To MAX_GOUKI
         If chkMode(iCnt).Visible = True Then
            iIC_KenshuMode_Sts(iCnt) = HUKA_STS
            chkMode(iCnt).Caption = HUKA
            chkMode(iCnt).Value = 1
         End If
    Next
  End If
  
  bBUTTOM_STS = False
  
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : chkMode_Click
'//  �@�\����  : �e���@�ʖt����������
'//  �@�\�T�v  : �e�����@�ʖt��������(��/�s��)���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub chkMode_Click(Index As Integer)
  
  On Error Resume Next
 
  If bBUTTOM_STS = True Then
     'IC/���C�S���@�ꊇ�t�������́A�ȍ~�������ȍ~�̏������s��Ȃ��B
     Exit Sub
  End If
 
  '�u�戵���탂�[�h�ݒ���:���@�ʐݒ�ύX�v���O�o��
  Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_GOUKIBETU_BUTTOM, 0)

  '�e���@�ʖt�l��ύX
  If chkMode(Index).Value = 1 Then
     chkMode(Index).Caption = HUKA
  Else
     chkMode(Index).Caption = KA
  End If
  
  '�戵�G���A�`�F�b�N���s���A�ΏۃG���A�̒l�����ݒl�ɕύX
  If Index < MOVE_JIKI_INDEX Then
    'IC�戵�e���@�F�ݒ�
    If chkMode(Index).Value = 1 Then
       iIC_KenshuMode_Sts(Index) = HUKA_STS
    Else
       iIC_KenshuMode_Sts(Index) = KA_STS
    End If
  Else
     '���C�戵�e���@�F�s�ݒ�
    If chkMode(Index).Value = 1 Then
       iJIKI_KenshuMode_Sts(Index - MOVE_JIKI_INDEX) = HUKA_STS
    Else
       iJIKI_KenshuMode_Sts(Index - MOVE_JIKI_INDEX) = KA_STS
    End If
  End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdJIKISelect_All_Click
'//  �@�\����  : ���C�戵�S���@�t����������
'//  �@�\�T�v  : ���C�戵�S���@�t��������(��/�s��)���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdJIKISelect_All_Click(Index As Integer)
 
  Dim iCnt As Integer '���@�J�E���^�[
  
  On Error Resume Next
  
  bBUTTOM_STS = True

  If Index = 0 Then
     '�S���@�F�ݒ�
     '�u�戵���탂�[�h�ݒ���:���C�戵�S���@�t�����v���O�o��
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_JIKI_ALLGOUKI_KA_BUTTOM, 0)
  
     For iCnt = 0 To MAX_GOUKI
         If chkMode(iCnt + MOVE_JIKI_INDEX).Visible = True Then
            iJIKI_KenshuMode_Sts(iCnt) = KA_STS
            chkMode(iCnt + MOVE_JIKI_INDEX).Caption = KA
            chkMode(iCnt + MOVE_JIKI_INDEX).Value = 0
         End If
    Next
  Else
     '�S���@�F�s�ݒ�
     '�u�戵���탂�[�h�ݒ���:���C�戵�S���@�s�t�����v���O�o��
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KENSHUMODE_SETTEI_JIKI_ALLGOUKI_HUKA_BUTTOM, 0)
   
     For iCnt = 0 To MAX_GOUKI
         If chkMode(iCnt + MOVE_JIKI_INDEX).Visible = True Then
            iJIKI_KenshuMode_Sts(iCnt) = HUKA_STS
            chkMode(iCnt + MOVE_JIKI_INDEX).Caption = HUKA
            chkMode(iCnt + MOVE_JIKI_INDEX).Value = 1
         End If
    Next
  End If
 
 bBUTTOM_STS = False
 
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdKakutei_Click
'//  �@�\����  : ��ʐݒ�l�𔽉f����B
'//  �@�\�T�v  : �����ݒ�G���A�A���͎����ݒ�t�@�C���ɒl�̔��f���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdKakutei_Click()
  
  On Error Resume Next
 
  '�u�戵���탂�[�h�ݒ���:�m��t�����v���O�o��
  Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KANSI_KENSHU_KAKUTEI_BUTTOM, 0)
  
  '��ʂ����b�N����B
  SetEnableFalse
  
  '��ʒl�ݒ蔽�f����
  psGamenSettei_Hanei
  
  '��ʃ��b�N�����B
  SetEnableTrue
  
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : tmrMail_Timer
'//  �@�\����  : �^�C���A�b�v������
'//  �@�\�T�v  : ���[����M�^�C���A�b�v���������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  
    On Error Resume Next
    
    '�ėp���C����M�������s��
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmToriatukaiKenshuModeSettei.Caption, False
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfDispSettei
'//  �@�\����  : �戵���탂�[�h�ݒ���(���[�h��)
'//  �@�\�T�v  : �戵���탂�[�h�ݒ��ʂ̏����������s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfDispSettei()
    Dim iCnt As Integer '���@�J�E���^�[
    Dim iSetti_Gouki As Integer '���@�ݒu/���ݒu��ԃt���O
    
    On Error Resume Next
    
    For iCnt = 0 To MAX_GOUKI
        'INI�t�@�C����荆�@�ݒu/���ݒu�����擾����B
        iSetti_Gouki = pfGetGoukiNo(iCnt + 1)
        If iSetti_Gouki = 1 Then
           '�ݒu�L
          lblGokiBetsuNumber(iCnt).Visible = True
          lblGokiBetsuNumber(iCnt + MOVE_JIKI_INDEX).Visible = True
          lblGokiBetsuNumber(iCnt).Caption = iCnt + 1
          lblGokiBetsuNumber(iCnt + MOVE_JIKI_INDEX).Caption = iCnt + 1
         
          '���@�ʖt�\������
          pfGet_Sts iCnt
        Else
           '���ݒu�FIC/���C�戵�̍��@�ԍ��E���@�ʖt���\���ɂ���B
           lblGokiBetsuNumber(iCnt).Visible = False
           lblGokiBetsuNumber(iCnt + MOVE_JIKI_INDEX).Visible = False
           chkMode(iCnt).Visible = False
           chkMode(iCnt + MOVE_JIKI_INDEX).Visible = False
        End If
    Next
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfGetGoukiNo
'//  �@�\����  : �ݒu���@���擾����B
'//  �@�\�T�v  : GATE.INI���ݒu���@���擾����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Integer            [OUT]0�F���ݒu/�擾�ُ�
'//                                      1�F�ݒu
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetGoukiNo(iGouki As Integer) As Integer

    Dim lngRet As Long          '�֐��̕Ԃ�l
    Dim iGate As Integer        '����INDEX
    Dim j As Integer            '���[�NINDEX
    Dim sKeyName As String
    Dim sGateData As String * RMENTE_GATE_SIZE    '�P�s���t�@�C�����e�擾�p
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim iRet As Integer
   
    On Error Resume Next

    '�������D�@���擾
    sKeyName = "gate" & Format(iGouki, "00")
    iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                   sKeyName, _
                                   DEFAILT, sGateData, Len(sGateData), _
                                   PATH_GATE_FILE)
    If iRet = 0 Then
       '�u�戵���탂�[�h�ݒ��ʁF�������D�@INI�t�@�C���Ǎ��ُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
       '�擾�ُ�
       pfGetGoukiNo = 0
       Exit Function
    End If
             
    If Len(sGateData) <> 0 Then
       '�f�[�^�̎擾
       ReDim sFData(15)
       iFCnt = 1
               
       For iFLoop = 1 To Len(sGateData)
           If Mid(sGateData, iFLoop, 1) <> " " And Mid(sGateData, iFLoop, 1) <> "," Then
              iFLoop2 = iFLoop
              Do
               iFLoop2 = iFLoop2 + 1
               If iFLoop2 > Len(sGateData) Then
                  sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                  iFCnt = iFCnt + 1
                  If iFCnt >= 16 Then
                     Exit For
                  End If
                  
                  iFLoop = iFLoop2
                  Exit Do
                End If
                            
                If Mid(sGateData, iFLoop2, 1) = " " Or Mid(sGateData, iFLoop2, 1) = "," Then
                   sFData(iFCnt) = Mid(sGateData, iFLoop, iFLoop2 - iFLoop)
                   iFCnt = iFCnt + 1
                   If iFCnt >= 16 Then
                      Exit For
                   End If
                      
                   iFLoop = iFLoop2
                   Exit Do
                End If
              Loop
           End If
       Next
    End If
           
    If Trim(sFData(4)) = EGR Then
       '�����^�C�v�FEG-R����/�ݒu�L
       pfGetGoukiNo = 1
       Exit Function
    ElseIf Trim(sFData(4)) = NEG Then
       '�����^�C�v�FNEG����/�ݒu�L
       pfGetGoukiNo = 1
       Exit Function
    Else
       '��L�ȊO�F���ݒu���@����
       pfGetGoukiNo = 0
    End If
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfGet_Sts
'//  �@�\����  : �����ݒ�t�@�C��/�G���A��茻�ݒl���擾����
'//  �@�\�T�v  : �����ݒ�t�@�C��/�G���A��茻�ݒl�̎擾���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-11   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGet_Sts(iGouki As Integer)
    Dim iKansiAplSts As Integer '�Ď��ՃA�v���N�����
    Dim iRet As Integer         '�l�擾�����߂�l
    
    On Error Resume Next
   
    '�Ď��ՃA�v���N���`�F�b�N���s���B
    iKansiAplSts = CheckAppStart(PROC_KANRI)
    If iKansiAplSts <> 0 Then
       '�Ď��ՋN����:�����ݒ�G���A���l�擾
        pfAreaGet_Sts iGouki
       
    Else
       '�Ď��Ֆ��N�����F�����ݒ�t�@�C�����l�擾
       pfFileGet_Sts iGouki
    End If
    
    '�l�擾�`�F�b�N
    '�擾����F�t�\���@�擾�ُ�F���@�ԍ��̂ݕ\��
    'IC/���C�F�擾����
    If iIC_KenshuMode_Sts(iGouki) <> HUTEI And _
       iJIKI_KenshuMode_Sts(iGouki) <> HUTEI Then
       
       'IC�戵�l�擾��������F�t�\��
       chkMode(iGouki).Visible = True
       If iIC_KenshuMode_Sts(iGouki) = HUKA_STS Then
          chkMode(iGouki).Caption = HUKA
          chkMode(iGouki).Value = 1
       Else
          chkMode(iGouki).Caption = KA
          chkMode(iGouki).Value = 0
       End If
       
       '���C�戵�l�擾��������F�t�\��
       chkMode(iGouki + MOVE_JIKI_INDEX).Visible = True
       If iJIKI_KenshuMode_Sts(iGouki) = HUKA_STS Then
         chkMode(iGouki + MOVE_JIKI_INDEX).Caption = HUKA
         chkMode(iGouki + MOVE_JIKI_INDEX).Value = 1
       Else
         chkMode(iGouki + MOVE_JIKI_INDEX).Caption = KA
         chkMode(iGouki + MOVE_JIKI_INDEX).Value = 0
       End If
       
       Exit Function
    
    'IC�擾����/���C�擾�ُ�
    ElseIf iIC_KenshuMode_Sts(iGouki) <> HUTEI And _
           iJIKI_KenshuMode_Sts(iGouki) = HUTEI Then

       'IC�戵�l�擾��������F�t�\��
       chkMode(iGouki).Visible = True
       If iIC_KenshuMode_Sts(iGouki) = HUKA_STS Then
          chkMode(iGouki).Caption = HUKA
          chkMode(iGouki).Value = 1
       Else
          chkMode(iGouki).Caption = KA
          chkMode(iGouki).Value = 0
       End If
       
       '���C�戵���͔�\��
       chkMode(iGouki + MOVE_JIKI_INDEX).Visible = False

       Exit Function
    
    'IC�擾�ُ�/���C�擾����
    ElseIf iIC_KenshuMode_Sts(iGouki) = HUTEI And _
           iJIKI_KenshuMode_Sts(iGouki) <> HUTEI Then
       
       '���C�戵�l�擾��������F�t�\��
       chkMode(iGouki + MOVE_JIKI_INDEX).Visible = True
       If iJIKI_KenshuMode_Sts(iGouki) = HUKA_STS Then
         chkMode(iGouki + MOVE_JIKI_INDEX).Caption = HUKA
         chkMode(iGouki + MOVE_JIKI_INDEX).Value = 1
       Else
         chkMode(iGouki + MOVE_JIKI_INDEX).Caption = KA
         chkMode(iGouki + MOVE_JIKI_INDEX).Value = 0
       End If
       
       'IC�戵���͔�\��
       chkMode(iGouki).Visible = False
       
       Exit Function
    Else
       'IC/���C�擾�����ُ�F�t��\��/���@�ԍ��̂ݕ\��
       chkMode(iGouki).Visible = False
       chkMode(iGouki + MOVE_JIKI_INDEX).Visible = False
    End If
    
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfAreaGet_Sts
'//  �@�\����  : IC/���C�戵�̌��ݒl���擾����(�G���A�Q��)
'//  �@�\�T�v  : IC/���C�戵�̌��ݒl���擾���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iICSts �@�@[OUT]IC�戵���ݒl
'//  ����      : Integer�@iJIKISts �@[OUT]���C�戵���ݒl
'//              Integer�@iGouki  �@ [IN]�����Ώۍ��@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfAreaGet_Sts(iGouki As Integer)
    Dim strMutexName    As String           '�~���[�e�b�N�X��
    Dim lngMuHandle     As Long             '�r�������p�n���h��
    Dim iAreaSts        As Integer          '�G���A�l

    On Error Resume Next
    
    Set Idinf_JikaiSettei = New IdInfProc              '�����ݒ�G���A
    '�����ݒ�G���A���I�[�v������B
    Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
    Idinf_JikaiSettei.IdOpen
    If Idinf_JikaiSettei.Errsts <> 0 Then
       Set Idinf_JikaiSettei = Nothing
       'IC/���C�戵�l�擾�G���A��l�s��ɐݒ�
       iIC_KenshuMode_Sts(iGouki) = HUTEI
       iJIKI_KenshuMode_Sts(iGouki) = HUTEI
       '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Exit Function
    End If
    
    '�����ݒ�G���A���k�n�b�j����B
    Idinf_JikaiSettei.IdLock
    If Idinf_JikaiSettei.Errsts <> 0 Then
       Idinf_JikaiSettei.IdFree
       '�f�[�^�Q�ƈُ펞
       Set Idinf_JikaiSettei = Nothing
       'IC/���C�戵�l�擾�G���A��l�s��ɐݒ�
       iIC_KenshuMode_Sts(iGouki) = HUTEI
       iJIKI_KenshuMode_Sts(iGouki) = HUTEI
       '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Exit Function
     End If
     
     'IC�戵�̓��e��ǂݍ��ށB
     Idinf_JikaiSettei.id = IdGate.IC_TORIATUKAI_KENSHU_STS
     Idinf_JikaiSettei.GetJikai_Sts iGouki
     If Idinf_JikaiSettei.Errsts <> 0 Then
        'IC�戵�擾�ُ�FIC�戵����擾�G���A��l�s��ɐݒ�
        iIC_KenshuMode_Sts(iGouki) = HUTEI
        '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
     Else
        'IC�戵�擾����FIC�戵����擾�G���A�Ɏ擾�l��ݒ�
        iAreaSts = Idinf_JikaiSettei.DataArea(iGouki)
        iIC_KenshuMode_Sts(iGouki) = iAreaSts
     End If
     
     '���C�戵�̓��e��ǂݍ��ށB
     Idinf_JikaiSettei.id = IdGate.JIKI_TORIATUKAI_KENSHU_STS
     Idinf_JikaiSettei.GetJikai_Sts iGouki
     If Idinf_JikaiSettei.Errsts <> 0 Then
        '���C�戵�擾�ُ�F���C�戵����擾�G���A��l�s��ɐݒ�
        iJIKI_KenshuMode_Sts(iGouki) = HUTEI
        '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
     Else
        '���C�戵�擾����F���C�戵����擾�G���A�Ɏ擾�l��ݒ�
        iAreaSts = Idinf_JikaiSettei.DataArea(iGouki)
        iJIKI_KenshuMode_Sts(iGouki) = iAreaSts
     End If
   
     Idinf_JikaiSettei.IdFree
     Set Idinf_JikaiSettei = Nothing
     
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfFileGet_Sts
'//  �@�\����  : IC/���C�戵�̌��ݒl���擾����(�t�@�C���Q��)
'//  �@�\�T�v  : IC/���C�戵�̌��ݒl���擾���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iJikaiSts [OUT]�\���X�e�[�^�X
'//              Integer�@iGouki  �@[IN]�����Ώۍ��@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfFileGet_Sts(iGouki As Integer)
    Dim iAreaSts        As Integer          '�����ݒ�t�@�C����Ԓl
    Dim lSts            As Long             '�֐��߂�l
    Dim udtAreaR255     As GATE_INFO        '�Ǎ��ݗp�G���A�i255�ݒ�p�j
    Dim lngSts          As Long             '�q�b�g�G���AID
    Dim lngLoop1        As Long             '�J�E���^�[
    Dim lngHandle       As Long             '�n���h��
    Dim FileName        As String           '�t�@�C���L���`�F�b�N
    Dim lngRet          As Long             '�߂�l
    Dim bRet            As Boolean          '�ǂݍ��݌��ʖ߂�l
    Dim sSetteiFile     As String           '�t�@�C���p�X�@'V1.4.0.1�@ADD
    
    On Error Resume Next
   
     '�����ݒ�t�@�C�����I�[�v��
    lngHandle = CreateFile(G_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 ADD

    '�t�@�C���I�[�v��������ɍs��ꂽ���H
    If lngHandle = INVALID_HANDLE_VALUE Then
       '�I�[�v���ُ펞�͎Q�ƕs�̂��ߎQ�ƈُ�
       'IC/���C�戵�l�擾�G���A��l�s��ɐݒ�
       iIC_KenshuMode_Sts(iGouki) = HUTEI
       iJIKI_KenshuMode_Sts(iGouki) = HUTEI
       '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Exit Function
    End If
        
    '�����ݒ�t�@�C���ǂݍ���
    For lngLoop1 = 0 To iGouki
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        If bRet = False Then
           '�n���h���̃N���[�Y
           Call CloseHandle(lngHandle)
           'IC/���C�戵�l�擾�G���A��l�s��ɐݒ�
           iIC_KenshuMode_Sts(iGouki) = HUTEI
           iJIKI_KenshuMode_Sts(iGouki) = HUTEI
           '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           Exit Function
        End If
    Next
        
    '�n���h���̃N���[�Y
    Call CloseHandle(lngHandle)
        
    'IC�戵�FID����
    lngSts = SerchId(udtAreaR255, IdGate.IC_TORIATUKAI_KENSHU_STS)
    If lngSts >= 0 Then
       'ID���L�����ꍇ
       iAreaSts = ChgData(udtAreaR255.GateInfo(lngSts))         '�f�[�^�ϊ�
       iIC_KenshuMode_Sts(iGouki) = iAreaSts
    Else
       ' �Y���h�c�����̏ꍇ�Q�ƈُ�
        iIC_KenshuMode_Sts(iGouki) = HUTEI
    End If
    
    '���C�戵�FID����
    lngSts = SerchId(udtAreaR255, IdGate.JIKI_TORIATUKAI_KENSHU_STS)
    If lngSts >= 0 Then
       'ID���L�����ꍇ
       iAreaSts = ChgData(udtAreaR255.GateInfo(lngSts))         '�f�[�^�ϊ�
       iJIKI_KenshuMode_Sts(iGouki) = iAreaSts
    Else
       ' �Y���h�c�����̏ꍇ�Q�ƈُ�
       iJIKI_KenshuMode_Sts(iGouki) = HUTEI
    End If

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : SerchId
'//  �@�\����  : �h�c��������(�戵���탂�[�h�ݒ��ʗp)
'//  �@�\�T�v  : �h�c�������s���B
'//
'//              �^        ����        �Ӗ�
'//  ����      : GATE_INFO udtArea255 [IN]�ϊ����f�[�^
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : Long�@�@�@         �@[OUT]�@0�ȏ�F����B-1�ȉ��F�G���[
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function SerchId(udtArea255 As GATE_INFO, lngId As Long) As Long

    Dim lngIndex As Long                '�����p�C���f�b�N�X
    Dim lngMin As Long                  '�ŏ��C���f�b�N�X
    Dim lngMax As Long                  '�ő�C���f�b�N�X
    Dim lngChkIndex As Long             '�Y���C���f�b�N�X
    Dim lngWorkId   As Long             '�W���h�c

    On Error Resume Next
    
    '������
    lngMin = 0
    lngMax = ID_GATE_MAX - 1
    lngChkIndex = -1

    '�����J�n
    Do While lngMin <= lngMax
        lngIndex = lngMin
        lngWorkId = udtArea255.GateInfo(lngIndex).intId             '�h�c���o��
        If lngId = lngWorkId Then                                  '�����H
            lngChkIndex = lngIndex                                  '�f�[�^���o����A�����I��
            Exit Do
        Else
            If (lngWorkId = 0) Or (lngId < lngId) Then         '�f�[�^���\����������
                lngMin = lngMin + 1
            Else
                lngMin = lngMin + 1
            End If
        End If
    Loop
            
    SerchId = lngChkIndex

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : ChgData
'//  �@�\����  : �f�[�^�ϊ���������(�戵���탂�[�h�ݒ��ʗp)
'//  �@�\�T�v  : �f�[�^�ϊ������������s���B
'//
'//              �^        ����        �Ӗ�
'//  ����      : ID_FMT �@DataArea �@[IN]�ϊ����f�[�^
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : String�@�@�@        [OUT]�@vbNullstring�ȊO�F����BvbNullString    �F�G���[
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function ChgData(DataArea As ID_FMT) As String

    Dim lngloop As Long
    Dim lngWork As Long
    Dim lngErrsts As Long

    On Error GoTo ChgDataErr
    
    lngErrsts = IdInfErr.OK
    
    Select Case DataArea.intType
    Case ID_TYPE.Flag   '���
        If (DataArea.bytDATA(0) <> 255) Then
            ChgData = str$(DataArea.bytDATA(0))
            
        Else
            ChgData = "-1"                      '�l���s��Ȃ�[�P�Z�b�g
            
        End If
            
    Case ID_TYPE.Count  '��
        lngWork = 0                              '������
        For lngloop = 3 To 0 Step -1
            lngWork = lngWork * 256 + DataArea.bytDATA(lngloop)
        Next lngloop
                        
        ChgData = str$(lngWork)
    
    Case ID_TYPE.Date_Type, ID_TYPE.time_type '���t�A����
        ChgData = StrConv(DataArea.bytDATA, vbUnicode)
        
    Case Else
        ChgData = vbNullString
        lngErrsts = IdInfErr.ID_TYPE_MISS
        Exit Function

    End Select
    
    Exit Function
    
ChgDataErr:
        ChgData = vbNullString
        lngErrsts = IdInfErr.PROC_ERR
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : SetEnableFalse
'//  �@�\����  : ��ʃ��b�N��������
'//  �@�\�T�v  : ��ʂ̃��b�N����������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()
    Dim iCnt As Integer
    
    On Error Resume Next

    'IC�戵�S���@��/�s�t�FFalse(���b�N)����B
    cmdICSelect_All(0).Enabled = False
    cmdICSelect_All(1).Enabled = False
    
    '���C�戵�S���@��/�s�t�FFalse(���b�N)����B
    cmdJIKISelect_All(0).Enabled = False
    cmdJIKISelect_All(1).Enabled = False
    
    '�m��t�FFalse(���b�N)����B
    cmdKakutei.Enabled = False
    
    '���j���[��ʂ֖߂�t�FFalse(���b�N)����B
    cmdReturn.Enabled = False
    
    For iCnt = 0 To MAX_GOUKI
        'IC�戵�G���A�FFalse(���b�N)����B
        chkMode(iCnt).Enabled = False
        '���C�戵�G���A�FFalse(���b�N)����B
        chkMode(iCnt + MOVE_JIKI_INDEX).Enabled = False
    Next

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
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()
    Dim iCnt As Integer
    
    On Error Resume Next

    'IC�戵�S���@��/�s�t�FTrue(���b�N����)����B
    cmdICSelect_All(0).Enabled = True
    cmdICSelect_All(1).Enabled = True
    
    '���C�戵�S���@��/�s�t�FTrue(���b�N����)����B
    cmdJIKISelect_All(0).Enabled = True
    cmdJIKISelect_All(1).Enabled = True
    
    '�m��t�FTrue(���b�N����)����B
    cmdKakutei.Enabled = True
    
    '���j���[��ʂ֖߂�t�FTrue(���b�N����)����B
    cmdReturn.Enabled = True
    
    For iCnt = 0 To MAX_GOUKI
        'IC�戵�G���A�FTrue(���b�N����)����B
        chkMode(iCnt).Enabled = True
        '���C�戵�G���A�FTrue(���b�N����)����B
        chkMode(iCnt + MOVE_JIKI_INDEX).Enabled = True
    Next

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psGamenSettei_Hanei
'//  �@�\����  : ��ʒl���f����
'//  �@�\�T�v  : ��ʒl���G���A���̓t�@�C���ɔ��f����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l �F
'///////////////////////////////////////////////////////////////////
Public Sub psGamenSettei_Hanei()
    Dim iKansiAplSts As Integer '�Ď��ՃA�v���N�����
    Dim iCnt As Integer         '�J�E���^�[
    Dim bRet As Boolean         '���f�����߂�l
    Dim iRet As Integer         '���b�Z�[�W�{�b�N�X�߂�l
    Dim bJikiRet As Boolean     '���C���f�����߂�l
    Dim bICRet As Boolean       '�h�b���f�����߂�l
    
    On Error Resume Next
   
    '�Ď��ՃA�v���N���`�F�b�N���s���B
    iKansiAplSts = CheckAppStart(PROC_KANRI)
    If iKansiAplSts <> 0 Then
       
       '�Ď��ՋN����:�����ݒ�G���A�X�V�������s��
        For iCnt = 0 To MAX_GOUKI
            'IC�戵����l�擾�G���A�`�F�b�N�F�l�s��ȊO
            If iIC_KenshuMode_Sts(iCnt) <> HUTEI Then
               bRet = pfAreaSet_Sts(iCnt, IC_KENSHU)
               bUpData_Flag = True
            End If
            If iICGOUKI_SETTEI(iCnt) = SETTEI_ARI Then
               '�ݒ�ύX�L��
               bICRet = True
            End If
                                 
            '���C�戵����l�擾�G���A�`�F�b�N�F�l�s��ȊO
            If iJIKI_KenshuMode_Sts(iCnt) <> HUTEI Then
               bRet = pfAreaSet_Sts(iCnt, JIKI_KENSHU)
               bUpData_Flag = True
            End If
            If iJIKIGOUKI_SETTEI(iCnt) = SETTEI_ARI Then
               '�ݒ�ύX�L��
               bJikiRet = True
            End If
        Next
        
        If bICRet = False And bJikiRet = False And bUpData_Flag = True Then
           '�X�V�����ُ펞�F��������(�ُ�I��)�|�b�v�A�b�v��ʕ\��
           iRet = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "���f��������")

           '�ݒ�ύX���@�t���O�F�ύX�����ɐݒ�
           For iCnt = 0 To MAX_GOUKI
               iICGOUKI_SETTEI(iCnt) = SETTEI_NASI
               iJIKIGOUKI_SETTEI(iCnt) = SETTEI_NASI
           Next
           '�G���A�X�V�����I��
           bUpData_Flag = False
           Exit Sub
        End If
        
        '�����ݒ�w�����ă}�ɑ��M����B
        bRet = pfSendMail
        If bRet = False Then
           '���M�ُ�F�u�����ݒ�w���F���M�ُ�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_MAIL, KENSHUMODE_SETTEI_JIKAIMAIL_ERROR, 0)
        Else
           '���M����F�u�����ݒ�w���F���M����v���O�o��
           Call sLogTraceReq(LTYP_NORMAL, L3AN_MAIL, KENSHUMODE_SETTEI_JIKAIMAIL_OK, 0)
        End If
                
    Else
       '�Ď��Ֆ��N�����F�����ݒ�t�@�C�����l�擾
       For iCnt = 0 To MAX_GOUKI
           'IC�戵����l�擾�G���A�`�F�b�N�F�l�s��ȊO
           If iIC_KenshuMode_Sts(iCnt) <> HUTEI Then
              bRet = pfFileSet_Sts(iCnt, IC_KENSHU)
              bUpData_Flag = True
           End If
           If iICGOUKI_SETTEI(iCnt) = SETTEI_ARI Then
               '�ݒ�ύX�L��
               bICRet = True
           End If
           '���C�戵����l�擾�G���A�`�F�b�N�F�l�s��ȊO
           If iJIKI_KenshuMode_Sts(iCnt) <> HUTEI Then
              bRet = pfFileSet_Sts(iCnt, JIKI_KENSHU)
              bUpData_Flag = True
           End If
           If iJIKIGOUKI_SETTEI(iCnt) = SETTEI_ARI Then
               '�ݒ�ύX�L��
               bJikiRet = True
           End If
        Next
        
        If bICRet = False And bJikiRet = False And bUpData_Flag = True Then
           '�X�V�����ُ펞�F��������(�ُ�I��)�|�b�v�A�b�v��ʕ\��
           iRet = MsgBox("�ُ�I�����܂����B", vbOKOnly + vbCritical, "���f��������")
           '�ݒ�ύX���@�t���O�F�ύX�����ɐݒ�
            For iCnt = 0 To MAX_GOUKI
                iICGOUKI_SETTEI(iCnt) = SETTEI_NASI
                iJIKIGOUKI_SETTEI(iCnt) = SETTEI_NASI
            Next
            '�G���A�X�V�����I��
            bUpData_Flag = False
            Exit Sub
        End If
    End If
      
    '�ݒ�ύX���@�t���O�F�ύX�����ɐݒ�
    For iCnt = 0 To MAX_GOUKI
        iICGOUKI_SETTEI(iCnt) = SETTEI_NASI
        iJIKIGOUKI_SETTEI(iCnt) = SETTEI_NASI
    Next
    '�G���A�X�V�����I��
    bUpData_Flag = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfAreaSet_Sts
'//  �@�\����  : �����ݒ�G���A��IC/���C�戵�̌��ݒl��ݒ菈��(�G���A�Q��)
'//  �@�\�T�v  : IC/���C�戵�̌��ݒl�̐ݒ���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iICSts �@�@[OUT]IC�戵���ݒl
'//  ����      : Integer�@iJIKISts �@[OUT]���C�戵���ݒl
'//              Integer�@iGouki  �@ [IN]�����Ώۍ��@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfAreaSet_Sts(iGouki As Integer, iUpData_ID As Integer) As Boolean
    Dim strMutexName    As String           '�~���[�e�b�N�X��
    Dim lngMuHandle     As Long             '�r�������p�n���h��
    Dim iAreaSts        As Integer          '�G���A�l

    On Error Resume Next
    
    Set Idinf_JikaiSettei = New IdInfProc              '�����ݒ�G���A
    '�����ݒ�G���A���I�[�v������B
    Idinf_JikaiSettei.ProcMode = DATA_ID.Data_Id_JikaiSettei
    Idinf_JikaiSettei.IdOpen
    If Idinf_JikaiSettei.Errsts <> 0 Then
      '�f�[�^�Q�ƈُ펞
      Set Idinf_JikaiSettei = Nothing
      '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
      pfAreaSet_Sts = False
      Exit Function
    End If
    
    '�����ݒ�G���A���k�n�b�j����B
    Idinf_JikaiSettei.IdLock
    If Idinf_JikaiSettei.Errsts <> 0 Then
       Idinf_JikaiSettei.IdFree
       '�f�[�^�Q�ƈُ펞
       Set Idinf_JikaiSettei = Nothing
       '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       pfAreaSet_Sts = False
       Exit Function
     End If
     
     If iUpData_ID = IC_KENSHU Then
        'IC�戵�̓��e��ǂݍ��ށB
        Idinf_JikaiSettei.id = IdGate.IC_TORIATUKAI_KENSHU_STS
        Idinf_JikaiSettei.SetICM_Sts iGouki, iIC_KenshuMode_Sts(iGouki)
        If Idinf_JikaiSettei.Errsts <> 0 Then
           '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           pfAreaSet_Sts = False
        Else
           iICGOUKI_SETTEI(iGouki) = SETTEI_ARI
        End If
     Else
       '���C�戵�̓��e��ǂݍ��ށB
       Idinf_JikaiSettei.id = IdGate.JIKI_TORIATUKAI_KENSHU_STS
       Idinf_JikaiSettei.SetICM_Sts iGouki, iJIKI_KenshuMode_Sts(iGouki)
       If Idinf_JikaiSettei.Errsts <> 0 Then
          '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
          Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
          pfAreaSet_Sts = False
       Else
          iJIKIGOUKI_SETTEI(iGouki) = SETTEI_ARI
       End If
     End If
     
     Idinf_JikaiSettei.IdFree
     Set Idinf_JikaiSettei = Nothing
     pfAreaSet_Sts = True
     
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfFileSet_Sts
'//  �@�\����  : IC/���C�戵�̌��ݒl�ݒ菈��(�t�@�C���Q��)
'//  �@�\�T�v  : IC/���C�戵�̌��ݒl�������ݒ�t�@�C���ɐݒ肷��B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iJikaiSts [OUT]�\���X�e�[�^�X
'//              Integer�@iGouki  �@[IN]�����Ώۍ��@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfFileSet_Sts(iGouki As Integer, iUpData_ID As Integer) As Boolean
    Dim iAreaSts        As Integer          '�����ݒ�t�@�C����Ԓl
    Dim lSts            As Long             '�֐��߂�l
    Dim udtAreaR255     As GATE_INFO        '�Ǎ��ݗp�G���A�i255�ݒ�p�j
    Dim lngSts          As Long             '�q�b�g�G���AID
    Dim lngLoop1        As Long             '�J�E���^�[
    Dim lngHandle       As Long             '�n���h��
    Dim FileName        As String           '�t�@�C���L���`�F�b�N
    Dim lngRet          As Long             '�߂�l
    Dim bRet            As Boolean          '�ǂݍ��݌��ʖ߂�l
    Dim sSetteiFile     As String           '�t�@�C���p�X
    Dim udtAreaR255Work As GATE_INFO        '�Ǎ��ݗp�G���A�i�|�C���^�ړ��p�j
    Dim iUpData_Sts     As Integer          '�ݒ�l
   
    On Error Resume Next
     
    '�����ݒ�t�@�C�����I�[�v��
    lngHandle = CreateFile(G_SETTEI_FILE, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0) 'V1.4.0.1 ADD

    '�t�@�C���I�[�v��������ɍs��ꂽ���H
    If lngHandle = INVALID_HANDLE_VALUE Then
       '�I�[�v���ُ펞�͎Q�ƕs�̂��ߎQ�ƈُ�
       '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       pfFileSet_Sts = False
       Exit Function
    End If
        
    '�����ݒ�t�@�C���ǂݍ���
    For lngLoop1 = 0 To iGouki
        bRet = ReadFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
        If bRet = False Then
           '�n���h���̃N���[�Y
           Call CloseHandle(lngHandle)
           '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
           Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
           pfFileSet_Sts = False
           Exit Function
        End If
    Next
        
    '�n���h���̃N���[�Y
    Call CloseHandle(lngHandle)
        
    'IC�戵�FID����
    If iUpData_ID = IC_KENSHU Then
       lngSts = SerchId(udtAreaR255, IdGate.IC_TORIATUKAI_KENSHU_STS)
       If lngSts >= 0 Then
       'ID���L�����ꍇ
          iUpData_Sts = iIC_KenshuMode_Sts(iGouki)
          udtAreaR255.GateInfo(lngSts).bytDATA(0) = iUpData_Sts
       Else
          ' �Y���h�c�����̏ꍇ�F�������Ȃ�
          pfFileSet_Sts = False
       End If
    Else
      '���C�戵�FID����
      lngSts = SerchId(udtAreaR255, IdGate.JIKI_TORIATUKAI_KENSHU_STS)
      If lngSts >= 0 Then
         'ID���L�����ꍇ
         iUpData_Sts = iJIKI_KenshuMode_Sts(iGouki)
         udtAreaR255.GateInfo(lngSts).bytDATA(0) = iUpData_Sts
      Else
         ' �Y���h�c�����̏ꍇ�F�������Ȃ�
         pfFileSet_Sts = False
      End If
    End If

    '�����ݒ�t�@�C�����I�[�v��
    lngHandle = CreateFile(G_SETTEI_FILE, _
                           GENERIC_READ + GENERIC_WRITE, _
                           FILE_SHARE_READ + FILE_SHARE_WRITE, _
                           0, _
                           OPEN_EXISTING, _
                           FILE_ATTRIBUTE_NORMAL, _
                           0)

    '�t�@�C���I�[�v��������ɍs��ꂽ���H
    If lngHandle = INVALID_HANDLE_VALUE Then
       '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       pfFileSet_Sts = False
       Exit Function
    End If
     
    '�t�@�C���|�C���^�ړ��̂��߂̓ǂݍ���
     For lngLoop1 = 0 To iGouki - 1
         bRet = ReadFile(lngHandle, udtAreaR255Work, LenB(udtAreaR255Work), lngRet, 0)
         If bRet = False Then
            '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
            Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
            Call CloseHandle(lngHandle)
            pfFileSet_Sts = False
            Exit Function
         End If
     Next
    
    '�����ݒ�t�@�C���ɏ�������
    bRet = WriteFile(lngHandle, udtAreaR255, LenB(udtAreaR255), lngRet, 0)
    If bRet = False Then
       '�u�戵���탂�[�h�ݒ��ʁF�G���A�E�t�@�C���Q�ƈُ�v���O�o��
       Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, KANSI_KENSHU_AREA_FILE_NOTACCESS_ERROR, 0)
       Call CloseHandle(lngHandle)
       pfFileSet_Sts = False
       Exit Function
    End If
    
    '�n���h���̃N���[�Y
     Call CloseHandle(lngHandle)
    
    '�ݒ�ύX���@�t���O�ݒ�L��
    If iUpData_ID = IC_KENSHU Then
       iICGOUKI_SETTEI(iGouki) = SETTEI_ARI
    Else
       iJIKIGOUKI_SETTEI(iGouki) = SETTEI_ARI
    End If

    pfFileSet_Sts = True
     
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfSendMail
'//  �@�\����  : �u�����ݒ�w���v���M
'//  �@�\�T�v  : IC/���C�戵�̕ύX��ʒm����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@iJikaiSts [OUT]�\���X�e�[�^�X
'//              Integer�@iGouki  �@[IN]�����Ώۍ��@�ԍ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.6.0.1) 2009-06-12   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfSendMail() As Boolean
    
    Dim udtMail     As MAIL_GATE_SET_ORD    '�����ݒ�w�����[�����M�G���A
    Dim lngRet      As Long                 '�֐��߂�l
    Dim intCnt      As Integer              '�J�E���^

    On Error Resume Next

    '���ʃw�b�_�ҏW
    udtMail.mlHeader.dwId = ML_ID_GATE_SET_ORD
    udtMail.mlHeader.dwSize = MlSize.GATE_SET_ORD
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    
    '�G���A��ʂ�ݒ�
    udtMail.dwCmnFile = G_SETTEI_FILE_NO
    
    '�ݒ���
    For intCnt = 0 To MAX_GATE_NO - 1
        If iICGOUKI_SETTEI(intCnt) = SETTEI_ARI Or iJIKIGOUKI_SETTEI(intCnt) = SETTEI_ARI Then
            udtMail.dwGateSet(intCnt) = 1
        Else
            udtMail.dwGateSet(intCnt) = 0
        End If
    Next intCnt

    '���[�����M
    pfSendMail = DssSendMail(MAIL_SLOT_KANMA, MlSize.GATE_SET_ORD, udtMail.mlHeader)

End Function

