VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUtility 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "���[�e�B���e�B�N��"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4680
      Top             =   8160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�O���"
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
      Index           =   0
      Left            =   1200
      TabIndex        =   42
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����"
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
      Index           =   1
      Left            =   5520
      TabIndex        =   41
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Index           =   9
      Left            =   8650
      TabIndex        =   40
      Top             =   6960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Index           =   8
      Left            =   8650
      TabIndex        =   39
      Top             =   6240
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Index           =   7
      Left            =   8650
      TabIndex        =   38
      Top             =   5520
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Index           =   6
      Left            =   8650
      TabIndex        =   37
      Top             =   4800
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Index           =   5
      Left            =   8650
      TabIndex        =   36
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Index           =   4
      Left            =   8650
      TabIndex        =   35
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Index           =   3
      Left            =   8650
      TabIndex        =   34
      Top             =   2640
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Index           =   2
      Left            =   8650
      TabIndex        =   33
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Index           =   1
      Left            =   8650
      TabIndex        =   32
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "������������������������"
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
      Index           =   0
      Left            =   8650
      TabIndex        =   31
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtExeName 
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
      Left            =   1440
      TabIndex        =   30
      Top             =   600
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "�N ��"
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
      Index           =   0
      Left            =   7680
      TabIndex        =   29
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "�N ��"
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
      Index           =   1
      Left            =   7680
      TabIndex        =   28
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
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
      Left            =   1440
      TabIndex        =   27
      Top             =   1320
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "�N ��"
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
      Index           =   2
      Left            =   7680
      TabIndex        =   26
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
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
      Left            =   1440
      TabIndex        =   25
      Top             =   2040
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "�N ��"
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
      Index           =   3
      Left            =   7680
      TabIndex        =   24
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
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
      Left            =   1440
      TabIndex        =   23
      Top             =   2760
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "�N ��"
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
      Index           =   4
      Left            =   7680
      TabIndex        =   22
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
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
      Left            =   1440
      TabIndex        =   21
      Top             =   3480
      Width           =   6135
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "�ݒ�ύX"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "�ݒ�ύX"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "�ݒ�ύX"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "�ݒ�ύX"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "�ݒ�ύX"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "�ݒ�ύX"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   120
      TabIndex        =   15
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "�ݒ�ύX"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   120
      TabIndex        =   14
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "�ݒ�ύX"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "�ݒ�ύX"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "�ݒ�ύX"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtExeName 
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
      Left            =   1440
      TabIndex        =   10
      Top             =   7080
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "�N ��"
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
      Index           =   9
      Left            =   7680
      TabIndex        =   9
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
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
      Left            =   1440
      TabIndex        =   8
      Top             =   6360
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "�N ��"
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
      Index           =   8
      Left            =   7680
      TabIndex        =   7
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
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
      Left            =   1440
      TabIndex        =   6
      Top             =   5640
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "�N ��"
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
      Index           =   7
      Left            =   7680
      TabIndex        =   5
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
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
      Left            =   1440
      TabIndex        =   4
      Top             =   4920
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "�N ��"
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
      Index           =   6
      Left            =   7680
      TabIndex        =   3
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "�N ��"
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
      Index           =   5
      Left            =   7680
      TabIndex        =   2
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtExeName 
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
      Left            =   1440
      TabIndex        =   1
      Top             =   4200
      Width           =   6135
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " �����e�i���X   ��ʂ֖߂�"
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
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "���[�e�B���e�B�N��"
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
      TabIndex        =   43
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmUtility.frm
'//  �p�b�P�[�W���F���[�e�B���e�B�N��(���������e�i���X)���
'//
'//  �T�v�F���[�e�B���e�B�N��(���������e�i���X)���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-16  REVISED BY [TCC] S.Yoshimori
'//                 �t�@�C���I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 �}�̎�O�s��C��
'//     REVISIONS :(2.8.0.1) 2011-02-07   REVISED BY [TCC] S.Terao
'//                 �z��Q�ƕs��C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const iHoshuAplMax = 19            '�o�^�ő匏��
Private sChangeExePass(0 To 31) As String  '�ύX�\�Œ�N���t�ɑΉ���������̧���߽���i��޴ر���܂ށj
Private sFixedExePass(0 To 31) As String   '�Œ�N���t�ɑΉ���������̧���߽���i��޴ر���܂ށj
Private sFixedExeName(0 To 31) As String   '�Œ�N���t�ɑΉ������t���́i��޴ر���܂ށj
Private iGamenSts As Integer               '���ݕ\����ʐ�
Private iHyoujiCnt As Integer              '�\���J�E���^�[
Private iKoteiHyouji_Flag As Integer       '�Œ�o�^��10���ȏ�t���O
Private iChangeHyouji_Flag As Integer      '�ύX�o�^��10���ȏ�t���O
Private iContinuFlag As Integer            '�u����ʁv�u�O��ʁv�t�\���L���t���O

'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000     '���[���^�C�}�̃C���^�[�o���l

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : ���[�e�B���e�B�N��(���������e�i���X)���(�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���N��
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
On Error Resume Next
    '���[����M�p�^�C�}���N������
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Deactivate
'//  �@�\����  : ���[�e�B���e�B�N��(���������e�i���X)���(�f�B�A�N�e�B�u��)
'//  �@�\�T�v  : ���[����M�p�^�C�}���~
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
    '���[����M�p�^�C�}���~�߂�
    tmrMail.Enabled = False
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Load
'//  �@�\����  : ���[�e�B���e�B�N��(���������e�i���X)���(���[�h��)
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
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim i As Integer    '�J�E���^�[
   
   On Error Resume Next
 
   '�uհè�è��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_START, 0)

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    '������
    iHyoujiCnt = 0  '�\���J�E���^�[
    iGamenSts = 0 '���ݕ\����ʐ�
    Command1(0).Visible = False '�u�O��ʁv�t��\���B
    Command1(1).Visible = False '�u����ʁv�t��\���B
    iKoteiHyouji_Flag = 0
    iChangeHyouji_Flag = 0
    
    'V1.3.0.1 ADD START
    '���[����M�p�̃^�C�}�l��ݒ肷��B
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    '1.3.0.1 ADD END
    
    For i = 0 To 31
        '�\�����G���A������
        sFixedExeName(i) = ""
    Next
    For i = 0 To 31
        '�c�[���p�X�G���A������
        sFixedExePass(i) = ""
    Next
    For i = 0 To 31
        '�ύX�\�G���A������
        sChangeExePass(i) = ""
    Next
    
    
    '�ύX�\�Œ�A�v���\������
    sFixedKoteiExeDisplay
    
    '�Œ�A�v���̏����擾���A�N���p�t��\������B
    sFixedExeDisplay
    
    '�o�^����10���ȏ�`�F�b�N���s���B
    If iKoteiHyouji_Flag = 1 Then
     '�Œ�A�v����10���ȏ゠��ꍇ
      Command1(0).Visible = True
      Command1(1).Visible = True
     '�u����ʁv�u�O��ʁv�t�\���t���O��ON�ɂ���B
      iContinuFlag = True
    End If
    
    If iChangeHyouji_Flag = 1 Then
     '�ύX�\�Œ�A�v����10���ȏ゠��ꍇ
      Command1(0).Visible = True
      Command1(1).Visible = True
     '�u����ʁv�u�O��ʁv�t�\���t���O��ON�ɂ���B
      iContinuFlag = True
    End If
   
   '���ݕ\����ʐ���1��ʖڂɐݒ肷��B
    iGamenSts = 1

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdChange_Click
'//  �@�\����  : �u�ݒ�ύX�v�t����������
'//  �@�\�T�v  : �A�v���̐ݒ��ύX���邽�߂́A�A�v���I����ʂ�\�����A
'//              �ݒ���X�V����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@�@[IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-16  REVISED BY [TCC] S.Yoshimori
'//                 �t�@�C���I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 �}�̎�O�s��C��
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdChange_Click(Index As Integer)
    Dim iResponse As Integer 'MsgBox�{�^���R�[�h
    Dim sFileName As String  '�I�����ꂽ���s�t�@�C����
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g  'V1.20.0.1 ADD
    
    On Error Resume Next
    
    '�uհè�è��ʁF�ݒ�ύX�t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_CHANGE_SETTEI_BUTTOM, 0)

    
    '��ʐݒ�C���f�b�N�X��0�`9�Ȃ̂ŁA�t�C���f�b�N�X�l���Z�o���A
    '�N���A�v���̃p�X�ŋN������B
    '�N���A�v���p�X�C���f�b�N�X=(���݉�ʐ�-1���)�~1��ʍő�t���{�����C���f�b�N�X(0�`9)
    '��F2��ʖڂ̉����t�C���f�b�N�X3���������ꂽ�ꍇ�A�N���A�v���p�X�C���f�b�N�X��13
    '13=(2-1)��10�{3
    Index = (iGamenSts - 1) * 10 + Index
    
    '�A�v���ݒ�ύX�̂��߂̃t�@�C���I����ʂ��o�͂���B
    'sFileName = pfFileSelection("D:", "*.exe;*.com;*.bat;*.cmd", _
                                        "���s�t�@�C���I��")    'V1.20.0.1 DEL
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
    Set objFso = Nothing
    '�g���q��ݒ�
    CommonDialog1.Filter = _
        "�v���O�����t�@�C���i*.exe;*.com;*.bat;*.cmd�j|*.exe;*.com;*.bat;*.cmd|"
    '�t�@�C���I����ʂ��J��
    CommonDialog1.ShowOpen
    '�I�������t�@�C�������擾
    sFileName = CommonDialog1.FileName
    'V1.20.0.1 ADD END
                                   
     Call ChDrive("D")  'V2.5.0.1 ADD
                              
    '�t�@�C���I����ʂł̃A�v���̑I��L�����`�F�b�N����B
    If sFileName <> "" Then
         If iGamenSts = 2 Then
             '���ݕ\����ʁF2��ʖځB
             '�\�����̑������f�̂��߂ɁA�C���f�b�N�X�ԍ���1��ʕ�(-10)���Đݒ肷��B
             txtExeName(Index - 10) = sFileName
         Else
             '���ݕ\����ʁF1��ʖځB
             '�\�����̑������f�̂��߂ɐݒ肷��B
             txtExeName(Index) = sFileName
        End If
         '�ύX�\�Œ�N���t�ɑΉ���������̧���߽�ɂ��ύX���ꂽ�߽����ݒ肷��B
         sChangeExePass(Index) = sFileName
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : txtExeName_DblClick
'//  �@�\����  : ̧���߽�\�����_�u���N���b�N������
'//  �@�\�T�v  : �폜�m�F���b�Z�[�W��\�����A�폜���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@�@[IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub txtExeName_DblClick(Index As Integer)
    Dim iResponse As Integer      'MsgBox�{�^���R�[�h
    Dim iSetupAplIndex As Integer '�N���A�v���C���f�b�N�X
    
    On Error Resume Next
   
   '�uհè�è��ʁF�ݒ�\��������ٸد��v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_DOUBLECLICK_SETTEI, 0)

    '��ʐݒ�C���f�b�N�X��0�`9�Ȃ̂ŁA�t�C���f�b�N�X�l���Z�o���A
    '�N���A�v���̃p�X�ŋN������B
    '�N���A�v���C���f�b�N�X=(���݉�ʐ�-1���)�~1��ʍő�t���{�����C���f�b�N�X(0�`9)
    '��F2��ʖڂ̉����t�C���f�b�N�X3���������ꂽ�ꍇ�A�N���A�v���p�X�C���f�b�N�X��13
    '13=(2-1)��10�{3
   iSetupAplIndex = (iGamenSts - 1) * 10 + Index

   '�ύX�\�Œ�N���t�t�@�C���߽�i�[�G���A�̒�`�L���`�F�b�N���s���B
   If sChangeExePass(iSetupAplIndex) <> "" Then
        '�u�o�^���O�v�|�b�v�A�b�v��ʂ�\������B
        iResponse = MsgBox(txtExeName(Index).Text & "��o�^���珜�O���܂��B" _
                            & Chr(vbKeyReturn) & " ��낵���ł����H", _
                            vbYesNo + vbExclamation, _
                            "���s�t�@�C�����̓o�^���O")
        If iResponse = vbYes Then
        ' [�͂�] �{�^����I�������ꍇ
            '���s�t�@�C�����̕\���������B
            txtExeName(Index).Text = ""
            sChangeExePass(iSetupAplIndex) = ""
            '�uհè�è��ʁF�o�^�폜�v���O�o��
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_TOOL_SETTEI_DELETE, 0)
        Else
        ' [������] �{�^����I�������ꍇ
            '�������Ȃ��B
            Exit Sub
        End If
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdExecute_Click
'//  �@�\����  : �u�N���v�t����������
'//  �@�\�T�v  : �A�v���̋N�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@�@[IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.20.0.1) 2010-03-16  REVISED BY [TCC] S.Yoshimori
'//                 �t�@�C���I����ʂ�OS�d�l�ɕύX
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 �}�̎�O�s��C��
'//     REVISIONS :(2.8.0.1) 2011-02-07   REVISED BY [TCC] S.Terao
'//                 �z��Q�ƕs��C��
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdExecute_Click(Index As Integer)
    Dim lRetVal As Double     'Shell�֐��߂�l
    Dim iResponse As Integer  'MsgBox�{�^���R�[�h
    Dim iSetupAplIndex As Integer '�N���A�v���C���f�b�N�X
    
    Dim objFso As New FileSystemObject   '�t�@�C���V�X�e���I�u�W�F�N�g  'V1.20.0.1 ADD
   
On Error GoTo ERROR_MSG
    '�uհè�è��ʁF�N���t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_KIDOU_BUTTOM, 0)

    '��ʐݒ�C���f�b�N�X��0�`9�Ȃ̂ŁA�t�C���f�b�N�X�l���Z�o���A
    '�N���A�v���̃p�X�ŋN������B
    '�N���A�v���C���f�b�N�X=(���݉�ʐ�-1���)�~1��ʍő�t���{�����C���f�b�N�X(0�`9)
    '��F2��ʖڂ̉����t�C���f�b�N�X3���������ꂽ�ꍇ�A�N���A�v���p�X�C���f�b�N�X��13
    '13=(2-1)��10�{3
    iSetupAplIndex = (iGamenSts - 1) * 10 + Index

    '�N���Ώۂ��߽����`�`�F�b�N���s���B
    If (sChangeExePass(iSetupAplIndex) = "") Then
        '�N���Ώ��߽����`�������ꍇ�A�N���A�v���I���̂��߂̃t�@�C���I����ʂ�\������B
        'txtExeName(Index) = pfFileSelection("D:", "*.exe;*.com;*.bat;*.cmd", _
                                            "���s�t�@�C���I��")     'V1.20.0.1 DEL
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
        Set objFso = Nothing
        '�g���q��ݒ�
        CommonDialog1.Filter = _
            "�v���O�����t�@�C���i*.exe;*.com;*.bat;*.cmd�j|*.exe;*.com;*.bat;*.cmd|"
        '�t�@�C���I����ʂ��J��
        CommonDialog1.ShowOpen
        '�I�������t�@�C�������擾
        txtExeName(Index) = CommonDialog1.FileName
        'V1.20.0.1 ADD END
       sChangeExePass(iSetupAplIndex) = txtExeName(Index)
    End If
  
    '�N���Ώ��߽����`������ꍇ�B
    'If (sChangeExePass(Index) <> "") Then           'V2.8.0.1 DEL
    If (sChangeExePass(iSetupAplIndex) <> "") Then   'V2.8.0.1 ADD
    '�ݒ藓�ɃA�v���P�[�V����������΁A
        '�ݒ藓�̃A�v���P�[�V���������s����B
        lRetVal = Shell(sChangeExePass(iSetupAplIndex), vbNormalFocus)
        '�uհè�è��ʁF�c�[���N������v���O�o��
        Call sLogTraceReq(LTYP_NORMAL, L3AN_API, UTILITY_GAMEN_TOOL_OK, 0)
    End If
    Call ChDrive("D")  'V2.5.0.1 ADD
    Exit Sub

ERROR_MSG:
    '�uհè�è��ʁF�c�[���N���ُ�v���O�o��
     Call sLogTraceReq(LTYP_ERROR, L3AN_API, UTILITY_GAMEN_TOOL_ERROR, 0)
    '�u�N���ُ�v�|�b�v�A�b�v��ʂ�\������B
     iResponse = MsgBox("���s����A�v���P�[�V������" _
                        & Chr(vbKeyReturn) & "�������ݒ肵�Ă�������", _
                        vbYes, _
                        "�A�v�����s�G���[")
     Exit Sub
     Set objFso = Nothing    'V1.20.0.1 ADD
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdFixedExe_Click
'//  �@�\����  : �u�N��(�Œ�)�v�t����������
'//  �@�\�T�v  : �A�v���̋N�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@�@[IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)
    Dim lRetVal As Double      'Shell�֐��߂�l
    Dim iResponse As Integer   'MsgBox�߂�l
    Dim iSetupAplIndex As Integer '�N���A�v���C���f�b�N�X

On Error GoTo ERROR_MSG
    '�uհè�è��ʁF�N���t�����v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_KIDOU_BUTTOM, 0)
  
    '��ʐݒ�C���f�b�N�X��0�`9�Ȃ̂ŁA�t�C���f�b�N�X�l���Z�o���A
    '�N���A�v���̃p�X�ŋN������B
    '�N���A�v���p�X�C���f�b�N�X=(���݉�ʐ�-1���)�~1��ʍő�t���{�����C���f�b�N�X(0�`9)
    '��F2��ʖڂ̉����t�C���f�b�N�X3���������ꂽ�ꍇ�A�N���A�v���p�X�C���f�b�N�X��13
    '13=(2-1)��10�{3
    iSetupAplIndex = (iGamenSts - 1) * 10 + Index
        
    '�Y���{�^���̃A�v���P�[�V�������N������B
    lRetVal = Shell(sFixedExePass(iSetupAplIndex), vbNormalFocus)
    '�uհè�è��ʁF�c�[���N������v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_API, UTILITY_GAMEN_TOOL_OK, 0)
    Exit Sub
    
ERROR_MSG:
'===�A�v���N���G���[�̏ꍇ�A
    '�uհè�è��ʁF�c�[���N���ُ�v���O�o��
    Call sLogTraceReq(LTYP_ERROR, L3AN_API, UTILITY_GAMEN_TOOL_ERROR, 0)
    '�u�N�����s�v�|�b�v�A�b�v��ʂ�\������B
    iResponse = MsgBox(cmdFixedExe(Index).Caption & "�t�A��`�G���[�B" & _
                Chr(vbKeyReturn) & _
                sFixedExePass(iSetupAplIndex) & "���N���ł��܂���B", _
                vbYes, _
               "�Œ�N���A�v�����s�G���[")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Command1_Click
'//  �@�\����  : �u����ʁv�u�O��ʁv�t����������
'//  �@�\�T�v  : �u����ʁv�u�O��ʁv�t�����ɂ��A�Ώۉ�ʂ�\������B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Index�@�@�@[IN]�����t�C���f�b�N�X
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Command1_Click(Index As Integer)

  On Error Resume Next

  Select Case Index
   Case 0
     '�uհè�è��ʁF�O��ʖt�����v���O�o��
     Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_BACK_BUTTOM, 0)
     If iGamenSts = 1 Then
       '���ݕ\����ʐ��F1��ʖځB
       '���\����ʐ���2��ʖڂ̂��߁A���ݕ\����ʐ���2��ݒ肷��B
       iGamenSts = iGamenSts + 1
     Else
       '���ݕ\����ʐ��F2��ʖځB
       '�\���J�n�_��0�A���\����ʐ���1��ʖڂ̂��߁A���ݕ\����ʐ���1�ɐݒ肷��B
       iGamenSts = 1
       iHyoujiCnt = 0
     End If
      
     '�Œ�t�A�ύX�\�Œ�t�\���������s���B
     sSetAplClickDisplay
     sAplClick_Display
    
     '���ݕ\����ʐ��F1��ʎ��̂݁A�\���J�E���^�[�̃J�E���g�A�b�v�͂����ōs���B
     If iGamenSts = 1 Then
       iHyoujiCnt = 10
     End If
     
    Case 1
      '�uհè�è��ʁF����ʖt�����v���O�o��
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_NEXT_BUTTOM, 0)
 
       If iGamenSts = 2 Then
         '���ݕ\����ʐ��F2��ʖځB
         '�\���J�n�_��0�A���\����ʐ���1��ʖڂ̂��߁A���ݕ\����ʐ���1��ݒ肷��B
         iGamenSts = 1
         iHyoujiCnt = 0
       Else
         '���ݕ\����ʐ��F1��ʖځB
         '���\����ʐ���2��ʖڂ̂��߁A���ݕ\����ʐ���2��ݒ肷��B
         iGamenSts = iGamenSts + 1
       End If
     
        sSetAplClickDisplay
        sAplClick_Display
     
       '���ݕ\����ʐ��F1��ʎ��̂݁A�\���J�E���^�[�̃J�E���g�A�b�v�͂����ōs���B
       If iGamenSts = 1 Then
          iHyoujiCnt = 10
       End If
   Case Else
   '��������
   End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : cmdReturn_Click
'//  �@�\����  : �u�����e�i���X��ʂ֖߂�v�t����������
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
  
   '�uհè�è��ʁF�����v���O�o��
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UTILITY_GAMEN_END, 0)
   Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sFixedKoteiExeDisplay
'//  �@�\����  : �ύX�\�Œ�A�v���N���t�\����������
'//  �@�\�T�v  : �ύX�\�A�v���̏����������s���B
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
Private Sub sFixedKoteiExeDisplay()
    Dim lSts As Long    'INI�t�@�C���ݒ�擾�֐��̖߂�l
    Dim lCnt As Long    'HOSHUAPL.INI�̓o�^����
    Dim iMax As Integer '�A�v���C���f�b�N�X�ő�l
    Dim sWork As String * UTILITY_SIZE
    Dim i As Integer

    On Error Resume Next

    iMax = txtExeName.UBound '�A�v���p�XINDEX�̍ő�l�𓾂�B

    ' ���O�ݒ�t�@�C������ݒ�ύX�\�A�v���́u�o�^�����v����o��
    lCnt = GetPrivateProfileInt(PROFILE_SECTION_NAME, PROFILE_KEY_NAME_COUNT, _
                                DEFAILT_Int, HOSHUAPL_FILE)
    If (lCnt > 0) Then
    '�ݒ�ύX�\�A�v��������΁A
        ' �ݒ�ύX�\�A�v���̓o�^���������o��
        For i = 0 To lCnt - 1
            lSts = GetPrivateProfileString(PROFILE_SECTION_NAME, _
                                           PROFILE_KEY_NAME_HEAD & i, _
                                           DEFAILT, sWork, Len(sWork), HOSHUAPL_FILE)
             'INI�t�@�C���̎擾���ʃ`�F�b�N���s���B
             If lSts > 0 Then
               If i <= iMax Then
               'INI�t�@�C�����擾����A�܂�10���ȓ��̂��߁A��ʕ\������B
                txtExeName(i) = sWork
               End If
               '�ύX�\�Œ�N���t����̧���߽�G���A�ɁA�擾�߽���i�[����B
               sChangeExePass(i) = sWork
            End If
        Next i
    End If
    
    '�o�^�������A1��ʕ\���ő�10�ȏォ�ǂ����`�F�b�N����B
    If iMax < lCnt Then
     '10���ȏ�̏ꍇ�A�ύX�o�^��10���ȏ�t���O��ON�ɂ���B
     iChangeHyouji_Flag = 1
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sFixedExeDisplay
'//  �@�\����  : �Œ�A�v���N���t�\����������
'//  �@�\�T�v  : �Œ�A�v���̏����������s���B
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
Private Sub sFixedExeDisplay()
Dim i As Integer          'INI̧�كL�[�J�E���^�FDSPi ���N���tINDEX
Dim iMax As Integer       '�Œ�N���tINDEX�ő�l
Dim sLine As String * 256 '�P�s���̕�����B�i������hDSPi=�h�������j
Dim lSize As Long         '�P�s�����޲Đ��B�i������hDSPi=�h�������j
Dim iK As Integer         '�J���}�L�q�ʒu

On Error Resume Next
 
'�S�Ă̌Œ�N���t�ɂ��āA�ȉ������{����B
iMax = cmdFixedExe.UBound     '�Œ�N���tINDEX�̍ő�l�𓾂�B
 
 For i = 0 To iHoshuAplMax
   '�A�v���N�������lINI�t�@�C������A�P�s���̕�����iDSPi=�������j��Ǎ��ށB
    lSize = GetPrivateProfileString(PROFILE_SECTION_NAME_FIXED_EXE, _
                                    PROFILE_KEY_NAME_FIXED_EXE & CStr(i), _
                                    DEFAILT, sLine, Len(sLine), HOSHUAPL_FILE)
    iK = InStr(sLine, ",")        '�t�@�C�����i�t���p�X�j�̋�ؕ����ʒu�𓾂�B
    'INI�t�@�C���ɁA�Y���s�̒�`������ꍇ�A
    If lSize > 0 And iK <> 0 Then
     '�t�@�C�����Ɩt���̂���o���A�ۑ����Ă����B
      sFixedExePass(i) = Trim$(Left$(sLine, iK - 1))
      sFixedExeName(i) = Trim$(Mid$(sLine, iK + 1, lSize - iK))
    End If
Next i

For i = 0 To iMax
   '�Œ�N���t���\���ɂ���B
    cmdFixedExe(i).Visible = False
    '�N���A�v���p�X���ƁA�\���t���̂̒�`�`�F�b�N���s���B
    If sFixedExePass(i) <> "" And sFixedExeName(i) <> "" Then
       '��`�L��̏ꍇ�A�L���v�V�����ɋN���t�\��������������݁A�N���t��\������B
       cmdFixedExe(i).Visible = True
       cmdFixedExe(i).Caption = sFixedExeName(i)
    End If
    '�\���J�E���^�A�b�v����B
    iHyoujiCnt = iHyoujiCnt + 1
Next i

 For i = 0 To iHoshuAplMax
    If sFixedExePass(i) <> "" And sFixedExeName(i) <> "" And i > 9 Then
      iKoteiHyouji_Flag = 1
    End If
 Next i

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sSetAplClickDisplay
'//  �@�\����  : �u����ʁv�u�O��ʁv�t�����������B
'//  �@�\�T�v  : �Œ�A�v���N�����̕\���������s���B
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
Private Sub sSetAplClickDisplay()
Dim i As Integer          'INI̧�كL�[�J�E���^�FDSPi ���N���tINDEX
Dim iMax As Integer       '�Œ�N���tINDEX�ő�l
Dim iCnt As Integer       '�������[�v�J�E���^�[

On Error Resume Next

'�\���J�E���^�[��������[�v�J�E���^�[�Ɏ擾����B
iCnt = iHyoujiCnt
 
'�S�Ă̌Œ�N���t�ɂ��āA�ȉ������{����B
iMax = cmdFixedExe.UBound     '�Œ�N���tINDEX�̍ő�l�𓾂�B
For i = CNT_MIN To iMax
  '�Œ�N���t�������Ă����B
     cmdFixedExe(i).Visible = False
       '�N���A�v���p�X���ƁA�\���t���̂̒�`�`�F�b�N���s���B
       If sFixedExePass(iCnt) <> "" And sFixedExeName(iCnt) <> "" Then
         '��`�L��̏ꍇ�A�L���v�V�����ɋN���t�\��������������݁A�N���t��\������B
          cmdFixedExe(i).Visible = True
          cmdFixedExe(i).Caption = sFixedExeName(iCnt)
        End If
        '�������[�v�J�E���^�[���J�E���g�A�b�v����B
        iCnt = iCnt + 1
Next i
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : sAplClick_Display
'//  �@�\����  : �u����ʁv�u�O��ʁv�t�����������B
'//  �@�\�T�v  : �ύX�\�Œ�A�v���N�����̕\���������s���B
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
Private Sub sAplClick_Display()

Dim i As Integer          'INI̧�كL�[�J�E���^�FDSPi ���N���tINDEX
Dim iMax As Integer       '�Œ�N���tINDEX�ő�l
Dim iCnt As Integer       '�������[�v�J�E���^�[

On Error Resume Next

'�\���J�E���^�[��������[�v�J�E���^�[�Ɏ擾����B
iCnt = iHyoujiCnt

iMax = txtExeName.UBound '�A�v���p�XINDEX�̍ő�l�𓾂�B
For i = CNT_MIN To iMax
  '�\���������s���B
  txtExeName(i) = sChangeExePass(iCnt)
  '�������[�v�J�E���^�[���J�E���g�A�b�v����B
  iCnt = iCnt + 1
Next i

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Unload
'//  �@�\����  : ��ʏ������̐ݒ��INI�t�@�C���ɔ��f����B
'//  �@�\�T�v  : �u�����e�i���X��ʂ֖߂�v�t�����������F
'//              HOSHUAPL.INI�ւ̐ݒ���X�V����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : Integer�@Cancel
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer            '�J�E���^�[
    Dim l As Integer            '�o�^�����J�E���^�[
    Dim iMax As Integer         '���s�A�v���\����INDEX�̍ő�l
    Dim lSts As Boolean         'INI�t�@�C�����f�߂�l

    On Error Resume Next
    
    l = 0
    iMax = txtExeName.UBound   '���s�A�v���\����INDEX�̍ő�l���Z�b�g����B
   
   '�u���E�O��ʁv�t�\���L���`�F�b�N�B
   '�u���E�O��ʁv�t�̕\��������ꍇ�A�ő僋�[�v�J�E���^�[���ő�20�ɂ���B
   If iContinuFlag = True Then
      iMax = (iMax + 1) * 2
   End If

    For i = CNT_MIN To iMax
      If (sChangeExePass(i) <> "") Then
        l = l + 1
      End If
       '�A�v���N�������l�t�@�C���ɋN���A�v���̎��s�t�@�C�����������ށB
        lSts = WritePrivateProfileString(PROFILE_SECTION_NAME, _
               PROFILE_KEY_NAME_HEAD & CStr(i), sChangeExePass(i), HOSHUAPL_FILE)
    Next i
  
  '�o�^����(��`�L�茏��)���X�V����B
   lSts = WritePrivateProfileString(PROFILE_SECTION_NAME, _
          PROFILE_KEY_NAME_COUNT, CStr(l), HOSHUAPL_FILE)
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
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
  '���[������M����B
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       '�ێ��ʃA�N�e�B�u�v������M������A����ʂ�O�ʂɕ\��������B
        AppActivate frmUtility.Caption, False
        pfFormActive (frmUtility.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

