VERSION 5.00
Begin VB.Form dlgLogDataKekka 
   BorderStyle     =   3  '�Œ��޲�۸�
   ClientHeight    =   3195
   ClientLeft      =   3015
   ClientTop       =   4200
   ClientWidth     =   6030
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   9
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
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "�n�j"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  '���̨���
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackStyle       =   0  '����
      Caption         =   "�����I�����܂����B"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   5295
   End
End
Attribute VB_Name = "dlgLogDataKekka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'/
'/  �֐�����     : cmdOK_Click
'/  �@�\����     : �uOK�v�{�^����������
'/  �@�\�T�v     : �uOK�v�{�^�������������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.1.1) 2006-10-10  CODED BY [TCC] K.Hayashi
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx  CODED BY [xxx]
'/ ���l:
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()
    
'EG20 �ێ��ʃ��b�N�A�b�v�쐬 DEL START
'    '���O�o��
'    Call psPutLog(LOG_DLGLOGDATAKEKKA_CMDOKCLICK)
'EG20 �ێ��ʃ��b�N�A�b�v�쐬 DEL END

    '��ʂ�Unload
    Unload Me

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'/
'/  �֐�����     : Form_Load
'/  �@�\����     : Form_Load������
'/  �@�\�T�v     : Form_Load���������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.1.1) 2006-10-10  CODED BY [TCC] K.Hayashi
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx  CODED BY [xxx]
'/ ���l:
'/////////////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
'EG20 �ێ��ʃ��b�N�A�b�v�쐬 DEL START
'    '���O�o��
'    Call psPutLog(LOG_DLGLOGDATAKEKKA_FORMLOAD)
'
'    '�z�u�ݒ�
'    Me.Top = DIALOGTOP
'    Me.Left = DIALOGLEFT
'    Me.Height = DIALOGHEIGHT
'    Me.Width = DIALOGWIDTH
'
'    lblTitle = "����I�����܂����B"
'EG20 �ێ��ʃ��b�N�A�b�v�쐬 DEL END
'EG20 �ێ��ʃ��b�N�A�b�v�쐬 ADD START
    '�z�u�ݒ�
    Me.Top = 3495
    Me.Left = 2985
    Me.Height = 3375
    Me.Width = 6165
'EG20 �ێ��ʃ��b�N�A�b�v�쐬 ADD END

End Sub

