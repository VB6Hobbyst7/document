VERSION 5.00
Begin VB.Form dlgLogDataKakunin 
   BorderStyle     =   3  '�Œ��޲�۸�
   ClientHeight    =   2955
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
   ScaleHeight     =   2955
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�L�����Z��"
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
      Left            =   3240
      Style           =   1  '���̨���
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
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
      Left            =   1200
      Style           =   1  '���̨���
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '����
      Caption         =   " �����̎��ԑѕʃf�[�^��S�ăN���A���܂���    ��낵���ł����H"
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5535
   End
End
Attribute VB_Name = "dlgLogDataKakunin"
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
'    Call psPutLog(LOG_DLGLOGDATAKAKUNIN_CMDOKCLICK)
'
'    frmICMLogKanri.blnOk = True
'EG20 �ێ��ʃ��b�N�A�b�v�쐬 DEL END

'EG20 �ێ��ʃ��b�N�A�b�v�쐬 ADD START
    blnOk = True
'EG20 �ێ��ʃ��b�N�A�b�v�쐬 ADD END

    '��ʂ�Unload
    Unload Me

End Sub

'/////////////////////////////////////////////////////////////////////////////
'/    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'/
'/  �֐�����     : cmdCancel_Click
'/  �@�\����     : �u�L�����Z���v�{�^����������
'/  �@�\�T�v     : �u�L�����Z���v�{�^�������������s��
'/
'/                   �^          ����            �Ӗ�
'/  ����         :
'/  �߂�l       :
'/
'/  ORIGINAL     : (1.0.1.1) 2006-10-10  CODED BY [TCC] K.Hayashi
'/  REVISIONS    : (x.x.x.x) xxxx-xx-xx  CODED BY [xxx]
'/ ���l:
'/////////////////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    
'EG20 �ێ��ʃ��b�N�A�b�v�쐬 DEL START
'    '���O�o��
'    Call psPutLog(LOG_DLGLOGDATAKAKUNIN_CMDCANCELCLICK)
'
'    frmICMLogKanri.blnOk = False
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
'    Call psPutLog(LOG_DLGLOGDATAKAKUNIN_FORMLOAD)
'
'    '�z�u�ݒ�
'    Me.Top = DIALOGTOP
'    Me.Left = DIALOGLEFT
'    Me.Height = DIALOGHEIGHT
'    Me.Width = DIALOGWIDTH
'
'    If DISPSTS_CPU = gintDispStatus Then
'        '�������\��CPU
'        lblTitle(0).Caption = DEF_LOG_CPULABEL
'    Else
'        'EGX���D�@
'        lblTitle(0).Caption = DEF_LOG_EGXLABEL
'    End If
'EG20 �ێ��ʃ��b�N�A�b�v�쐬 DEL END
'EG20 �ێ��ʃ��b�N�A�b�v�쐬 ADD START
    
    'OK�t�����t���O������
    blnOk = False
    
    '�z�u�ݒ�
    Me.Top = 3495
    Me.Left = 2985
    Me.Height = 3375
    Me.Width = 6165
'EG20 �ێ��ʃ��b�N�A�b�v�쐬 ADD END


End Sub

