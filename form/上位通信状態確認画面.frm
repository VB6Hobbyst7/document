VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOverConectSts 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "��ʒʐM��Ԋm�F"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   11.25
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
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
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
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Left            =   7560
      Top             =   8040
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   10425
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   $"��ʒʐM��Ԋm�F���.frx":0000
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
      Left            =   9120
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid GridIni 
      Height          =   6355
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11218
      _Version        =   393216
      Rows            =   10
      Cols            =   3
      RowHeightMin    =   50
      WordWrap        =   -1  'True
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      HighLight       =   0
      ScrollBars      =   2
      MergeCells      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00800000&
      Caption         =   "��ʒʐM��Ԋm�F"
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
      TabIndex        =   2
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmOverConectSts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  �t�@�C����  �FfrmOverConectSts.frm
'//  �p�b�P�[�W���F��ʒʐM��Ԋm�F���
'//
'//  �T�v�F��ʒʐM��Ԋm�F���
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.3.0.1) 2009-03-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R�Ď��Ձ@������Ή�
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y�ڑ��@�팩�����Ή��z
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Option Explicit

'///////////////////////////////////////////////////////////////////
'�h�m�h�t�@�C�����i�[�G���A
'///////////////////////////////////////////////////////////////////
Private iConectSts As Integer          '�O���@��ʐM��Ԓl
Private iSts_Naiyou As Integer         '�O���@��ُ��Ԓl
Private iSts_Type As Integer           '�O���@��ُ��ʒl
Private Const CONECTSTS_NORMAL = 1     '�ʐM����
Private Const CONECTSTS_SOKET = 1      '�\�P�b�g���x���ُ�
Private Const CONECTSTS_TCP = 2        'TCP���x���ُ�
Private Const CONECTSTS_APL = 3        '�A�v���P�[�V�������x���ُ�
Private Const CONECTSTS_GETERR = 4     '��Ԏ擾�ُ�
Private udtAreaR255 As GATE_INFO                                    '�Ǎ��ݗp�G���A�i255�ݒ�p�j

'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   '���[���^�C�}�̃C���^�[�o���l

' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ��J�n
' ��ʋ@��ݒ�\��
Private Type TRANSKIKI_INFO
    bStatus As Boolean              ' �ݒ�L���iTRUE:�L��,FALSE:�����j
    sGetInf As String               ' ��ʕ\���p����
    iAreaID As Integer              ' �ΏۊO���@���ʋ@��ʐM��ԃG���AID
    nIniListNo As Integer           ' �O���@�탊�X�g�ԍ�
    nCorner As Integer              ' �R�[�i�ԍ�
    nProcType As Integer            ' �����^�C�v
    iErrorInfoID As Integer         ' �ʐM�ُ��ԃG���AID
    iErrorTypeID As Integer         ' �ʐM�ُ��ʃG���AID
End Type
Private gTransKikiInfo(1 To CONECT_KIKI_INI_MAX) As TRANSKIKI_INFO

Private Const PROCTYPE_NORMAL = 0   ' �ʏ폈���i�Q�ƃG���A����ʋ@��ʐM��ԃG���A�j
Private Const PROCTYPE_ENKAKU = 1   ' ���u�����i�Q�ƃG���A�����u�^�C�v�j
' EG20 V3.4.0.1�y�ڑ��@�팩�����Ή��z�ǉ��I��


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : Form_Activate
'//  �@�\����  : ��ʒʐM��Ԋm�F���(�A�N�e�B�u��)
'//  �@�\�T�v  : ��ʂ̍őO�ʕ\���������s���B
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
'//  �@�\����  : ��ʒʐM��Ԋm�F���(�f�B�A�N�e�B�u��)
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
'//  �@�\����  : ��ʒʐM��Ԋm�F���(���[�h��)
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
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R�Ď��Ձ@������Ή�
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim i As Integer
    Dim ii As Integer
    Dim iWide As Integer
    
    On Error Resume Next

   '�u��ʒʐM��Ԋm�F��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, OVER_CONECT_STS_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
      
    'V2.3.0.1 ADD START
    'IDU�k�ރ`�F�b�N
    psIDUCheck
    'V2.3.0.1 ADD END

    '��ʒʐM��ԕ\������
    psConectSts
   
   'V1.3.0.1 ADD START
   '���[����M�^�C�}�̃C���^�[�o����'�P�b�ɃZ�b�g
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   'V1.3.0.1 ADD END
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
      
   '�u��ʒʐM��Ԋm�F��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, OVER_CONECT_STS_GAMEN_END, 0)
    frmOverConectSts.ZOrder
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  �֐�����  : Command1_Click
'//  �@�\����  : �u�\���X�V�v�t����������
'//  �@�\�T�v  : ��ʒʐM��ԕ\���������ĂсA�\���X�V���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V2.0.1.1) 2011-11-21  CODED  BY [TCC] T.Koyama
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub Command1_Click()

    On Error Resume Next

   '�u��ʒʐM��Ԋm�F��ʁF�\���v���O�o��
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, OVER_CONECT_STS_GAMEN_START, 0)
    
    'IDU�k�ރ`�F�b�N
    psIDUCheck

    '��ʒʐM��ԕ\������
    psConectSts

End Sub
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : psConectSts
'//  �@�\����  : ��ʒʐM��Ԃ�\������B
'//  �@�\�T�v  : �Ώۏ�ʋ@��̒ʐM��Ԃ̎擾�\�����s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : �Ȃ�
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(2.3.0.1) 2010-10-19   REVISED BY [TCC] T.Arai
'//                 EG-R�Ď��Ձ@������Ή�
'//     REVISIONS :(EG20 V2.0.1.1) 2011-11-21   REVISED BY [TCC] T.Koyama
'//                 �d�f�Q�O�t�F�[�Y�Q�Ή��y�c��54�z
'//                 �E��ԕ\�����ւ̃X�N���[���o�[�ǉ������
'//                   �\���X�V�t�������̃Z������ǉ�
'//     REVISIONS :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y�ڑ��@�팩�����Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psConectSts()
  Dim iCnt As Integer                     'INI�t�@�C���ǂݍ��݃J�E���^�[
  Dim sKey As String                      '�L�[��
  Dim sGetInf As String * OVERCONECT_SIZE '�擾���(�\������)
  Dim lSts As Long                        'INI�擾�����߂�l
  Dim i As Integer                        '�O���b�h�̍����J�E���^�[
  Dim iErrCnt   As Integer                '�O�l�ߗp�J�E���^�[
  Dim sErr_TCP As String                  'TCP���x���ُ핶��
  Dim sErrCode As String                  '�G���[�R�[�h
  Dim iAreaID As Integer                  '�擾���(�G���AID) 'V2.3.0.1 ADD
  Dim iAddRow As Integer                  ' �o�^�s��            ' EG20 V3.4.0.1�ǉ�
  Dim szResultName As String              ' �o�͖���            ' EG20 V3.4.0.1�ǉ�
   
  On Error Resume Next
   
' EG20 V3.4.0.1�ǉ��J�n
  '���@���擾
  Call gsGetGateInfo
  ' �R�[�i���̐ݒ菈��
  Call gsGetCornerName
' EG20 V3.4.0.1�ǉ��I��
   
  '�O���b�h�̕ύX
  With GridIni
        '�O���b�h�̏�����
        .Clear

        '�O���b�h�̃Z�����̕ύX
'        .Rows = 11                     ' EG20 V3.4.0.1�폜
        iAddRow = 1                     ' EG20 V3.4.0.1�ǉ�
        .Rows = iAddRow                 ' EG20 V3.4.0.1�ǉ�
        .Cols = 4

        '�ݒ�l�̃^�C�g���Z�b�g
        .Row = 0
        .Col = 1: .Text = "��ʋ@��"
        .CellAlignment = flexAlignCenterCenter
           
        .Col = 2: .Text = "�ʐM���"
        .CellAlignment = flexAlignCenterCenter
        
        .Col = 3: .Text = "�ڍ�"
        .CellAlignment = flexAlignCenterCenter
        
        '�O���@�햼�̂�\��
        For iCnt = 1 To CONECT_KIKI_INI_MAX

            gTransKikiInfo(iCnt).bStatus = False               ' �ݒ�L���iTRUE:�L��,FALSE:�����j
            gTransKikiInfo(iCnt).sGetInf = ""                  ' ��ʕ\���p����
            gTransKikiInfo(iCnt).iAreaID = 0                   ' �ΏۊO���@���ʋ@��ʐM��ԃG���AID
            gTransKikiInfo(iCnt).nIniListNo = 0                ' �O���@�탊�X�g�ԍ�
            gTransKikiInfo(iCnt).nCorner = 0                   ' �R�[�i�ԍ�
            gTransKikiInfo(iCnt).nProcType = 0                 ' �����^�C�v
            gTransKikiInfo(iCnt).iErrorInfoID = 0              ' �ʐM�ُ��ԃG���AID
            gTransKikiInfo(iCnt).iErrorTypeID = 0              ' �ʐM�ُ��ʃG���AID

         'V2.3.0.1 ADD START
         ' OUTKIKI_LIST.ini�����ʒʐM�G���AID���擾����B
         sKey = ""
         sKey = PROFILE_KEY_KIKIAREA_NAME & Format(iCnt, "00")
         iAreaID = GetPrivateProfileInt(PROFILE_SECTION_LIST_NAME, _
                                        sKey, _
                                        DEFAILT_Int, _
                                        OUTKIKI_LIST_FILE)

         'IDU�ݒu���������݂̏�ʋ@��ʐM��ԃG���A��ID�T�[�o�Ŗ����ꍇ
         '�܂��́AIDU�ݒu�L��̏ꍇ�́A�ȍ~�̕\���������s���B
         If (pbIDUSts = 1 And iAreaID <> IdKikiComSts.ID_SERVER_COM) Or _
            (pbIDUSts = 0) Then
         'V2.3.0.1 ADD END

           ' OUTKIKI_LIST.ini����\���p�O���@�햼�̂��擾����B
           sKey = PROFILE_KEY_KIKINAME & Format(iCnt, "00")
           lSts = GetPrivateProfileString(PROFILE_SECTION_LIST_NAME, _
                                          sKey, _
                                          DEFAILT, _
                                          sGetInf, _
                                          Len(sGetInf), _
                                          OUTKIKI_LIST_FILE)
' EG20 V3.4.0.1�ǉ��J�n
           If lSts <> False Then
                ' �o�͖��̎擾����
                Call psAddKikiCornerName(sGetInf, iAreaID, iCnt)
                If gTransKikiInfo(iCnt).bStatus = False Then
                    lSts = False
                End If
           End If
' EG20 V3.4.0.1�ǉ��I��
           If lSts = False Then
             'INI�ݒ薳���̏ꍇ�A�������Ȃ�
           Else
             iAddRow = iAddRow + 1
             .Rows = iAddRow                                    ' EG20 V3.4.0.1�ǉ�
             iErrCnt = iErrCnt + 1
             .Row = iErrCnt
'             .Col = 1: .Text = sGetInf                         ' EG20 V3.4.0.1�폜
             .Col = 1: .Text = gTransKikiInfo(iCnt).sGetInf     ' EG20 V3.4.0.1�ǉ�
        
' EG20 V3.4.0.1�폜�J�n
'             '�e�O���@��ʐM��Ԏ擾�������s���B
'             pfGetConectSts iCnt
' EG20 V3.4.0.1�폜�I��
' EG20 V3.4.0.1�ǉ��J�n
            If gTransKikiInfo(iCnt).nProcType = PROCTYPE_NORMAL Then
                '��ʋ@��ʐM��Ԏ擾�������s���B
                pfGetConectSts iCnt
            Else
                '��ʋ@��ʐM��Ԏ擾�������s���B
                pfGetConectStsJikai iCnt
            End If
' EG20 V3.4.0.1�ǉ��I��
             
             '�ʐM��ԃX�e�[�^�X�Q��
             Select Case iConectSts
               Case CONECTSTS_NORMAL
                  '�ʐM��ԁF����
                  .Col = 2: .Text = "����"
                  .CellAlignment = flexAlignCenterCenter
               Case CONECTSTS_GETERR
                  '�ʐM��ԁF�擾�ُ�
                  .Col = 2: .Text = ""
                  .CellAlignment = flexAlignCenterCenter
               Case Else
                  '��L�ȊO�F�\�P�b�g���x���ُ�,TCP���x���ُ�,�A�v�����x���ُ�
                  .Col = 2: .Text = "�ُ�"
                  .CellAlignment = flexAlignCenterCenter
             End Select
           
             '�e�O���@��ڍו\���������s���B
             '�ʐM��ԃX�e�[�^�X�Q��
             Select Case iSts_Naiyou
              Case CONECTSTS_SOKET
                 '�\�P�b�g���x���ُ�
                 .Col = 3: .Text = "�\�P�b�g���Ȃ���Ȃ�"
                 .CellAlignment = flexAlignCenterCenter
              Case CONECTSTS_TCP
                 'TCP���x���ُ�
                 '16�i���ɕϊ��B
                 sErrCode = Hex(iSts_Type)
                 sErrCode = sErrCode & "h"
                 sErr_TCP = "TCP���x���łȂ���Ȃ�(�G���[�R�[�h:" & sErrCode & ")"
                 .Col = 3: .Text = sErr_TCP
                 .CellAlignment = flexAlignCenterCenter
              Case CONECTSTS_APL
                 '�A�v���P�[�V�������x���ُ�
                 .Col = 3: .Text = "�A�v���P�[�V�������x���łȂ���Ȃ�"
                 .CellAlignment = flexAlignCenterCenter
              Case Else
                 '�ʐM��Ԑ���/�ʐM��Ԏ擾�ُ펞
                 .Col = 3: .Text = ""
                 .CellAlignment = flexAlignCenterCenter
             End Select
           End If
         End If 'V2.3.0.1 ADD
        Next
   
        '�O���b�h�̕��ύX
        .ColWidth(0) = 0
' EG20 V3.4.0.1�폜�J�n
'        .ColWidth(1) = 2500
'        .ColWidth(2) = 1500
'' EG20 V2.0.1.1 DEL START
''        .ColWidth(3) = 8000
'' EG20 V2.0.1.1 DEL END
'' EG20 V2.0.1.1 ADD START
'        .ColWidth(3) = 7775
'' EG20 V2.0.1.1 ADD END
' EG20 V3.4.0.1�폜�I��
' EG20 V3.4.0.1�ǉ��J�n
        .ColWidth(1) = 3000
        .ColWidth(2) = 1200
        .ColWidth(3) = 7575
        For i = iAddRow To 10
            .Rows = i + 1
        Next
' EG20 V3.4.0.1�ǉ��I��
        
        For i = 0 To CONECT_KIKI_INI_MAX
        '1�O���b�h�̍����ݒ�
         .RowHeight(i) = 570
        Next
         
' EG20 V2.0.1.1 ADD START
        .TopRow = 1
' EG20 V2.0.1.1 ADD END
    
    End With
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  �֐�����  : pfGetConectSts
'//  �@�\����  : ��ʋ@��ʐM��ԃG���A����Ԏ擾����
'//  �@�\�T�v  : ��ʋ@��ʐM��ԃG���A���
'//              �ʐM��ԁA�ُ��ԁA�ُ��ʂ̎擾���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : iCnt�@�@Integer�@�@[IN]�擾�J�E���^�[
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetConectSts(iCnt As Integer)
    Dim iAreaID As Integer  '�G���AID
    Dim sKey As String      '�L�[��
    Dim strMutexName    As String               '�~���[�e�b�N�X��
    Dim lngMuHandle     As Long                 '�r�������p�n���h��
    Dim udtMapInf       As MAP_MEM              '�������}�b�s���O�I�u�W�F�N�g
    Dim GetSts          As Long
         
    On Error Resume Next
   
    strMutexName = "Mu_" & GOverComSts
    lngMuHandle = dllOpenMutex(strMutexName)         '�r������(OPEN)
    If lngMuHandle = 0 Then
        '�f�[�^�Q�ƈُ펞�͏�ԃX�e�[�^�X�Ɏ擾�ُ��ݒ�B
        iConectSts = CONECTSTS_GETERR
        iSts_Naiyou = CONECTSTS_GETERR
        Exit Function
    End If
    
    dllCloseHandle (lngMuHandle)                 '�r������(CLOSE)
    
    Set Idinf_Jyoui = New IdInfProc                    '��ʒʐM��ԃG���A

    Idinf_Jyoui.ProcMode = DATA_ID.Data_Id_Jyoui                '��ʒʐM��ԃG���A
    Idinf_Jyoui.IdOpen
    If Idinf_Jyoui.Errsts <> 0 Then
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Exit Function
    End If
    
    ' OUTKIKI_LIST.ini����G���AID���擾����B
    sKey = PROFILE_KEY_KIKIAREA_NAME & Format(iCnt, "00")
    iAreaID = GetPrivateProfileInt(PROFILE_SECTION_LIST_NAME, _
                                   sKey, _
                                   DEFAILT_Int, _
                                   OUTKIKI_LIST_FILE)
    If iAreaID = 0 Then
      '�擾�ُ�̏ꍇ���̓ǂݍ��݂ցB
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Idinf_Jyoui.IdFree
       Exit Function
    Else
    
        '�Q��(��ʋ@��ʐM���)�G���A����ݒ�
         Idinf_Jyoui.ProcMode = DATA_ID.Data_Id_Jyoui
         Idinf_Jyoui.IdOpen
         If Idinf_Jyoui.Errsts <> 0 Then
           '�f�[�^�Q�ƈُ펞�͏�ԃX�e�[�^�X�Ɏ擾�ُ��ݒ�B
           iConectSts = CONECTSTS_GETERR
           iSts_Naiyou = CONECTSTS_GETERR
           Exit Function
         End If
         
         '�Q��(��ʋ@��ʐM���)�G���A���k�n�b�j����B
         Idinf_Jyoui.IdLock
         If Idinf_Jyoui.Errsts <> 0 Then
           '�f�[�^�Q�ƈُ펞�͏�ԃX�e�[�^�X�Ɏ擾�ُ��ݒ�B
           iConectSts = CONECTSTS_GETERR
           iSts_Naiyou = CONECTSTS_GETERR
           Idinf_Jyoui.IdFree
           Exit Function
         End If
                    
         '�G���A�̓��e��ǂݍ��ށB
         Idinf_Jyoui.id = iAreaID
           
         '�ʐM��Ԃ��擾
         Idinf_Jyoui.GetInf (CONECT)
         If Idinf_Jyoui.Errsts <> 0 Then
            '�f�[�^�Q�ƈُ펞�͏�ԃX�e�[�^�X�Ɏ擾�ُ��ݒ�B
            iConectSts = CONECTSTS_GETERR
            iSts_Naiyou = CONECTSTS_GETERR
            Idinf_Jyoui.IdFree
            Exit Function
         End If
         iConectSts = CInt(Idinf_Jyoui.DataArea(0))
           
         '�ُ��Ԃ��擾
         Idinf_Jyoui.GetInf (STS)
         If Idinf_Jyoui.Errsts <> 0 Then
            '�f�[�^�Q�ƈُ펞�͏�ԃX�e�[�^�X�Ɏ擾�ُ��ݒ�B
            iConectSts = CONECTSTS_GETERR
            iSts_Naiyou = CONECTSTS_GETERR
            Idinf_Jyoui.IdFree
            Exit Function
         End If
         iSts_Naiyou = CInt(Idinf_Jyoui.DataArea(0))
           
         '�ُ��ʂ��擾
         Idinf_Jyoui.GetInf (ERR_TYPE)
         If Idinf_Jyoui.Errsts <> 0 Then
            '�f�[�^�Q�ƈُ펞�͏�ԃX�e�[�^�X�Ɏ擾�ُ��ݒ�B
             iConectSts = CONECTSTS_GETERR
             iSts_Naiyou = CONECTSTS_GETERR
             Idinf_Jyoui.IdFree
             Exit Function
          End If
           iSts_Type = CInt(Idinf_Jyoui.DataArea(0))
          
    End If
     
    Idinf_Jyoui.IdFree
    
    Set Idinf_Jyoui = Nothing                     '��ʒʐM��ԃG���A
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
        AppActivate frmOverConectSts.Caption, False
        pfFormActive (frmOverConectSts.hwnd)
    End If
End Sub
'V1.3.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : psAddKikiCornerName
'//  �@�\����  : ��ʋ@��R�[�i���̒ǉ�����
'//  �@�\�T�v  : ��ʋ@�햼�̂ɑ΂��ăR�[�i���̂�t������K�v������Βǉ�����B
'//
'//              �^        ����      �Ӗ�
'//  ����      : String �@ sName     [IN]��ʋ@�햼��
'//  ����      : Integer�@ iAreaID   [IN]��ʋ@��ʐM��ԃG���AID
'//  ����      : Integer�@ nIndex    [IN]��ʋ@��ݒ�\��
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(EG20 V3.4.0.1) 2012-02-13  CODED BY  [TCC] H.Sugimoto
'//                 EG20�t�F�[�Y�Q�Ή��y�ڑ��@�팩�����Ή��z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Sub psAddKikiCornerName(sName As String, iAreaID As Integer, nIndex As Integer)

    Dim nCorner As Integer                  ' �R�[�i�C���f�b�N�X
    Dim szCornerName As String              ' �R�[�i����
    Dim nNullIndex As Integer               ' ���������[�N
    Dim szResultName As String              ' �o�͖���

    szResultName = ""
    nCorner = 0                                         ' �R�[�i�ݒ�s�v
    gTransKikiInfo(nIndex).nIniListNo = nIndex          ' �O���@�탊�X�g�ԍ�
    gTransKikiInfo(nIndex).nProcType = PROCTYPE_NORMAL  ' �����^�C�v
    gTransKikiInfo(nIndex).iAreaID = iAreaID            ' �Q�ƃG���A
    ' 1.�ΏۊO���@���ʋ@��ʐM��ԃG���AID���`�F�b�N����
    '   �ڑ��Ώۂ�I�ʂ���B
    Select Case iAreaID
    Case IdKikiComSts.ID_DESYU_COM                                       ' 1:�f�W�ʐM���
        nCorner = 1
    Case IdKikiComSts.ID_DESYU2_COM                                      ' 9:�f�W2�ʐM���
        nCorner = 2
    Case IdKikiComSts.ID_DESYU3_COM                                      ' 10:�f�W3�ʐM���
        nCorner = 3
    Case IdKikiComSts.ID_DESYU4_COM                                      ' 11:�f�W4�ʐM���
        nCorner = 4
    Case IdKikiComSts.ID_DESYU5_COM                                      ' 12:�f�W5�ʐM���
        nCorner = 5
    Case IdKikiComSts.ID_DESYU6_COM                                      ' 13:�f�W6�ʐM���
        nCorner = 6
    Case IdKikiComSts.ID_ENKAKU_COM                                      ' 2:���u�ʐM���
        gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU
        nCorner = 1
    Case IdKikiComSts.ID_ENKAKU2_COM                                     ' 21:���u2�ʐM��ԁi�G���A��`�Ȃ��j
        gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU
        nCorner = 2
    Case IdKikiComSts.ID_ENKAKU3_COM                                     ' 22:���u3�ʐM��ԁi�G���A��`�Ȃ��j
        gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU
        nCorner = 3
    Case IdKikiComSts.ID_ENKAKU4_COM                                     ' 23:���u4�ʐM��ԁi�G���A��`�Ȃ��j
        gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU
        nCorner = 4
    Case IdKikiComSts.ID_ENKAKU5_COM                                     ' 24:���u5�ʐM��ԁi�G���A��`�Ȃ��j
        gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU
        nCorner = 5
    Case IdKikiComSts.ID_ENKAKU6_COM                                     ' 25:���u6�ʐM��ԁi�G���A��`�Ȃ��j
        gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU
        nCorner = 6
    Case Else
    End Select

    gTransKikiInfo(nIndex).nCorner = nCorner
    If gTransKikiInfo(nIndex).nProcType = PROCTYPE_ENKAKU Then
        gTransKikiInfo(nIndex).iAreaID = IdGate.ENKAKUKIKI_JIKAIAREAID
        gTransKikiInfo(nIndex).iErrorInfoID = IdGate.ENKAKUKIKI_JIKAIERRSTATUSID
        gTransKikiInfo(nIndex).iErrorTypeID = IdGate.ENKAKUKIKI_JIKAIERRTYPEID
    End If
    
    If nCorner <> 0 Then
        If gblnCornerSet(nCorner - 1) <> True Then
            Exit Sub
        End If
        ' �R�[�i���̂̕t��
        nNullIndex = InStr(gstrCornerName(nCorner - 1), Chr(0))
        If nNullIndex <> 0 Then
            szCornerName = vbCrLf & Left(gstrCornerName(nCorner - 1), nNullIndex - 1)
        Else
            szCornerName = vbCrLf & gstrCornerName(nCorner - 1)
        End If
    End If
    szResultName = Left(sName, InStr(sName, Chr(0)) - 1)
    szResultName = szResultName + szCornerName
    gTransKikiInfo(nIndex).sGetInf = szResultName
    gTransKikiInfo(nIndex).bStatus = True

End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2012 All Rights Reserved
'//
'//  �֐�����  : pfGetConectStsJikai
'//  �@�\����  : ������ԃG���A����Ԏ擾����
'//  �@�\�T�v  : ������ԃG���A���
'//              �ʐM��ԁA�ُ��ԁA�ُ��ʂ̎擾���s���B
'//
'//              �^        ����      �Ӗ�
'//  ����      : iCnt�@�@Integer�@�@[IN]�擾�J�E���^�[
'//
'//              �^        �l        �Ӗ�
'//  �߂�l    : �Ȃ�
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  ���l�F
'///////////////////////////////////////////////////////////////////
Private Function pfGetConectStsJikai(iCnt As Integer)
    Dim iAreaSts As Integer                 ' �Ď��ݒ��Ԓl
        
    Dim iJikaiaArea_Jyotai As Integer       ' ������ԃG���A��Ԓl
    Dim lngMuHandle As Long                 ' �r�������p�n���h��
    Dim strMutexName As String
        
    Dim iAreaID As Integer                  ' �ʐM��ԃG���AID
    Dim iGokiNo As Integer                  ' ������Ԃ̍��@
    Dim iErrorInfoID As Integer             ' �ʐM�ُ��ԃG���AID
    Dim iErrorTypeID As Integer             ' �ʐM�ُ��ʃG���AID
        
    On Error Resume Next
    
    strMutexName = "Mu_" & GGateStatus
    lngMuHandle = dllOpenMutex(strMutexName)            '�r������(OPEN)
    If lngMuHandle = 0 Then
       '�G���A�Q�ƕs�̂��߁A�Q�ƈُ�
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Exit Function
    End If
  
    dllCloseHandle (lngMuHandle)                 '�r������(CLOSE)
    
    ' �ݒ���̎擾
    iAreaID = gTransKikiInfo(iCnt).iAreaID
    iErrorInfoID = gTransKikiInfo(iCnt).iErrorInfoID
    iErrorTypeID = gTransKikiInfo(iCnt).iErrorTypeID
    iGokiNo = gTransKikiInfo(iCnt).nCorner
    
    Set Idinf_JikaiJyotai = New IdInfProc              '������ԃG���A
    '�Q��(�������)�G���A����ݒ�
    Idinf_JikaiJyotai.ProcMode = DATA_ID.Data_Id_JkaiJyotai    '������ԃG���A
    Idinf_JikaiJyotai.IdOpen
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
        Set Idinf_JikaiJyotai = Nothing               '������ԃG���A
       Exit Function
    End If
    
    '�Q��(�������)�G���A���k�n�b�j����B
    Idinf_JikaiJyotai.IdLock
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       '�f�[�^�Q�ƈُ펞�̓u�����N�\���ݒ���s���B
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Idinf_JikaiJyotai.IdFree
       Set Idinf_JikaiJyotai = Nothing               '������ԃG���A
       Exit Function
    End If
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // �ʐM���
    '�G���A�̓��e��ǂݍ��ށB
    Idinf_JikaiJyotai.id = iAreaID
    Idinf_JikaiJyotai.GetJikai_Sts iGokiNo - 1
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       '�f�[�^�Q�ƈُ펞�̓u�����N�\���ݒ���s���B
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Idinf_JikaiJyotai.IdFree
       Set Idinf_JikaiJyotai = Nothing               '������ԃG���A
       Exit Function
    End If
   
    '�ʐM��Ԃ��擾
    iJikaiaArea_Jyotai = CInt(Idinf_JikaiJyotai.DataArea(iGokiNo - 1))
    iConectSts = iJikaiaArea_Jyotai
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // �ʐM�ُ���
    '�G���A�̓��e��ǂݍ��ށB
    Idinf_JikaiJyotai.id = iErrorInfoID
    Idinf_JikaiJyotai.GetJikai_Sts iGokiNo - 1
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       '�f�[�^�Q�ƈُ펞�̓u�����N�\���ݒ���s���B
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Idinf_JikaiJyotai.IdFree
       Set Idinf_JikaiJyotai = Nothing               '������ԃG���A
       Exit Function
    End If
  
    '�ʐM�ُ��Ԃ��擾
    iJikaiaArea_Jyotai = CInt(Idinf_JikaiJyotai.DataArea(iGokiNo - 1))
    iSts_Naiyou = iJikaiaArea_Jyotai
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // �ʐM�ُ���
    '�G���A�̓��e��ǂݍ��ށB
    Idinf_JikaiJyotai.id = iErrorTypeID
    Idinf_JikaiJyotai.GetJikai_Sts iGokiNo - 1
    If Idinf_JikaiJyotai.Errsts <> 0 Then
       '�f�[�^�Q�ƈُ펞�̓u�����N�\���ݒ���s���B
       iConectSts = CONECTSTS_GETERR
       iSts_Naiyou = CONECTSTS_GETERR
       Idinf_JikaiJyotai.IdFree
       Set Idinf_JikaiJyotai = Nothing               '������ԃG���A
       Exit Function
    End If
  
    '�ʐM�ُ��ʂ��擾
    iJikaiaArea_Jyotai = CInt(Idinf_JikaiJyotai.DataArea(iGokiNo - 1))
    iSts_Type = iJikaiaArea_Jyotai
     
    Idinf_JikaiJyotai.IdFree
    Set Idinf_JikaiJyotai = Nothing                 '������ԃG���A
     
End Function


