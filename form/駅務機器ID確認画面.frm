VERSION 5.00
Begin VB.Form frmEkimKikiId 
   BorderStyle     =   0  'Èµ
   Caption         =   "w±@íIDmF"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "lr SVbN"
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
   ScaleMode       =   0  'Õ°»Þ°
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMail 
      Left            =   9120
      Top             =   3480
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "}ÌæO"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9720
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "eLXg\¦"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   9720
      TabIndex        =   5
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "}ÌoÍ"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   9720
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ListBox ListEkimId 
      Height          =   7710
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   8775
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "lr SVbN"
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
      Caption         =   "  @íîñÝè    æÊÖßé"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9360
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblKan 
      Alignment       =   2  'µ¦
      BorderStyle     =   1  'Àü
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   8
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label lblKan 
      Alignment       =   2  'µ¦
      BorderStyle     =   1  'Àü
      Caption         =   "¼Ì"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'µ¦
      BackColor       =   &H00800000&
      Caption         =   "w±@íIDmF"
      BeginProperty Font 
         Name            =   "lr SVbN"
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
Attribute VB_Name = "frmEkimKikiId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  t@C¼  FfrmEkimKikiId.frm
'//  pbP[W¼Fw±@íIDmFæÊ
'//
'//  TvFw±@íIDmFæÊ
'//     ORIGINAL  :(1.4.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 Ûç_C³
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 tF[YR@¸@sïC³
'//     REVISIONS :(1.12.0.1) 2009-11-11  REVISED BY [TCC] C.Terui
'//                 w±@íIDÝæfBNgÊuÏX
'//     REVISIONS :(1.20.0.1) 2010-03-10  REVISED BY [TCC] S.Yoshimori
'//                 tH_IðæÊðOSdlÉÏX
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 t@CN[YÇÁ
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 yvOXo[\¦@\©¼µÎz
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 yeLXgoÍA}ÌoÍ{^Ì}~Îz
'//  õlF
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000   '[^C}ÌC^[ol
Private iSendType As Integer            'víÊl
Private Const EKIMU_DEFU = "APL\APL_WORK"

Private Const APL = "APL"
Private Const LOG = "LOG"
Private Const Data = "DATA"
Private Const BACKUP = "BACKUP"

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : Form_Activate
'//  @\¼Ì  : w±@íIDmFæÊ(ANeBu)
'//  @\Tv  : [óM^C}N®
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 Ûç_C³
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 yvOXo[\¦@\©¼µÎz
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 yeLXgoÍA}ÌoÍ{^Ì}~Îz
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    On Error Resume Next
    '[óM^C}ðN®·éB
    tmrMail.Enabled = True
    
'V1.7.0.1 ADD START
    Dim bRet As Boolean                 'ßèl
    Dim bFlag As Boolean                'tO
    Dim lngErrCode As Long              'G[R[h
    Dim udtMail As MAIL_INFO_CMD          'æÊ\¦v
    Dim uMail As ML_KYOTU_INF           '[
    Dim lLen  As Long
  
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ð\¦·é
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
  
   'obt@tbVvðOvZXÉM·é
   'îñvCMD(w±@íID=0)ðID§äÉM·é
   udtMail.mlHeader.dwId = ML_ID_INFO_CMD
   udtMail.mlHeader.dwSize = MlSize.INFO_CMD
   udtMail.mlHeader.dwProid = RHOSHU_ID
   udtMail.mlHeader.dwSubArea = 0
   udtMail.dwRequestType = MailCmdType.ML_DT_EKIMU_ID
   iSendType = MailCmdType.ML_DT_EKIMU_ID
   bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
   If bRet = False Then
      'uw±@íIDmFFîñvCMDMÙívOoÍ
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
      
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
      'vOXo[ðÁ·é
      Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
      
      Exit Sub
   Else
      'uw±@íIDmFFîñvCMDM³ívOoÍ
      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
      'æÊbN
      cmdVer(1).Enabled = False
      cmdVer(2).Enabled = False
      cmdInstall.Enabled = False
      cmdCancel.Enabled = False
   End If
   
    'obt@tbVI¹ÊmóM
    bFlag = False
    Do Until bFlag = True
       '[óMðs¤
       lLen = DssMailRead(plMSlot_MN, uMail)
       If lLen > 0 Then                            'óM³íÌ
         If ML_ID_INFO_RES = uMail.udtlHeader.dwId Then '[hc
            'îñvRESðóMµ½çAæÊ\¦pt@Cì¬ðs¤B
            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GETINFO_RES_RECV, 0)
            'víÊAÊðæ¾
            Call psDispID(uMail.lngData(1))
           'æÊbNð
' EG20 V6.3.0.1yeLXgoÍA}ÌoÍ{^Ì}~ÎzíJn
'           cmdVer(1).Enabled = True
'           cmdVer(2).Enabled = True
' EG20 V6.3.0.1yeLXgoÍA}ÌoÍ{^Ì}~ÎzíI¹
' EG20 V6.3.0.1yeLXgoÍA}ÌoÍ{^Ì}~ÎzÇÁJn
            If ListEkimId.ListCount > 0 Then
                cmdVer(1).Enabled = True
                cmdVer(2).Enabled = True
            End If
' EG20 V6.3.0.1yeLXgoÍA}ÌoÍ{^Ì}~ÎzÇÁI¹
           cmdInstall.Enabled = True
           cmdCancel.Enabled = True
           Exit Do
         End If
        End If
        Sleep (MN_MAIL_INTERVAL)
    Loop
'V1.7.0.1 ADD END
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ðÁ·é
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : Form_Deactivate
'//  @\¼Ì  : w±@íIDmFæÊ(fBANeBu)
'//  @\Tv  : [óM^C}â~
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
    On Error Resume Next
    '[óM^C}ðâ~·éB
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : Form_Load
'//  @\¼Ì  : w±@íIDmFæÊ([h)
'//  @\Tv  : úðs¤B
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-17   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 Ûç_C³
'//     REVISIONS :(EG20 V6.3.0.1) 2012-06-16   CODED   BY [TCC] H.Sugimoto
'//                 yeLXgoÍA}ÌoÍ{^Ì}~Îz
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
 'V1.7.0.1 DEL START
'   Dim udtMail As MAIL_INFO_CMD          'æÊ\¦v
'   Dim iResponse As Integer            'bZ[W{bNXßèl
'   Dim bRet As Boolean                 '[Mßèl
'   Dim lngErrCode As Long              'G[R[h
'   Dim bFlag As Boolean
'   Dim lId As Long
 'V1.7.0.1 DEL END
 
   On Error Resume Next
   
   Me.Top = 0
   Me.Left = 0
   Me.Height = 9000
   Me.Width = 12000
    
   'uw±@íIDmFæÊF\¦vOoÍ
   Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMUKIKI_ID_GAMEN_START, 0)
    
   '[óM^C}ÌC^[oð'PbÉZbg
   tmrMail.Interval = MN_MAIL_INTERVAL
   tmrMail.Enabled = False
 
' EG20 V6.3.0.1yeLXgoÍA}ÌoÍ{^Ì}~ÎzÇÁJn
    cmdVer(1).Enabled = False
    cmdVer(2).Enabled = False
' EG20 V6.3.0.1yeLXgoÍA}ÌoÍ{^Ì}~ÎzÇÁI¹
 'V1.7.0.1 DEL START
'   'îñvCMD(w±@íID=0)ðID§äÉM·é
'   udtMail.mlHeader.dwId = ML_ID_INFO_CMD
'   udtMail.mlHeader.dwSize = MlSize.INFO_CMD
'   udtMail.mlHeader.dwProid = RHOSHU_ID
'   udtMail.mlHeader.dwSubArea = 0
'   udtMail.dwRequestType = MailCmdType.ML_DT_EKIMU_ID
'   iSendType = MailCmdType.ML_DT_EKIMU_ID
'   bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
'   If bRet = False Then
'      'uw±@íIDmFFîñvCMDMÙívOoÍ
'      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
'      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
'   Else
'      'uw±@íIDmFFîñvCMDM³ívOoÍ
'      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
'      'æÊbN
'      cmdVer(1).Enabled = False
'      cmdVer(2).Enabled = False
'      cmdInstall.Enabled = False
'      cmdCancel.Enabled = False
'   End If
 'V1.7.0.1 DEL END
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'//
'//  Ö¼Ì  : cmdCancel_Click
'//  Tv     : uj[æÊÖßévtº
'//  à¾     : ©æÊðÁ·éB
'//  Êß×Ò°À   :
'//           :
'//
'//  ORIGINAL  F(1.4.0.1) 2009-03-23  CODED BY  [TCC] S.Terao
'//  REVISIONS F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
   
    On Error Resume Next
    
    'uw±@íIDmFæÊFÁvOoÍ
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMUKIKI_ID_GAMEN_END, 0)
    
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'//
'//  Ö¼Ì  : cmdInstall_Click
'//  Tv     : u}ÌæOvtº
'//  à¾     : }ÌðæèO·B
'//  Êß×Ò°À   :
'//           :
'//
'//  ORIGINAL  F(1.4.0.1) 2009-03-23  CODED BY  [TCC] S.Terao
'//  REVISIONS F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub cmdInstall_Click()
 On Error Resume Next
   
   'u}ÌæOtºvOoÍ
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '}ÌæO
    Call pfRemove(Me)
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'//
'//  Ö¼Ì  : cmdVer_Click
'//  Tv     : ueLXg\¦vu}ÌoÍvtº
'//  à¾     : etºðs¤B
'//  Êß×Ò°À   :
'//           :
'//
'//  ORIGINAL  F(1.4.0.1) 2009-03-23  CODED BY  [TCC] S.Terao
'//  REVISIONS F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub cmdVer_Click(Index As Integer)
    Dim lRetVal As Long             'ßèl
    Dim sCommand As String          'R}h¶ñ
    Dim lngErrCode As Long
    Dim bRet As Boolean
    
    On Error Resume Next
  
    Select Case Index

      Case 1
           'ueLXg\¦tFºvOoÍ
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMUKIKI_ID_TEXT_BUTTOM, 0)
           ' ÀsR}hðì¬
           sCommand = MN_EXE_MEMO & MN_VERSI_FILE
           ' ðN®·é¡
           lRetVal = Shell(sCommand, vbMaximizedFocus)
           ' ðANeBuiOÊ\¦jÉ·é
           AppActivate lRetVal, True
           SendKeys "{LEFT}", True
      Case 2
           'u}ÌoÍtFºvOoÍ
           Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, EKIMUKIKI_ID_OUTPUT_BUTTOM, 0)
           bRet = Text_OutPut
           If bRet = False Then
              'u}ÌoÍÙívOoÍ
              Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, EKIMUKIKI_ID_OUTPUT_ERROR, 0)
           Else
              'u}ÌoÍ³ívOoÍ
              Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, EKIMUKIKI_ID_OUTPUT_OK, 0)
           End If
           
    End Select
End Sub

'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2006 All Right Reserved
'//
'//  Ö¼Ì  : Text_Output
'//  Tv     : u}ÌoÍv
'//  à¾     : }ÌoÍðs¤B
'//  Êß×Ò°À   :
'//           :
'//
'//  ORIGINAL  F(1.4.0.1) 2009-03-23  CODED BY  [TCC] S.Terao
'//  REVISIONS F(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 w±@íIDÝæfBNgÊuÏX
'//  REVISIONS F(1.20.0.1) 2010-03-10   REVISED BY [TCC] S.Yoshimori
'//                 tH_IðæÊðOSdlÉÏX
'//  REVISIONS F(EG20 V2.0.1.1) 2011-11-21  REVISED BY [TCC] T.Koyama
'//                 dfQOtF[YQÎyc54z
'//                 EoÍt@C¼ÏX
'//  REVISIONS F(EG20 V3.0.0.2) 2011-12-22  CODED BY  [TCC] H.Sugimoto
'//                 EG20tF[YQÎ
'//                 EG20ÄÕUSDMÎÔyMainte_03_01z
'//  REVISIONS F(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 yvOXo[\¦@\©¼µÎz
'//  REVISIONS F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Function Text_OutPut() As Boolean
    Dim sCopyfile As String         'Rs[æ
    Dim sCopyTargetFile As String   'Rs[³
    Dim sLzhDirName As String
    Dim iResponse           As Integer          'MsgBoxßèl
    
' EG20 V2.0.1.1 ADD START
    Dim strStationName As String                ' w¼æ¾GA
' EG20 V2.0.1.1 ADD END
' EG20 V3.0.0.2ÇÁJn
    Dim fso         As New FileSystemObject     ' t@CVXeIuWFNg
    Dim textWrite   As TextStream               ' eLXgiCgj
    Dim textRead    As TextStream               ' eLXgi[hj
    Dim bWOpen      As Boolean
    Dim bROpen      As Boolean
    Dim strRecord   As String                   ' [N
' EG20 V3.0.0.2ÇÁI¹
    
On Error GoTo FileCopyError
  
    Text_OutPut = False

' EG20 V3.0.0.2ÇÁJn
    bWOpen = False
    bROpen = False
' EG20 V3.0.0.2ÇÁI¹
   
    'tH_IðæÊð\¦³¹At@Ci[fBNg¼ð¾éB
'    sLzhDirName = pfDirSelection("a:", "w±@íIDÝæÌfBNgIð")     'V1.12.0.1 DEL
    'sLzhDirName = pfDirSelection("H:", "w±@íIDÝæÌfBNgIð")      'V1.12.0.1 ADD    'V1.20.0.1 DEL
    sLzhDirName = ShowFolders(Me.hwnd, "tH_ðwèµÄ­¾³¢", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
    If sLzhDirName = "" Then
       Text_OutPut = True
       Exit Function
    End If
    
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ð\¦·é
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
    
' EG20 V2.0.1.1 DEL START
'    sCopyfile = sLzhDirName & EKIMU_ID_TXT
' EG20 V2.0.1.1 DEL END
' EG20 V2.0.1.1 ADD START
    'w¼æ¾
    strStationName = gsGetStationEkiName
    ' oÍt@C¼ì¬
    sCopyfile = sLzhDirName & strStationName & "_" & EKIMU_ID_TXT
' EG20 V2.0.1.1 ADD END
    
    sCopyTargetFile = MN_VERSI_FILE
    
' EG20 V3.0.0.2íJn
'    FileCopy sCopyTargetFile, sCopyfile
' EG20 V3.0.0.2íI¹
    
' EG20 V3.0.0.2ÇÁJn
    Set textWrite = fso.CreateTextFile(sCopyfile, True)
    bWOpen = True
    textWrite.WriteLine ("Ýuw@F" & strStationName)
    textWrite.WriteBlankLines (1)
    Set textRead = fso.OpenTextFile(sCopyTargetFile, ForReading, False)
    bROpen = True
    Do Until textRead.AtEndOfStream = True
        strRecord = textRead.ReadLine
        textWrite.WriteLine strRecord
    Loop
    textWrite.Close
    bWOpen = False
    textRead.Close
    bROpen = False
    Set textWrite = Nothing
    Set textRead = Nothing
    Set fso = Nothing
' EG20 V3.0.0.2ÇÁI¹
    
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ðÁ·é
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
    
    iResponse = MsgBox("³íI¹µÜµ½B", _
                       vbOKOnly, _
                       "}ÌoÍÊ")
    
    
    'fBXNîñðæ¾
    Text_OutPut = True
    
    Exit Function

FileCopyError:
' EG20 V3.1.0.2ÇÁJn
    If bWOpen = True Then
        textWrite.Close
        bWOpen = False
    End If
    If bROpen = True Then
        textRead.Close
        bROpen = False
    End If
    Set textWrite = Nothing
    Set textRead = Nothing
    Set fso = Nothing
' EG20 V3.1.0.2ÇÁI¹
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ðÁ·é
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
    
    iResponse = MsgBox("ÙíI¹µÜµ½B", _
                       vbOKOnly, _
                       "}ÌoÍÊ")
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : tmrMail_Timer
'//  @\¼Ì  : [óM^C}A^CAbv
'//  @\Tv  : [ðóM·éB
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 Ûç_C³
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
 'V1.7.0.1 DEL START
'  Dim lLen  As Long
'  Dim uMail As ML_KYOTU_INF           '[
'
'  On Error Resume Next
'
'  '[óM
'  lLen = DssMailRead(plMSlot_MN, uMail)
'  If lLen > 0 Then                            'óM³íÌ
'
'      Select Case uMail.udtlHeader.dwId  '[hc
'        Case ML_ID_HOSHU_ACTIVE_REQ
'            'ÛçæÊANeBuvðóMµ½çA©æÊðOÊÉ\¦³¹éB
'            AppActivate frmEkimKikiId.Caption, False
'            pfFormActive (frmEkimKikiId.hwnd)
'        Case ML_ID_INFO_RES
'            'îñvRESðóMµ½çAæÊ\¦pt@Cì¬ðs¤B
'            Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GETINFO_RES_RECV, 0)
'
'            'víÊAÊðæ¾
'            Call psDispID(uMail.lngData(1))
'        Case Else
'     End Select
'  End If
'  'æÊbNð
'  cmdVer(1).Enabled = True
'  cmdVer(2).Enabled = True
'  cmdInstall.Enabled = True
'  cmdCancel.Enabled = True
'V1.7.0.1 DEL END
'V1.7.0.1 ADD START
    'G[[`ðé¾
    On Error Resume Next
    
    'ÄpCóMðs¤
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmEkimKikiId.Caption, False
        pfFormActive (frmEkimKikiId.hwnd)
    End If
'V1.7.0.1 ADD END
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : psDispID
'//  @\¼Ì  : æÊ\¦
'//  @\Tv  : w±@íIDîñæÊ\¦ðs¤B
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Long     lngSts    [IN]Ê
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.8.0.1) 2009-08-27   REVISED BY [TCC] S.Terao
'//                 tF[YR@¸@sïC³
'//     REVISIONS :(1.21.0.1) 2010-04-09  REVISED BY [TCC] S.Terao
'//                 t@CN[YÇÁ
'//     REVISIONS :(EG20 V3.3.0.1) 2012-01-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20tF[YQÎyìì No.36ÖAz
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õl F
'///////////////////////////////////////////////////////////////////
Private Function psDispID(lngSts As Long)
    Dim sEkimuIDFile    As String   'w±@íIDt@CpX
    Dim iRet            As Integer  'INIæ¾ßèl
    Dim sFolder         As String * MAX_PATH_SIZE  'tH_¼
    Dim sFile           As String   't@C¼
    Dim MyName          As String   't@CõÊ
    Dim bRet            As Boolean  'ßèl
    Dim lngErrCode      As Long     'G[R[h
    Dim intFileNo       As Integer  't@CÔ
    Dim strWork         As String   'ìÆGA
    Dim dwErrsts        As Long
    Dim sFolderName     As String
        
    'ÊXe[^Xª³íÌêAiFàðs¤B
    If lngSts = 0 Then
      sFolder = ""
      
      'ÊF³íÍæÊ\¦
      iRet = GetPrivateProfileString(IDU_SECTION_NAME, _
                                     IDU_EKIMUID_KEY, _
                                     EKIMU_DEFU, sFolder, Len(sFolder), _
                                     PATH_IDU_INI_FILE)
      If iRet = 0 Then
        sFolder = EKIMU_DEFU
      End If
      sEkimuIDFile = ""
      'víÊlæèt@C¼ì¬
      sFile = Replace(EKIMU_ID_FILE, "##", Format(iSendType, "0#"))
      If iRet = 0 Then
         sFolderName = RTrim(sFolder)
      Else
         sFolderName = Mid(sFolder, 1, iRet)
      End If
      'pXÏ·
      sFolderName = pfChangeFolderName(sFolderName)
      'w±@íIDt@CpXì¬
      sEkimuIDFile = sFolderName & "\" & sFile
      't@CL³`FbN
      If Dir(sEkimuIDFile, vbNormal) = "" Then
         Exit Function
      End If
      
      '/////////////////////////////////////////////////////////////////////
      '//ÛçêpÖFw±@íIDæÊ\¦pt@Cì¬
      '////////////////////////////////////////////////////////////////////
      'bRet = dllEKIMUKIKI(sEkimuIDFile, dwErrsts, MN_VERSI_FILE) 'V1.8.0.1 DEL
      bRet = dllEKIMUKIKI(sEkimuIDFile, dwErrsts, MN_VERSI_FILE, PATH_IDU_APP) 'V1.8.0.1 ADD

      If dwErrsts = 1 Then
         'G[R[hF³í
         'Xgú»
         ListEkimId.Clear

        'VBG[
        On Error GoTo Error_psVersionDisp
    
        'w±@íIDæÊ\¦pt@CÌt@CÔðæ¾·éB
        intFileNo = FreeFile
      
        'w±@íIDæÊ\¦pt@CI[v
        Open MN_VERSI_FILE For Input As #intFileNo
    
        'Xg\¦ªÇÝÝit@CI[ÜÅ[vðJèÔ·j
'        Do While Not EOF(1)                                ' EG20 V3.3.0.1í
        Do While Not EOF(intFileNo)                         ' EG20 V3.3.0.1ÇÁ
           'ìÆGAðú»
           strWork = ""

           Line Input #intFileNo, strWork
           
           'üsR[hÌÝÍÇÝÆÎ·
           If Trim(strWork) <> "" Then
              'XgÉoÍ
              ListEkimId.AddItem (strWork)
           End If
        Loop
         
        't@CN[Y
        Close #intFileNo
      Else
        'G[R[hFÙí
        Exit Function
     End If
   Else
     'ÊFÙíÍ½àµÈ¢
   End If
Exit Function

'VBG[
Error_psVersionDisp:
    'V1.21.0.1 ADD  START
    If intFileNo > 0 Then
        Close #intFileNo
    End If
    'V1.21.0.1 ADD  END
    'uw±@íIDmFæÊFo[Wîñt@Cì¬ÙívOoÍ
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : pfChangeFolderName
'//  @\¼Ì  : tH_pXÏ·
'//  @\Tv  : INIt@Cæèæ¾µ½tH_è`ÌÏ·ðs¤B
'//
'//              ^        ¼Ì         Ó¡
'//  ø      : String sFolderName    [IN]INIè`
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õl F
'///////////////////////////////////////////////////////////////////
Private Function pfChangeFolderName(sFolderName As String) As String
   Dim iPath As Integer
   Dim sRootPath As String
   Dim sFolder As String
      
   'uvÊuðæ¾
   iPath = InStr(sFolderName, "\")
   If iPath = 0 Then
     sRootPath = Mid(sFolderName, 1)
   Else
     'uvO¶ñðæ¾
     sRootPath = Mid(sFolderName, 1, iPath - 1)
     'uvã¶ñðæ¾
     sFolder = Mid(sFolderName, iPath + 1)
   End If
   Select Case sRootPath
      Case APL
        'Av[g
        sRootPath = PATH_IDU_APP
      Case LOG
        'O[g
        sRootPath = PATH_IDU_LOG
      Case Data
        'DB[g
        sRootPath = PATH_IDU_DB
      Case BACKUP
        'obNAbv[g
        sRootPath = PATH_BUC
   End Select
    'pXA
    pfChangeFolderName = sRootPath + "\" + sFolder
End Function
