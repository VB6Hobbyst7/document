VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRMente 
   BorderStyle     =   0  'Èµ
   Caption         =   "Og[XiEG-R©®üD@j"
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
   PaletteMode     =   1  'Z µ°ÀÞ°
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Ìùèl
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdRemove 
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
      Left            =   9500
      TabIndex        =   11
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "t@Cí"
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
      Index           =   5
      Left            =   9500
      TabIndex        =   8
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "³k}ÌoÍ"
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
      Index           =   3
      Left            =   9500
      TabIndex        =   6
      Top             =   3480
      Width           =   2415
   End
   Begin VB.ListBox lstTraceFile 
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   240
      MultiSelect     =   2  'g£
      TabIndex        =   5
      Top             =   1080
      Width           =   9135
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "³kÊmF"
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
      Index           =   4
      Left            =   9500
      TabIndex        =   4
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "   t@C     }ÌoÍ"
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
      Left            =   9500
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "\¦XV"
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
      Left            =   9500
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdTraceFile 
      Caption         =   "f[^ûW "
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
      Index           =   0
      Left            =   9500
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "[geiXæÊÖßé"
      BeginProperty Font 
         Name            =   "lr oSVbN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9500
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   9120
      Top             =   8040
   End
   Begin VB.Label Label1 
      Alignment       =   2  'µ¦
      BackColor       =   &H00800000&
      Caption         =   "©®üD@[geiX"
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
      TabIndex        =   12
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblListItem 
      BorderStyle     =   1  'Àü
      Caption         =   "    g[Xt@C¼"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   10
      Top             =   840
      Width           =   7455
   End
   Begin VB.Label lblListItem 
      Alignment       =   2  'µ¦
      BorderStyle     =   1  'Àü
      Caption         =   "oCgTCY"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Caption         =   "©®üD@  g[Xt@C"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   450
      Width           =   4335
   End
End
Attribute VB_Name = "frmRMente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  t@C¼  FfrmRMente.frm
'//  pbP[W¼F©®üD@[geiXæÊ
'//
'//  TvF©®üD@[geiXæÊ
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-07-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 Ûç_C³
'//     REVISIONS :(1.12.0.1) 2009-11-11  REVISED BY [TCC] C.Terui
'//                 g[Xt@CÝæfBNgÊuÏX
'//                 g[Xt@C³kÝæfBNgÊuÏX
'//                 ³kt@CIðæfBNgÊuÏX
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 tH_IðæÊðOSdlÉÏX
'//                 t@CIðæÊðOSdlÉÏX
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 }ÌæOsïC³
'//     REVISIONS :(EG20 V3.6.0.1) 2012-02-20  CODED BY  [TCC] H.Sugimoto
'//                 EG20tF[YQÎyTR-No.272C³Îz
'//     REVISIONS :(EG20 V5.6.0.1) 2012-04-05  CODED BY  [TCC] H.Sugimoto
'//                 y³ktH_wèÎz
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 O}ÌoÍAãÀðTPQÆ·é
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 yvOXo[\¦@\©¼µÎz
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014Nx{ô yEG20_KANSI05_01z
'//  õlF
'///////////////////////////////////////////////////////////////////
Option Explicit

Private Const MN_MAIL_INTERVAL = 1000     '[^C}ÌC^[ol

'Xg{bNXÉÖ·él
Private Const LIST_FILE_SIZE_LENGTH = 11   'ÊÞ²Ä»²½ÞÌ¶
Private Const LIST_FILE_ELIMITTER = " -- " 'ÊÞ²Ä»²½ÞÆÄÚ°½Ì§²ÙÔÌæØ¶ñ
Private Const LIST_HEDDER_LENGTH = LIST_FILE_SIZE_LENGTH + 4 'ãLAQÂÌ¶v
Private sTOOLPass As String
'Private sHyoujiGoukiNo(0 To 18) As String         '\¦@Ôi[GA          ' EG20 V3.6.0.1yTR-No.272C³Îzí
Private sHyoujiGoukiNo(0 To 31) As String         '\¦@Ôi[GA           ' EG20 V3.6.0.1yTR-No.272C³ÎzÇÁ
Private Const TITLENAME_CORNER = "R[i#"        ' R[i¼                        ' EG20 V6.6.0.1ÇÁ
Private sRonriCornerNo(0 To 31) As String         '_R[iÔi[GA         ' EG20 V6.6.0.1ÇÁ

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : CmdRemove_Click
'//  @\¼Ì  : u}ÌæOvtº
'//  @\Tv  : }ÌÌæèOµðs¤B
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub cmdRemove_Click()
   On Error Resume Next
   
   'u}ÌæOtºvOoÍ
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, USB_OUT, 0)
 
   '}ÌæO
    Call pfRemove(Me)
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : Form_Activate
'//  @\¼Ì  : ©®üD@[geiX(ANeBu)
'//  @\Tv  : [óMp^C}AN®
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
On Error Resume Next
    '^C}ðN®·é
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : Form_Deactivate
'//  @\¼Ì  : ©®üD@[geiX(fBANeBu)
'//  @\Tv  : [óMp^C}Aâ~
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next

   If blnCabfrmOpenFlg = True Then
      Call fnTsbCabCallDiverge
     Exit Sub
   End If

    '^C}ð~ßé
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : Form_Load
'//  @\¼Ì  : ©®üD@[geiX([h)
'//  @\Tv  : úðs¤B
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim iRet As Integer
    
On Error Resume Next
    'u©®üD@ØÓ°ÄÒÝÃÅÝ½æÊF\¦vOoÍ
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_START, 0)

' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ð\¦·é
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹

   'GLTt@Cðì¬µAàeðXV·éB
    iRet = fMakeGLTFile
    
    If iRet = 0 Then
        'Xg{bNXÉg[Xt@C¼ð\¦·éB
        fListDisplay
    End If
    
    '[óMpÌ[óMpÌ^C}lðÝè·é
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ðÁ·é
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : cmdTraceFile_Click
'//  @\¼Ì  : etº
'//  @\Tv  : et¼ÌÌðs¤B
'//              uf[^ûWvu\¦XVvut@C}ÌoÍv
'//              u³k}ÌoÍvu³kÊmFvut@Cív
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Integer@Index@@ [IN]ºtCfbNX
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.6.0.1) 2009-07-16   REVISED BY [TCC] S.Terao
'//     REVISIONS :(1.7.0.1) 2009-07-28   REVISED BY [TCC] S.Terao
'//                 Ûç_C³
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub cmdTraceFile_Click(Index As Integer)
    Dim lRetVal As Double      'ShellÖßèl
    Dim iResponse As Integer   'MsgBoxßèl
    Dim sWriteDir As String    'g[Xt@CÝæÌfBNg
    Dim lngErrCode As Long   'G[R[h
   
   On Error Resume Next

    Select Case Index
    Case 0
       'u©®üD@ØÓ°ÄÒÝÃÅÝ½æÊFf[^ûWtºvOoÍ
        Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_DATA_SHUSHU_BUTTOM, 0)
        '©w.GLTt@CÖ©üîñðÞB
        fMakeGLTFile
        '©üSWÛçf[^ì¬ðs¤B
        'If sSWFileCopy > 0 Then   'V1.6.0.1 DEL
        sSWFileCopy  'V1.6.0.1 ADD
          '[gec[ðN®·éB
          psGATERMenteTool
          '©®üD@c[N®
          lRetVal = Shell(sTOOLPass, vbNormalFocus)
          If 0 = lRetVal Then
             GoTo ERROR_MSG_RMENTE
          End If
          '[gec[ðANeBuiOÊ\¦jÉ·é
        '  AppActivate lRetVal, True 'V1.7.0.1 DEL
        'V1.6.0.1 DEL START
        'Else
        '  'uØÓ°ÄÒÝÃÅÝ½æÊF©üÛçSWf[^t@CRs[ÙívOoÍ
        '  lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
        '  Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, CREATE_FILE_ERROR, lngErrCode)
        'End If
        'V1.6.0.1 DEL END
    Case 1
      'u©®üD@ØÓ°ÄÒÝÃÅÝ½æÊF\¦XVtºvOoÍ
       Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, UPDATE_BUTTOM, 0)
       'Xg{bNXÉg[Xt@C¼ð\¦·éB
       fListDisplay
    Case 2
      'u©®üD@ØÓ°ÄÒÝÃÅÝ½æÊFt@C}ÌoÍtºvOoÍ
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_FILE_OUTPUT_BUTTOM, 0)
      sCopyTraceFile
    Case 3
      'u©®üD@ØÓ°ÄÒÝÃÅÝ½æÊF³k}ÌoÍtºvOoÍ
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_LZH_OUTPUT_BUTTOM, 0)
      sLzhFileWrite
    Case 4
      'u©®üD@ØÓ°ÄÒÝÃÅÝ½æÊF³kÊmFtºvOoÍ
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_LZH_KAKUNIN_BUTTOM, 0)
      '³kt@CÌàeð\¦·éB
      sLzhFileDisplay
    Case 5
      'u©®üD@ØÓ°ÄÒÝÃÅÝ½æÊFt@CítºvOoÍ
      Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_FILE_DELETE_BUTTOM, 0)
       'Iðt@Cðí·éB
        If fSelectedFilesDelete = True Then
            'ít@Cª Á½ÈçAXg{bNXð\¦XV·éB
            fListDisplay
        End If
    Case Else
 End Select

 Exit Sub

ERROR_MSG_RMENTE:
'===g[Xf[^ûWG[ÌêA
    'u©®üD@ØÓ°ÄÒÝÃÅÝ½æÊF[gec[N®ÙívOoÍ
     Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, RMENTE_GAMEN_KIDOU_ERROR, 0)
    'u[gec[N®Ùív|bvAbvð\¦·éB
    iResponse = MsgBox("[gec[iR_Mente.exejðN®Å«Ü¹ñB", _
                vbYes, _
               "[gec[ÀsG[")
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : cmdReturn_Click
'//  @\¼Ì  : uj[æÊÖßévtº
'//  @\Tv  : ©æÊðÁ·éB
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
On Error Resume Next
    'u©®üD@ØÓ°ÄÒÝÃÅÝ½æÊFÁvOoÍ
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, RMENTE_GAMEN_END, 0)
    '©æÊðÁ·B
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : fMakeGLTFile
'//  @\¼Ì  : ©w.GLTt@CÖÌ©üîñð«Ý
'//  @\Tv  : GATE.INIðQÆµA©w.GLTt@CÖA
'//              @ÔA\¦¶AIPAhXð«ÞB
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.6.0.1) 2012-07-01 REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(EG20 V6.7.0.1)  2012-06-28  CODED BY  [TCC] H.Sugimoto
'//                 yÚ`FbNÌÎÛðüD@îñÌÝÆ·éC³z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Function fMakeGLTFile() As Integer
    Dim lngRet As Long          'ÖÌÔèl
    Dim iGate As Integer        '©üINDEX
    Dim j As Integer            '[NINDEX
    Dim sGoukiNo As String      'GLTt@CR[hf[^(@Ô\¦¶)
    Dim cWork As Byte           '[NGA
    Dim lngErrCode As Long      'G[R[h
    Dim sKeyName As String
    Dim sGateData As String * RMENTE_GATE_SIZE    'Psªt@Càeæ¾p
    Dim sFData() As String
    Dim iFCnt As Integer
    Dim iFLoop As Integer
    Dim iFLoop2 As Integer
    Dim iRet As Integer
    Dim intGLTFileNo As Integer     'Ì§²ÙÔ
    Dim szCorner As String      ' R[iÔ
    Dim szTitleName As String                       ' ^Cg¼                    ' EG20 V6.7.0.1ÇÁ
    Dim fso As New FileSystemObject                 't@CVXeIuWFNg   ' EG20 V6.7.0.1ÇÁ

    On Error Resume Next
    MkDir PATH_RMENTE_GATE_DEN   '©üpdStH_ðì¬·éBiGLTt@Cpj
    
' EG20 V6.7.0.1ÇÁJn
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(PATH_RMENTE_GATE_DEN_JIEKI) = False Then
        'Rs[ætH_ì¬
        fso.CreateFolder (PATH_RMENTE_GATE_DEN_JIEKI)
    End If
    Set fso = Nothing
' EG20 V6.7.0.1ÇÁI¹
    
    
    'GLTt@CðJ­Bt@Cª¶ÝµÈ¯êÎVKÉì¬³êéB
    On Error GoTo ErrorHandlerGLTFile
    intGLTFileNo = FreeFile        ' ¢gpÌt@CÔðæ¾·éB
    Open GATE_GLT_FILE_FULLPASS For Output As #intGLTFileNo     ' GLTt@CðJ­B

    For iGate = CNT_MIN To MAX_GATE_NO - 1
      '©®üD@îñæ¾
      sKeyName = "gate" & Format(iGate + 1, "00")
      iRet = GetPrivateProfileString(SETTEIFILE_INZ_SECTION_NAME, _
                                     sKeyName, _
                                     DEFAILT, sGateData, Len(sGateData), _
                                     PATH_GATE_FILE)
      If iRet = 0 Then
         'uØÓ°ÄÒÝÃÅÝ½æÊF©®üD@INIt@CÇÙívOoÍ
         Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, GATE_INI_READ_ERROR, 0)
         Exit Function
      End If
        
      If Len(sGateData) <> 0 Then
         'f[^Ìæ¾
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
      
      If Len(Trim(sFData(1))) = 1 Then
         '@ÔªPÈçÎAæªÉOðtÁ·éB
'         sGoukiNo = "0" & Trim(sFData(1)) & "@"                                 ' EG20 V6.7.0.1í
         sGoukiNo = "0" & Trim(sFData(1))                                           ' EG20 V6.7.0.1ÇÁ
      Else
'         sGoukiNo = Trim(sFData(1)) & "@"                                       ' EG20 V6.7.0.1í
         sGoukiNo = Trim(sFData(1))                                                 ' EG20 V6.7.0.1ÇÁ
      End If
        
' EG20 V6.6.0.1 y@ÔÉR[iÔðtÁ·éÎzÇÁJn
'        szCorner = Replace(TITLENAME_CORNER, "#", Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))) ' EG20 V6.7.0.1í
        szCorner = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))                                  ' EG20 V6.7.0.1ÇÁ
        sRonriCornerNo(iGate) = Trim(sFData(GATE_IDX.IDX_RONRI_CORNER))
' EG20 V6.6.0.1 y@ÔÉR[iÔðtÁ·éÎzÇÁI¹
' EG20 V6.7.0.1 y@ÔÉR[iÔðtÁ·éÎzÇÁJn
        ' R[iÔÏ·
        szTitleName = Replace(RMENTE_GOKITITLENAME, "$", szCorner)
        ' @ÔÏ·
        szTitleName = Replace(szTitleName, "##", sGoukiNo)
' EG20 V6.7.0.1 y@ÔÉR[iÔðtÁ·éÎzÇÁJn
      
      If Trim(sFData(4)) <> "" Then
         'Gate.init@CÌ@Ô\¦¶AIPAhXðGLTt@CÉ«ÞB
'         Print #intGLTFileNo, sGoukiNo & "," & Trim(sFData(5))                     ' EG20 V6.6.0.1í
'         Print #intGLTFileNo, szCorner & "_" & sGoukiNo & "," & Trim(sFData(5))    ' EG20 V6.6.0.1ÇÁ ' EG20 V6.7.0.1í
         Print #intGLTFileNo, szTitleName & "," & Trim(sFData(5))                   ' EG20 V6.7.0.1ÇÁ
      End If
      
      '\¦@Ô
      If Len(Trim(sFData(1))) = 1 Then
         '@ÔªPÈçÎAæªÉOðtÁ·éB
         sHyoujiGoukiNo(iGate) = "0" & Trim(sFData(1))
      Else
         sHyoujiGoukiNo(iGate) = Trim(sFData(1))
      End If
    
    Next
    
    'GLTt@CðÂ¶éB
    Close #intGLTFileNo
    
    fMakeGLTFile = 0    '³íI¹
    Exit Function

ErrorHandlerGateIni:
   'u©®üD@ØÓ°ÄÒÝÃÅÝ½æÊFt@CANZXÙívOoÍ
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeGLTFile = 1
   'GLTt@CðÂ¶éB
   Close #intGLTFileNo
   Exit Function
ErrorHandlerGLTFile:
   'u©®üD@ØÓ°ÄÒÝÃÅÝ½æÊFt@CANZXÙívOoÍ
   lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
   Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_ACCESS_ERROR, lngErrCode)
   fMakeGLTFile = 2
   'GLTt@CðÂ¶éB
   Close #intGLTFileNo

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : sSWFileCopy
'//  @\¼Ì  : ©üÛçSWÝèf[^t@Cì¬
'//  @\Tv  : ©üÛçSWÝèf[^ðA©üÛçSWf[^t@CÉ
'//              Rs[·éB
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V6.6.0.1) 2012-07-01 REVISED BY [TCC] H.Sugimoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Function sSWFileCopy() As Integer

     Dim iCnt As Integer                     'JE^[
     Dim sSWDataPath As String               '©üÛçSWf[^t@C
     Dim sMyPath As String                   '©üÛçSWÝèf[^
     
     On Error Resume Next
   
     sSWFileCopy = 0                         't@C¶Ý
    
    '©üÅåª[v·éB
    For iCnt = 1 To MAX_GATE_NO
     'uGATE_SW##.datvÌu##vð01`16ÉÏ··éB
     sMyPath = Replace(GATE_SW_FILE, "##", Format(iCnt, "0#"))
     '©üÛçSWÝèf[^Ìõðs¤B
     If Dir(sMyPath) <> "" Then
        '©üÛçSWf[^t@CÌpXðì¬·éB
        sSWDataPath = PATH_RMENTE_GATE_DEN_JIEKI_GOUKI
' EG20 V6.6.0.1ÇÁJn
        'uR[i$vÌu$vð1`6ÉÏ··éB
        sSWDataPath = Replace(sSWDataPath, "$", sRonriCornerNo(iCnt - 1))
' EG20 V6.6.0.1ÇÁI¹
        'u##@vÌu##vð01`16ÉÏ··éB
        sSWDataPath = Replace(sSWDataPath, "##", Format(sHyoujiGoukiNo(iCnt - 1), "0#"))
        'tH_ì¬
        MkDir sSWDataPath
        sSWDataPath = sSWDataPath & TOOL_SW_File
        
        '©üÛçSWf[^ð©üÛçSWf[^t@CÉRs[·éB
        FileCopy sMyPath, sSWDataPath
        sSWFileCopy = sSWFileCopy + 1
     End If
   Next
   
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : fListDisplay
'//  @\¼Ì  : Xg{bNXÌàeð\¦XV·éB
'//  @\Tv  : Xg{bNXÌ\¦àeðÁãA
'//              ÅVÌg[Xt@C¼ð\¦·éB
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Function fListDisplay()
    Dim sInFolder(1) As String  'g[Xf[^tH_¼

    On Error Resume Next

    'Xg{bNXðóÉ·éB
    lstTraceFile.Clear
    'g[Xf[^tH_ÈºÌt@CðXg{bNXÉ\¦·éB
    sInFolder(0) = PATH_RMENTE_GATE_DEN_JIEKI  '{dStH_©çJn·éB
    sFileDisplay 1, sInFolder()
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : sFileDisplay
'//  @\¼Ì  : Xg{bNX\¦
'//  @\Tv  : wètH_¼ºÌt@C¼ðXg{bNXÉ\¦·éB
'//              ÅVÌg[Xt@C¼ð\¦·éB
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : String@@sFolder
'//        @@: Integer @iFolderNo
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub sFileDisplay(iFolderNo As Integer, sFolder() As String)
    Dim iInFileNo As Integer   'õÎÛtH_¼ºÌt@CÌÂ
    Dim sInFile() As String    '  ¯ã  t@C¼itpXj
    Dim iInFolderNo As Integer 'õÎÛtH_¼ºÌt@CÌÂ
    Dim sInFolder() As String  '  ¯ã  tH_¼itpXFÅI¶Íj
    Dim i     As Integer       '[NJE^
    Dim j     As Integer       '[NJE^
    Dim sFileSize As String * LIST_FILE_SIZE_LENGTH  '\¦t@CÌoCgTCY
    Dim sDisplay As String     'Xg{bNXÖ\¦·éPsªÌ¶ñ

    On Error Resume Next

    'wè³ê½tH_ÌSÄÉÂ¢ÄÀ{·éB
    For i = CNT_MIN To iFolderNo - 1
        'õÎÛtH_¼ºÌt@CEtH_ðæ¾·éB
        psFolderSearch sFolder(i), iInFileNo, sInFile(), iInFolderNo, sInFolder()
        'õÎÛtH_¼ºÌt@CðXg{bNXÖ\¦·éB
        For j = 0 To iInFileNo - 1
            'Ì§²Ù»²½ÞÍElßARÌJ}æØèÅ\¦·éB
            RSet sFileSize = Format$(FileLen(sInFile(j)), "#,###")
            't@C¼ÍAEE\©dS\©w\ÜÅÌtH_:RMENTE_DIR_TRACEÍ\¦µÈ¢B
            '            iæªÉæØè¶:LIST_FILE_ELIMITTERðt¯éBj
            sDisplay = sFileSize & LIST_FILE_ELIMITTER & _
                       Right(sInFile(j), Len(sInFile(j)) - Len(PATH_RMENTE_GATE_DEN_JIEKI))
            lstTraceFile.AddItem sDisplay
        Next
        'õÎÛtH_¼ºÌtH_ÈºÌt@CðXg{bNXÉ\¦·éB
        sFileDisplay iInFolderNo, sInFolder()
    Next
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : sCopyTraceFile
'//  @\¼Ì  : ut@C}ÌoÍvtº
'//  @\Tv  : t@CðwèfBNgÉoÍ·éB
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 g[Xt@CÝæfBNgÊuÏX
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 tH_IðæÊðOSdlÉÏX
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 O}ÌoÍAãÀðTPQÆ·é
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 yvOXo[\¦@\©¼µÎz
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub sCopyTraceFile()
    Dim iLine As Integer         'ÄÚ°½Ì§²ÙØ½ÄÎÞ¯¸½Ìs²ÝÃÞ¯¸½
    Dim iMaxLine As Integer      'ÄÚ°½Ì§²ÙØ½ÄÎÞ¯¸½Ìs
    Dim iFlag As Integer         'Iðt@CL³iP^Oj
    Dim iResponse As Integer     'MsgBox{^R[h
    Dim sFullPass As String      'Rs[³t@CtpX¼
    Dim sFileName As String      'Rs[³t@C¼
    Dim sCopyDir As String       'Rs[æfBNg
    Dim sCopyFileName As String  'Rs[æt@C¼
    Dim lSts As Long             '[Nißèlj
    Dim sWork As String          '[N
    Dim i As Integer             '[N
    Dim j As Integer             '[N
    Dim lngErrCode As Long       'G[R[h
    Dim iFileCounter As Integer  'ÎÛÌ§²ÙJE^    ' EG20 V5.9.0.1yOIðãÀÎzADD

On Error GoTo COPY_ERROR
    iFlag = 0   'Iðt@C³ÆµÄ¨­B
    'Xg{bNX\¦ÌSsÉÂ¢ÄÈºðÀ{·éB
    iMaxLine = lstTraceFile.ListCount  'ÄÚ°½Ì§²ÙØ½ÄÎÞ¯¸½Ìsð¾éB
    
' EG20 V5.9.0.1yOIðãÀÎzADD START
    iFileCounter = 0
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
            iFileCounter = iFileCounter + 1
        End If
    Next

    If iFileCounter > LOG_FILECNT_MAX Then
        ' x¶¾\¦
        MsgBox "Ið³ê½t@CªãÀð´¦Üµ½B" _
               & Chr(vbKeyReturn) & "IðÅ«ét@CÍ[" & LOG_FILECNT_MAX & "]ÜÅÅ·B", _
               vbOKOnly + vbCritical, _
               "t@CwèÙí"
        Exit Sub
    End If
' EG20 V5.9.0.1yOIðãÀÎzADD END
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
        'Ið³ê½sÈçÎA
            If iFlag = 0 Then
                ' æoµæfBNgðIð·é
'                sCopyDir = pfDirSelection("a:", "g[Xt@CÝæÌfBNgIð")  'V1.12.0.1 DEL
                'sCopyDir = pfDirSelection("H:", "g[Xt@CÝæÌfBNgIð")   'V1.12.0.1 ADD@'V1.20.0.1 DEL
                sCopyDir = ShowFolders(Me.hwnd, "tH_ðwèµÄ­¾³¢", SHOWFOLDER_DEFAULTFOLDER) 'V1.20.0.1 ADD
                If sCopyDir = "" Then
                'fBNgwèªÈ¯êÎA ðI¦éB
                    Exit Sub
                End If
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
                'vOXo[ð\¦·é
                Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
            End If
            iFlag = 1  'Iðt@CLèÆ·éB
            'Rs[³t@C¼\¦àeðZbg·éBisWork©ÊÞ²Ä»²½Þ--01@\CTRC2000xxx.xxxj
            sWork = lstTraceFile.List(iLine)
            'æª©çÊÞ²Ä»²½Þ¶i"ÊÞ²Ä»²½Þ--" ·³=LIST_HEDDER_LENGTHjðO·éB
            '                                     isFileName©01@\CTRC2000xxx.xxxj
            sFileName = Right$(sWork, Len(sWork) - LIST_HEDDER_LENGTH)
            'Rs[³t@C¼tpXðZbg·éBisFullPass©C:\tool\R_Mente\DATA\{dS\©w\01@\CTRC2000xxx.xxxj
            sFullPass = PATH_RMENTE_GATE_DEN_JIEKI & sFileName
            'ÝæfBNg{t@CiRs[ ³Æ¯¶j¼ðZbg·éB
            '                                 isCopyFileName©a:\01@\CTRC2000xxx.xxxj
            sCopyFileName = sCopyDir & sFileName
            'Rs[æfBNgÉtH_ðì¬·éB
            On Error Resume Next
            i = 1
            sWork = sCopyDir
            Do
                j = InStr(i, sFileName, "\")
                If j = 0 Then Exit Do
                j = j + 1
                sWork = sWork & Mid$(sFileName, i, j - i)
                MkDir sWork
                i = j
            Loop
            'Og[Xt@CðwèfBNgÉ«o·B
            On Error GoTo COPY_ERROR
            FileCopy sFullPass, sCopyFileName
        End If
    Next
    
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ðÁ·é
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
    
    If iFlag = 0 Then
    't@CªIð³êÄ¢È¯êÎAG[bZ[Wð\¦µAðI¹·éB
        MsgBox "æoµt@CªIð³êÄ¢Ü¹ñB" _
               & Chr(vbKeyReturn) & "IðµÄ­¾³¢B", _
               vbOKOnly + vbExclamation, _
                "[geiXi©®üD@j"
        Exit Sub
    End If
    
    'uØÓ°ÄÒÝÃÅÝ½æÊFt@C}ÌoÍ³ívOoÍ
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, RMENTE_GAMEN_FILE_OUTPUT_OK, 0)
    Exit Sub

COPY_ERROR:
    
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ðÁ·é
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
    
    Select Case Err.Number
        Case 61 ' Rs[æó«eÊs«
            iResponse = MsgBox("ó¯¤ÌhCuÌfBXNª¢ÁÏ¢Å·B" _
               & Chr(vbKeyReturn) & "Vµ¢fBXNð}üµÄ­¾³¢B", _
               vbOKOnly, _
               "Ot@Cæoµ")
        Case 70 ' CgveNg
            lSts = CopyFile(sFullPass, sCopyFileName, 0)
            If (lSts = 0) Then
                iResponse = MsgBox("t@Cðì¬Ü½Íu·Å«Ü¹ñB±ÌfBXNÍCgveNg³êÄÜ·B" _
                   & Chr(vbKeyReturn) & "CgveNgðð·é©@ÊÌfBXNðgÁÄ­¾³¢B", _
                   vbOKOnly, _
                   "Ot@Cæoµ")
            End If
        Case 71 ' fBXNð}üµÄ­¾³¢
            iResponse = MsgBox("hCuÉfBXNªüÁÄÜ¹ñB" _
               & Chr(vbKeyReturn) & "fBXNð}üµÄ©çâè¼µÄ­¾³¢B", _
               vbOKOnly, _
               "Ot@Cæoµ")
         Case Else
            iResponse = MsgBox("\ú¹ÊG[ª­¶µÜµ½B" _
               & Chr(vbKeyReturn) & "ìðâè¼µÄ­¾³¢B", _
               vbOKOnly, _
               "Ot@Cæoµ")
    End Select
    
    'uØÓ°ÄÒÝÃÅÝ½æÊFt@C}ÌoÍÙívOoÍ
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, RMENTE_GAMEN_FILE_OUTPUT_ERROR, lngErrCode)

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : sLzhFileWrite
'//  @\¼Ì  : u³k}ÌoÍvtº
'//  @\Tv  : Xg{bNXÅwè³ê½t@Cð³kµA
'//              wèfBNgÉoÍ·éB
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 g[Xt@C³kÝæfBNgÊuÏX
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 tH_IðæÊðOSdlÉÏX
'//     REVISIONS :(EG20V5.9.0.1) 2012-05-03  REVISED BY [TCC] M.Chiwaki
'//                 O}ÌoÍAãÀðTPQÆ·é
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 yvOXo[\¦@\©¼µÎz
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub sLzhFileWrite()
    Dim iLine As Integer         'ÄÚ°½Ì§²ÙØ½ÄÎÞ¯¸½Ìs²ÝÃÞ¯¸½
    Dim iMaxLine As Integer      'ÄÚ°½Ì§²ÙØ½ÄÎÞ¯¸½Ìs
    Dim iFlag As Integer         'Iðt@CL³iP^Oj
    Dim iResponse As Integer     'MsgBox{^R[h
    Dim sFullPass As String      '³k³t@CtpX¼
    Dim sFileName As String      '³k³t@C¼
    Dim sLzhDirName As String    '.LZHÌ§²Ùi[fBNg¼
    Dim sLzhFileName As String   '.LZHÌ§²Ù¼
    Dim iSts As Integer          'Ößèl
    Dim sWork As String          '[N
    Dim i As Integer             '[N
    Dim j As Integer             '[N
    Dim lngErrCode As Long       'G[R[h
    Dim nIndex As Integer        ' ¶                    ' EG20 V5.6.0.1ÇÁ
    Dim iFileCounter As Integer  'ÎÛÌ§²ÙJE^    ' EG20 V5.9.0.1yOIðãÀÎzADD
    
On Error GoTo WRITE_ERROR
    iFlag = 0   'Iðt@C³ÆµÄ¨­B
    'Xg{bNX\¦ÌSsÉÂ¢ÄÈºðÀ{·éB
    iMaxLine = lstTraceFile.ListCount  'ÄÚ°½Ì§²ÙØ½ÄÎÞ¯¸½Ìsð¾éB
    
' EG20 V5.9.0.1yOIðãÀÎzADD START
    iFileCounter = 0
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
            iFileCounter = iFileCounter + 1
        End If
    Next

    If iFileCounter > LOG_FILECNT_MAX Then
        ' x¶¾\¦
        MsgBox "Ið³ê½t@CªãÀð´¦Üµ½B" _
               & Chr(vbKeyReturn) & "IðÅ«ét@CÍ[" & LOG_FILECNT_MAX & "]ÜÅÅ·B", _
               vbOKOnly + vbCritical, _
               "t@CwèÙí"
        Exit Sub
    End If
' EG20 V5.9.0.1yOIðãÀÎzADD END
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
        'Ið³ê½sÈçÎA
            If iFlag = 0 Then
                ' æoµæfBNgðIð·é
'                sLzhDirName = pfDirSelection("a:", "g[Xt@C³kÝæÌfBNgIð")   'V1.12.0.1 DEL
                'sLzhDirName = pfDirSelection("H:", "g[Xt@C³kÝæÌfBNgIð")    'V1.12.0.1 ADD 'V1.20.0.1 DEL
                sLzhDirName = ShowFolders(Me.hwnd, "tH_ðwèµÄ­¾³¢", SHOWFOLDER_DEFAULTFOLDER)  'V1.20.0.1 ADD
                If sLzhDirName = "" Then
                'fBNgwèªÈ¯êÎA ðI¦éB
                    Exit Sub
                End If
' EG20 V5.6.0.1y³ktH_wèÎzÇÁJn
                ' oÍtH_É¼pXy[XªÜÜêÄ¢éêA³kÅÙíª­¶µÄµÜ¤½ß
                ' ³kOÉ`FbNµÄÙíð\¦·éB
                nIndex = InStr(sLzhDirName, " ")
                If nIndex <> 0 Then
                    ' x|bvAbvEBhEð\¦·éB
                    Call MsgBox(CABFOLDERSELECT_ERRORMESSAGE, vbCritical, CABFOLDERSELECT_ERRORTITLE)
                    Exit Sub  'fBNgªwè³êÈ¯êÎAI¹
                End If
' EG20 V5.6.0.1y³ktH_wèÎzÇÁI¹

' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
                'vOXo[ð\¦·é
                Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
            
            End If
            iFlag = 1  'Iðt@CLèÆ·éB
            '³k³t@C¼\¦àeðZbg·éBisWork©ÊÞ²Ä»²½Þ--01@\CTRC2000xxx.xxxj
            sWork = lstTraceFile.List(iLine)
            'æª©çÊÞ²Ä»²½Þ¶i"ÊÞ²Ä»²½Þ--" ·³=LIST_HEDDER_LENGTHjðO·éB
            '                                  isFileName©01@\CTRC2000xxx.xxxj
            sFileName = Right$(sWork, Len(sWork) - LIST_HEDDER_LENGTH)
            '³k³t@C¼tpXðZbg·éBisFullPass©C:\tool\R_Mente\DATA\{dS\©w\01@\CTRC2000xxx.xxxj
            sFullPass = PATH_RMENTE_GATE_DEN_JIEKI & sFileName
            'ÝæfBNg{t@Ci³k³Æ¯¶j¼ðZbgµAg£qÉA.CABðtÁ·éB
            '                                 isLzhFileName©a:\01@\CTRC2000xxx.xxx.CABj
            sLzhFileName = sLzhDirName & sFileName & ".CAB"
            '³kæfBNgÉtH_ðì¬·éB
            On Error Resume Next
            i = 1
            sWork = sLzhDirName
            Do
                j = InStr(i, sFileName, "\")
                If j = 0 Then Exit Do
                j = j + 1
                sWork = sWork & Mid$(sFileName, i, j - i)
                MkDir sWork
                i = j
            Loop
            On Error GoTo WRITE_ERROR
            'ÎÛt@CðA³kµ.CABt@CÉi[·éB
            Call psCabReqest(CABREQEST.CAB_COMPRESSION, sLzhFileName, sFullPass)
        End If
    Next
    
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ðÁ·é
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
    
    If iFlag = 0 Then
    't@CªIð³êÄ¢È¯êÎAG[bZ[Wð\¦µAðI¹·éB
        MsgBox "æoµt@CªIð³êÄ¢Ü¹ñB" _
               & Chr(vbKeyReturn) & "IðµÄ­¾³¢B", _
               vbOKOnly + vbExclamation, _
               "[geiXi©®üD@j"
        Exit Sub
    End If
    
    'uØÓ°ÄÒÝÃÅÝ½æÊF³k}ÌoÍ³ívOoÍ
    Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, RMENTE_GAMEN_LZH_OUTPUT_OK, 0)
  
    Exit Sub

WRITE_ERROR:
    'uØÓ°ÄÒÝÃÅÝ½æÊF³k}ÌoÍÙívOoÍ
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, RMENTE_GAMEN_LZH_OUTPUT_ERROR, lngErrCode)

' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ðÁ·é
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : sLzhFileDisplay
'//  @\¼Ì  : u³kÊmFvtº
'//  @\Tv  : wè³ê½³kt@CÌàeðæ¾µA \¦·éB
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(1.12.0.1) 2009-11-11   REVISED BY [TCC] C.Terui
'//                 ³kt@CIðæfBNgÊuÏX
'//     REVISIONS :(1.20.0.1) 2010-03-11  REVISED BY [TCC] S.Yoshimori
'//                 t@CIðæÊðOSdlÉÏX
'//     REVISIONS :(2.5.0.1) 2010-10-29  REVISED BY [TCC] S.Terao
'//                 }ÌæOsïC³
'//     REVISIONS :(EG20 V5.13.0.1) 2012-06-02 REVISED BY  [TCC] H.Sugimoto
'//                 yvOXo[\¦@\©¼µÎz
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub sLzhFileDisplay()
    Dim sLzhFileName As String   '.LZHÌ§²Ù¼
    Dim sLzhDataFile As String   '.LZHÌ§²ÙàeÝt@C¼iÌÙÊß½j
    Dim sCommand As String
    Dim lRetVal As Long
    
    Dim objFso As New FileSystemObject   't@CVXeIuWFNg  'V1.20.0.1 ADD
    
    On Error Resume Next

    '³kt@CIðæÊð\¦µA³kt@CðIð³¹éB
'    sLzhFileName = pfCabFileSelection("a:")        'V1.12.0.1 DEL
    'sLzhFileName = pfCabFileSelection("H:")         'V1.12.0.1 ADD 'V1.20.0.1 DEL
    'V1.20.0.1 ADD START
    'æ¾t@C¼ðú»
    CommonDialog1.FileName = ""
    'úfBNgðÝè
    If objFso.FolderExists(SHOWFILE_DEFAULTFOLDER1) = True Then    'tH_IðæÊftHgpXPª¶Ý·é©
        '¶Ý·é½ßAftHgpXPiH:jðÝè
        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER1
    Else
        '¶ÝµÈ¢½ßAftHgpXQiC:jðÝè
        CommonDialog1.InitDir = SHOWFILE_DEFAULTFOLDER2
    End If
    Set objFso = Nothing
    'g£qðÝè
    CommonDialog1.Filter = "³kt@Ci*.cabj|*.cab|"
    't@CIðæÊðJ­
    CommonDialog1.ShowOpen
    'Iðµ½t@C¼ðæ¾
    sLzhFileName = CommonDialog1.FileName
    'V1.20.0.1 ADD END
    If sLzhFileName = "" Then Exit Sub   't@CªIð³êÈ¯êÎAßéB
    
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ð\¦·é
    Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_HOSHUKINOU)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
    
    'Ið³ê½³kt@CÌàeðæ¾·éB
    Call psCabReqest(CABREQEST.CAB_DRAFT, sLzhFileName, vbNullString)
    
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁJn
    'vOXo[ðÁ·é
    Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ)
' EG20 V5.13.0.1yvOXo[\¦@\©¼µÎzÇÁI¹
    
    't@Càeæ¾lãü
    sLzhDataFile = gstrCabErrCd
    If sLzhDataFile = "" Then Exit Sub   't@Càeæ¾G[Å êÎAßéB
    ' ÌÀsR}hðì¬·é
    sCommand = MN_EXE_MEMO & sLzhDataFile
    lRetVal = Shell(sCommand, vbMaximizedFocus)
    ' ðANeBuiOÊ\¦jÉ·é
    AppActivate lRetVal, True
    SendKeys "{LEFT}", True
    
    Call ChDrive("D")  'V2.5.0.1 ADD
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : fSelectedFilesDelete
'//  @\¼Ì  : ut@Cívtº
'//  @\Tv  : IðÌt@Cðí·éB
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Boolean@@@@@@[OUT]ßèl
'//                                   True:t@Cí@FALSEFt@C¢í
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     ORIGINAL  :(1.1.0.2) 2009-02-XX   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Function fSelectedFilesDelete() As Boolean
    Dim iLine As Integer         'ÄÚ°½Ì§²ÙØ½ÄÎÞ¯¸½Ìs²ÝÃÞ¯¸½
    Dim iMaxLine As Integer      'ÄÚ°½Ì§²ÙØ½ÄÎÞ¯¸½Ìs
    Dim iDelLine As Integer      'ÄÚ°½Ì§²ÙØ½ÄÎÞ¯¸½ÌIðs
    Dim iResponse As Integer     'MsgBox{^R[h
    Dim sFullPass As String      'íÎÛt@CtpX¼
    Dim sFileName As String      'íÎÛt@C¼
    Dim sWork As String          '[N

On Error GoTo ErrorDeleteFile
    
    't@CíÈµÆµÄ¨­B
    fSelectedFilesDelete = False
    iDelLine = 0
    'Xg{bNX\¦ÌSsÉÂ¢ÄÈºðÀ{·éB
    iMaxLine = lstTraceFile.ListCount  'ÄÚ°½Ì§²ÙØ½ÄÎÞ¯¸½Ìsð¾éB
    For iLine = CNT_MIN To iMaxLine - 1
        If lstTraceFile.Selected(iLine) = True Then
        'Ið³ê½sÈçÎA
            If iDelLine = 0 Then
                'ímFbZ[Wð\¦·éB
                iResponse = MsgBox("IðÌt@CðíµÜ·B" _
                                    & Chr(vbKeyReturn) & " æëµ¢Å·©H", _
                                    vbYesNo + vbExclamation, _
                                    "g[Xt@CÌí")
                If iResponse = vbNo Then
                ' [¢¢¦] {^ðIðµ½êAí¹¸I¹·éB
                    Exit Function
                End If
            End If
            'Yst@C¼\¦àeðZbg·éBisWork©ÊÞ²Ä»²½Þ--01@\CTRC2000xxx.xxxj
            sWork = lstTraceFile.List(iLine)
            'æª©çÊÞ²Ä»²½Þ¶i"ÊÞ²Ä»²½Þ--" ·³=LIST_HEDDER_LENGTHjðO·éB
            '                                   isFileName©01@\CTRC2000xxx.xxxj
            sFileName = Right$(sWork, Len(sWork) - LIST_HEDDER_LENGTH)
            'Rs[³t@C¼tpXðZbg·éBisFullPass©:\tool\R_Mente\DATA\{dS\©w\01@\CTRC2000xxx.xxxj
            sFullPass = PATH_RMENTE_GATE_DEN_JIEKI & sFileName
            'YsÌt@Cðí·éB
            Kill sFullPass
            iDelLine = iDelLine + 1
            't@Cðíµ½B
            fSelectedFilesDelete = True
            'u©®üD@ØÓ°ÄÒÝÃÅÝ½æÊFt@CívOoÍ
            Call sLogTraceReq(LTYP_NORMAL, L3AN_FILE, FILE_DELETE, 0)
        End If
    Next
Exit Function

ErrorDeleteFile:

    MsgBox "t@CÌíÅG[ª­¶µÜµ½B", _
           vbOKOnly + vbExclamation, _
           "g[Xt@CÌí"

    fSelectedFilesDelete = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : tmrMail_Timer
'//  @\¼Ì  : [óMp^C}A^CAbv
'//  @\Tv  : [óMðs¤B
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(EG20 V8.1.0.1) 2014-06-05  REVISED BY  [TCC] S.Kuroda
'//                 2014Nx{ô yEG20_KANSI05_01z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
On Error Resume Next
    'Äp[óMðs¤
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
        AppActivate frmRMente.Caption, False
        pfFormActive (frmRMente.hwnd)           ' EG20 V8.1.0.1yEG20_KANSI05_01zADD
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  Ö¼Ì  : psGATERMenteTool
'//  @\¼Ì  : ©®üD@Ì[geiXc[pXðæ¾
'//  @\Tv  : ©®üD@[geiXc[pXðæ¾ðs¤B
'//
'//              ^        ¼Ì      Ó¡
'//  ø      : Èµ
'//
'//              ^        l        Ó¡
'//  ßèl    : Èµ
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  õlF
'///////////////////////////////////////////////////////////////////
Public Sub psGATERMenteTool()
 
    Dim sPath As String * MAX_PATH_SIZE
    Dim iRet As Integer
    
    On Error Resume Next

    ' HOSHU.INIæè©®üD@c[pXðæ¾·éB
    iRet = GetPrivateProfileString(KANSI_HOSHU_GATE_RMENTE_SEC, _
                                    KANSI_HOSHU_GATE_RMENTE_KEY, _
                                    DEFAILT, sPath, Len(sPath), _
                                    HOSHU_FILE)

      If iRet = 0 Then
        sTOOLPass = ""
      Else
        sTOOLPass = sPath
      End If
      
End Sub


