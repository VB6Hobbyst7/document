VERSION 5.00
Begin VB.Form frmKikiSettei 
   BorderStyle     =   0  'Θ΅
   Caption         =   "@νξρέθ"
   ClientHeight    =   9000
   ClientLeft      =   2175
   ClientTop       =   2430
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'Z ΅°ΐή°
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Μωθl
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "w±@νIDmF"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   6360
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Left            =   240
      Top             =   8280
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "Wυέθ ΫΆ^³"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "LANJ[hέθ"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   6360
      TabIndex        =   4
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "wsxf[^mF"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   6360
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "@ν\¬έθ"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cmdFixedExe 
      Caption         =   "wsxf[^έθ"
      BeginProperty Font 
         Name            =   "lr SVbN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   " eiX   ζΚΦίι"
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
      Left            =   9720
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  '΅¦
      BackColor       =   &H00800000&
      Caption         =   "@νξρέθ"
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
      TabIndex        =   6
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmKikiSettei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 ALL Rights Reserved
'//
'//  t@CΌ  F@νξρέθj[ζΚ.frm
'//  pbP[WΌF@νξρέθj[ΜtH[W[
'//
'//  TvFpX[hόΝζΚ
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 tF[YQΞ@w±@νIDmFζΚALANJ[hγΚΊΚέθζΚΗΑ
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 tΜΊΒ^sΒΗΑ
'//                 ζΚbN^ζΚbNπΗΑ
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                ζΚΔOΚ\¦C³(sοC³)
'//  υlF
'///////////////////////////////////////////////////////////////////
Option Explicit
Private Const MN_MAIL_INTERVAL = 1000       'C^C}ΜC^[ol

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  ΦΌΜ  : Form_Activate
'//  @\ΌΜ  : @νξρέθj[ζΚ(ANeBuFCxgvV[W)
'//  @\Tv  : [σM^C}N?
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()

    'G[[`πιΎ
    On Error Resume Next
    
    '^C}πN?·ι
    tmrMail.Enabled = True

End Sub

'EG20 V2.1.0.1 ADD START ytF[YQΞz
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2011 All Rights Reserved
'//
'//  ΦΌΜ  : Form_Deactivate
'//  @\ΌΜ  : @νξρέθj[ζΚ(fBANeBu)
'//  @\Tv  : [σMpA^C}β~
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(EG20 V2.1.0.1) 2011-11-26   CODED   BY [TCC] M.Matsumoto
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
On Error Resume Next
    '^C}πβ~·ι
    tmrMail.Enabled = False
End Sub
'EG20 V2.1.0.1 ADD END


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  ΦΌΜ  : Form_Load
'//  @\ΌΜ  : @νξρέθj[ζΚ([hFCxgvV[W)
'//  @\Tv  : ϊπs€B
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 tF[YQΞ
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    
    'G[[`πιΎ
    On Error Resume Next
    
    'ζΚμOoΝ
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_START, 0)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
    'V1.4.0.1@ADD START
    'IDUkή`FbN
    psIDUCheck
    
    If pbIDUSts = 1 Then
     'w±@νIDmFρ\¦
      cmdFixedExe(5).Visible = False
    End If
    'V1.4.0.1@ADD END
    
    'CσMpΜ^C}lπέθ·ι
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  ΦΌΜ  : cmdFixedExe_Click
'//  @\ΌΜ  : etΊ
'//  @\Tv  : ©ζΚπΑ·ιB
'//
'//              ^        ΌΜ     @@@Σ‘
'//  ψ      : Integer@ Index          IπtΜCfbNX
'//
'//              ^        l        @@ Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.4.0.1) 2009-03-23   REVISED BY [TCC] S.Terao
'//                 tF[YQΞ
'//                 Ewsxf[^mFj[ζΚ¨wsxf[^mF(wξρ)ζΚ
'//                 E@ν\¬έθj[ζΚ¨@ν\¬έθ(wξρ)ζΚ
'//                 ELANJ[hγΚΊΚέθtΊAζΚ\¦ΗΑ
'//                 Ew±@νIDmFζΚΗΑ
'//     REVISIONS :(1.12.0.1) 2009-11-10  REVISED BY [TCC] C.Terui
'//                 tΜΊΒ^sΒΗΑ
'//
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub cmdFixedExe_Click(Index As Integer)

    'G[[`πιΎ
    On Error Resume Next
    
'V1.12.0.1 ADD START
    'S{^πΊsΒΖ·ιB
    Call SetEnableFalse
'V1.12.0.1 ADD END
    
    Select Case Index
        
        Case 0                                 'wsxf[^έθ
            'ζΚμOoΝ
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_EKITUDO_DATA_SETTEI, 0)
            
            'ζΚ\¦
            Load frmEkisettei
            frmEkisettei.Show 1

        Case 1                                 'wsxf[^mF
            'ζΚμOoΝ
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_EKITUDO_DATA_KAKUNIN, 0)
            'V1.4.0.1 DEL START
            'ζΚ\¦
            'Load frmEkiDataGateMenu
            'frmEkiDataGateMenu.Show 1
            'V1.4.0.1 DEL END
            'V1.4.0.1 ADD START
            'ζΚ\¦
            Load frmEkiData
            frmEkiData.Show 1
            'V1.4.0.1 ADD END
   
        Case 2                                 '@ν\¬έθ
            'ζΚμOoΝ
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_KIKI_KOUSEI_SETTEI, 0)
            'V1.4.0.1 DEL START
            'ζΚ\¦
            'Load frmKikiDataMenu
            'frmKikiDataMenu.Show 1
            'V1.4.0.1 DEL END
            'V1.4.0.1 ADD START
            'ζΚ\¦
            Load frmKikiData
            frmKikiData.Show 1
            'V1.4.0.1 ADD END
            
        Case 3                                 'LANJ[hγΚΊΚέθ
            'ζΚμOoΝ
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_LAN_CARD_SETTEI, 0)

'V1.4.0.1 ADD START
            'ζΚ\¦
            Load frmLanSettei
            frmLanSettei.Show 1
'V1.4.0.1 ADD END

        Case 4                                 'WυέθΫ³
            'ζΚμOoΝ
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_KAKARI_SAVE_RESTORE, 0)
            
            'ζΚ\¦
            Load frmRenewData
            frmRenewData.Show 1
'V1.4.0.1 ADD START
        Case 5                                 'w±@νID
            'ζΚμOoΝ
            Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_EKIMUKIKI_ID, 0)
            'ζΚ\¦
            Load frmEkimKikiId
            frmEkimKikiId.Show 1
'V1.4.0.1 ADD END

        Case Else
            'Θ΅
            
    End Select

'V1.12.0.1 ADD START
    'S{^πΊΒΖ·ιB
    Call SetEnableTrue
'V1.12.0.1 ADD END
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  ΦΌΜ  : cmdReturn_Click
'//  @\ΌΜ  : ueiXζΚΦίιvtΊ
'//  @\Tv  : ©ζΚπΑ·ιB
'//
'//              ^        ΌΜ     @@@Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        @@ Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()

    'G[[`πιΎ
    On Error Resume Next
    
    'ζΚμOoΝ
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, KIKIINFSETMENU_GAMEN_END, 0)
    
    '©ζΚΑ
    Unload Me

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  ΦΌΜ  : tmrMail_Timer
'//  @\ΌΜ  : [σMp^C}i^CAbvFCxgvV[Wj
'//  @\Tv  : ΔpCσMπs€
'//
'//              ^        ΌΜ     @@@Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        @@ Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(1.1.0.1) 2008-12-01   CODED   BY [TCC] E.Watanabe
'//     REVISIONS :(1.17.0.1) 2009-01-05   REVISED BY [TCC] S.Terao
'//                ζΚΔOΚ\¦C³(sοC³)
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()
    
    'G[[`πιΎ
    On Error Resume Next
    
    'ΔpCσMπs€
    If pfComMailRecieve = ML_ID_HOSHU_ACTIVE_REQ Then
       'V1.17.0.1 DEL START
'        AppActivate frmRenewData.Caption, False
'        pfFormActive (frmRenewData.hwnd)
       'V1.17.0.1 DEL START
       'V1.17.0.1 ADD START
        AppActivate frmKikiSettei.Caption, False
        pfFormActive (frmKikiSettei.hwnd)
       'V1.17.0.1 ADD END
    End If

End Sub

'V1.12.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2009 All Rights Reserved
'//
'//  ΦΌΜ  : SetEnableFalse
'//  @\ΌΜ  : ζΚbN
'//  @\Tv  : ζΚπbN·ιB
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(1.12.0.1) 2009-11-10   CODED   BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υl F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableFalse()
    
    'G[[`πιΎ
    On Error Resume Next

    'S{^πΊsΒΖ·ιB
    cmdFixedExe(0).Enabled = False
    cmdFixedExe(1).Enabled = False
    cmdFixedExe(2).Enabled = False
    cmdFixedExe(3).Enabled = False
    cmdFixedExe(4).Enabled = False
    cmdFixedExe(5).Enabled = False
    cmdReturn.Enabled = False
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2009 All Rights Reserved
'//
'//  ΦΌΜ  : SetEnableTrue
'//  @\ΌΜ  : ζΚbNπ
'//  @\Tv  : ζΚΜbNππ·ιB
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(1.12.0.1) 2009-11-10   CODED   BY [TCC] C.Terui
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υl F
'///////////////////////////////////////////////////////////////////
Public Sub SetEnableTrue()
    
    'G[[`πιΎ
    On Error Resume Next

    'S{^πΊΒΖ·ιB
    cmdFixedExe(0).Enabled = True
    cmdFixedExe(1).Enabled = True
    cmdFixedExe(2).Enabled = True
    cmdFixedExe(3).Enabled = True
    cmdFixedExe(4).Enabled = True
    cmdFixedExe(5).Enabled = True
    cmdReturn.Enabled = True
    
End Sub
'V1.12.0.1 ADD END

