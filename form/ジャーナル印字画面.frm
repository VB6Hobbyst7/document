VERSION 5.00
Begin VB.Form frmJprPrint 
   BorderStyle     =   0  'Θ΅
   Caption         =   "W[iσ"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkJprKind 
      Caption         =   "έθlκ"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   32
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Frame FraPrintKind 
      Caption         =   "σΪwθ"
      Height          =   1815
      Left            =   120
      TabIndex        =   29
      Top             =   2760
      Width           =   11655
      Begin VB.CheckBox chkJprKind 
         Caption         =   "όD@Ϋηέθf[^"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   39
         Top             =   1440
         Width           =   5055
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "wsxf[^mF(΄έΊ°ΔήΊ°Ε@ξρθ`)"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   38
         Top             =   1080
         Width           =   5055
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "w±@νhc"
         Height          =   255
         Index           =   7
         Left            =   8520
         TabIndex        =   37
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "χΨItCoΝ"
         Height          =   255
         Index           =   6
         Left            =   8520
         TabIndex        =   36
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "?­o[Wκ"
         Height          =   255
         Index           =   5
         Left            =   8520
         TabIndex        =   35
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "pΰzf[^"
         Height          =   255
         Index           =   4
         Left            =   5400
         TabIndex        =   34
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "Κίf[^"
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   33
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "wsxf[^mF(©ό)"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   31
         Top             =   720
         Width           =   3255
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "wsxf[^mF(wξρ)"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.CheckBox chkGouki 
      Caption         =   "PT@"
      Height          =   375
      Index           =   14
      Left            =   10080
      TabIndex        =   27
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox chkGouki 
      Caption         =   "PS@"
      Height          =   375
      Index           =   13
      Left            =   10080
      TabIndex        =   26
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chkGouki 
      Caption         =   "V@"
      Height          =   375
      Index           =   6
      Left            =   6960
      TabIndex        =   19
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame FraGouki 
      Caption         =   "@wθ"
      Height          =   1935
      Left            =   4800
      TabIndex        =   12
      Top             =   600
      Width           =   6975
      Begin VB.CheckBox chkGouki 
         Caption         =   "PU@"
         Height          =   375
         Index           =   15
         Left            =   5280
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "PR@"
         Height          =   375
         Index           =   12
         Left            =   5280
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "PQ@"
         Height          =   375
         Index           =   11
         Left            =   3600
         TabIndex        =   24
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "PP@"
         Height          =   375
         Index           =   10
         Left            =   3600
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "PO@"
         Height          =   375
         Index           =   9
         Left            =   3600
         TabIndex        =   22
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "X@"
         Height          =   375
         Index           =   8
         Left            =   3600
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "W@"
         Height          =   375
         Index           =   7
         Left            =   2160
         TabIndex        =   20
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "U@"
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "T@"
         Height          =   375
         Index           =   4
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "S@"
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "R@"
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "Q@"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "P@"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FraCorner 
      Caption         =   "R[iwθ"
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4455
      Begin VB.CheckBox chkCorner 
         Caption         =   "R[iU"
         Height          =   225
         Index           =   5
         Left            =   2640
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "R[iT"
         Height          =   225
         Index           =   4
         Left            =   2640
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "R[iS"
         Height          =   225
         Index           =   3
         Left            =   2640
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "R[iR"
         Height          =   225
         Index           =   2
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "R[iQ"
         Height          =   225
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "R[iP"
         Height          =   225
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "σ"
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
      Left            =   120
      TabIndex        =   4
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Timer tmrMail 
      Left            =   9600
      Top             =   360
   End
   Begin VB.ListBox LstStatus 
      Height          =   2310
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   11655
   End
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "lr oSVbN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   15000
      Width           =   2895
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "f[^ϋWEoΝ  ζΚΦίι"
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
      Left            =   9480
      TabIndex        =   1
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  '΅¦
      BackColor       =   &H00800000&
      Caption         =   "W[iσ"
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
      TabIndex        =   3
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmJprPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 ALL Rights Reserved
'//
'//  t@CΌ  FfrmJprPrint.frm
'//  pbP[WΌFW[iσζΚ
'/
'//  TvFVXeϊ»(ΔΥ)ζΚ
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-18   CODED   BY [TCC] T.Takajima
'//     REVISIONS :(EG20 V7.4.0.1) 2013-07-22   CODED   BY [TCC] T.Nakajima
'//                 ϊά½ͺθoκt[έθζΚΞ
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED   BY [TCC] T.Nakajima
'//                 k€V²όtF[YQΞ
'//                 yHKRK_Kansi07_003_01z SUB_GATE_KAN.INItH[}bg©Ό΅Ξ
'//     REVISIONS :(EG30 V32.1.0.1) 2016-06-10  CODED   BY [TCC] T.Nakajima
'//                 2016Nx{τΞ
'//     REVISIONS :(EG30 V32.2.0.1) 2016-07-20  CODED   BY [TCC] T.Nakajima
'//                 2016Nx{τΞ vOXo[ρ\¦Ξ
'//     REVISIONS :(EG30 V35.3.0.1) 2019-07-03   REVISED BY [TCC] H.Kondoh
'//                2019Nx{τΞ
'//     REVISIONS :
'//
'//  υlF
'///////////////////////////////////////////////////////////////////
Option Explicit

'ϊ»ΐstO
Private bSysFormat As Boolean

Private Const APL_INTERVAL = 390000     'AvN?^C}ftHgl
Public glbFilePath  As String             't@CpX     'V1.12.0.1 ADD
Dim lngMAX_Time As Long                    'INIζΎέθl
Dim lngtime     As Long                    '»έ^C}l
Private iSendType As Integer            'vνΚl
'V1.5.0.1 ADD END
'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   '[^C}ΜC^[ol
'V1.20.0.1 ADD START
Private Const LOG_INTERVAL = 30000        'ON?^C}ftHgl(30b)
Dim lngLogMAX_Time As Long                'INIζΎέθl(Oj
'V1.20.0.1 ADD END
Dim intJprFile        As Integer        'EG20 V30.1.0.1 ADD


' W[ioΝέθξρ
Private Type JPR_PRINT_SETTING_INFO
    iCornerCount        As Integer          ' `FbN³κ½R[i
    iCorner(5)          As Integer          ' `FbN³κ½R[iκ
    iGoukiCount         As Integer          ' `FbN³κ½@
    iGouki(15)          As Integer          ' `FbN³κ½@κ
    iJprCount           As Integer          ' `FbN³κ½W[iνή
'    iJprKind(7)         As Integer          ' `FbN³κ½W[iκ      'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z DEL
'    iJprKind(8)         As Integer          ' `FbN³κ½W[iκ      'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z ADD    'EG30 V32.1.0.1 DEL
    iJprKind(9)         As Integer          ' `FbN³κ½W[iκ      'EG30 V32.1.0.1 ADD
End Type
Private Enum JPR_KIND
    JPR_KIND_EKI_INFO = 0           ' wsxf[^mF(wξρ)
    JPR_KIND_JIKAI_INFO = 1         ' wsxf[^mF(©ό)
    JPR_KIND_SETTING_LST = 2        ' έθlκ
    JPR_KIND_TUKA_DATA = 3          ' Κίf[^
    JPR_KIND_RIYO_KINGAKU = 4       ' pΰzf[^
    JPR_KIND_KADO_VER = 5           ' ??o[Wκ
    JPR_KIND_SIMEKIRI = 6           ' χΨItCoΝ
    JPR_KIND_EKIMU_ID = 7           ' w±@νID
    JPR_KIND_SUBGATE_INFO = 8       ' wsxf[^mF(GR[hR[i@ξρθ`)    'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z ADD
    JPR_KIND_GATE_CFG = 9          ' όD@Ϋηέθf[^  'EG30 V32.1.0.1 ADD
End Enum
Dim udtJprPrintSetteingInfo    As JPR_PRINT_SETTING_INFO
Dim udtInitJprSetting           As JPR_PRINT_SETTING_INFO
Dim iJprIdx                     As Integer          'ΜW[i

'@ν\¬f[^iwξρjC[Wt@CΗζpΜ\’Μ
Private Type EKIINFO_IMAGE_FILE
    sType       As String                'νΚ
    sGoki       As String                '@
    sNo         As String                'νΚΚΤ
    sCorner     As String                'R[i        ' EG20 V2.1.0.1[Mainte_03_01 wsxΞ]ΗΑ
    sTuuban     As String                'ΚΤ
    sKoumoku    As String                'Ϊ
    sKubun      As String                'ζͺ
    sSettei     As String                'έθl
    sSyosai     As String                'έθlΪΧ
End Type

'wsxf[^iόD@)C[Wt@CΗέζθpΜ\’Μ
Private Type JIKAIINFO_IMAGE_FILE
    strBunrui_Dai  As String               'εͺή
    strBunrui_Tyu  As String               'ͺή
    srtBunrui_Sho   As String               '¬ͺή
    strCorner       As String               'R[i
    strKomoku       As String               'Ϊ
    strKubun        As String               'ζͺ
    strData         As String               'f[^
    strSetShosai    As String               'ΪΧ
    
End Type

'wsxf[^mF(©ό)W[ioΝt@Cμ¬e[u
Private Type JIKAI_JPREDIT_TBL
    strKomoku       As String               'ΪΌi{ζͺ)
    strBunrui_Sho   As String               '»ΜΪπw·¬ͺήR[h
    strKubun        As String               '»ΜΪπw·ζͺ
End Type

'??o[WoΝζͺ
Private Enum mintDispDiv
    KADOVER_FILE_DISP = 0
    KADOVER_FILE_OUTPUT
End Enum

'}ΜoΝt@CΗέζθpΜ\’Μ(Κί/pΰz)
Private Type BAITAI_OUTPUT_IMAGE_FILE
    strKomokuName       As String          'ΪΌ
    strGoukei           As String          'Κίv
    srtGoukiValue(15)   As String          '@ΚΜl(’gp)
End Type

'EG20 V30.1.0.1 ADD START
'}ΜoΝt@CΗέζθpΜ\’Μ(Κί/pΰz)y²όpz
Private Type BAITAI_OUTPUT_IMAGE_FILE_KAN
    strKomokuName       As String          'ΪΌ
    strGoukei           As String          'Κίv
    strNorikae          As String          'Κίζ·(’gp)
    strTukaChoku        As String          'ΚίΌΪ(’gp)
    srtGoukiValue(31)   As String          '@ΚΜl(’gp)
End Type
'EG20 V30.1.0.1 ADD END

'έθκt@CΗέζθp\’Μ(OPERATE_SET##.CSVj
Private Type SETTEI_OUTPUT_IMAGE_FILE
    strDaiKomoku        As String           'εΪΌ
    strKomoku           As String           'ΪΌ
    strValue            As String           'έθl
    strChangeFlg        As String           'ΟXtO 'EG30 V32.1.0.1 ADD
End Type

'?­o[Wt@CΗέζθp\’Μ(KadoVerDisp.csv)
Private Type KADO_VER_DISP_IMAGE_FILE
    strKishu            As String           '@νͺήit@CΗέέpj
    strCorner           As String           'R[iͺήit@CΗέέpj
    strGokiDiv          As String           '@ͺήit@CΗέέpj
    strName             As String           '@νΌit@CΗέέpj
    strMaker            As String           '[JΌit@CΗέέpj
    strVer              As String           'o[Wit@CΗέέpj
    strDate             As String           'μ¬ϊtit@CΗέέpj
End Type

'EG30 V32.1.0.1 ADD START
'όD@Ϋηέθf[^ ΗέζθpΜ\’Μ(JP_CFGR[i@Τ.csv)
Private Type GATE_CFG_DATA_FILE
    strInfoName         As String           'ξρΌ
    strBunrui_Dai       As String           'εΪ
    strBunrui_Chu       As String           'ͺή
    strBunrui_Syo       As String           '¬ͺή
    strValue            As String           'έθl
    strChangeFlg        As String           'ΟXL³tO
End Type
'EG30 V32.1.0.1 ADD END


'W[i?WΤt@C
Private Const EKIMU_DEFU = "APL\APL_WORK"
Private Const EDIT_DATA_EKIINFO = PATH_WORK & "EKI_DISP_EKIINFO.csv"    'wsxf[^mF(wξρ)"
Private Const EDIT_DATA_JIKAIINFO = PATH_WORK & "EKI_DISP_GATE_JPR.csv" 'wsxf[^mF(©ό)"
Private Const EDIT_DATA_SETTEI = PATH_WORK & "OPERATE_SET##.csv"        'έθlκ
Private Const EDIT_DATA_KADOVERSION = PATH_WORK & "KadoVerDisp####"     '??o[WκiKadoVerDispR[iΤA@Τj
Private Const EDIT_DATA_SIMEKIRI = PATH_WORK & "SIME##.txt"             'χΨItCoΝ
Private Const EDIT_DATA_EKIMUID = PATH_WORK & "MN_VERSI.txt"            'w±@νID
Private Const EDIT_DATA_TUKA = PATH_SHUKEI_SEND & "TUKA*.csv"           'Κίf[^
Private Const EDIT_DATA_RIYO = PATH_SHUKEI_SEND & "ICRIYO*.csv"         'pΰzf[^
Private Const EDIT_DATA_GATECFG = PATH_WORK & "JP_CFG####"              'όD@Ϋηέθf[^  'EG30 V32.1.0.1 ADD
Private Const APL = "APL"
Private Const LOG = "LOG"
Private Const Data = "DATA"
Private Const BACKUP = "BACKUP"

Private Const MAX_KOMOKU_NUM_TUKA = 51                      'ΚίO}ΜΕεΪ
Private Const MAX_KOMOKU_NUM_KINGAKU = 16                   'ΰzO}ΜΕεΪ
'EG20 V30.1.0.1 ADD START
Private Const MAX_TUKA_SHUKEI_KOUMOKU = 7                                 '²όΚίf[^ΜΕεWvΪiubNPΚj
Private Const MAX_KOMOKU_NUM_TUKA_KAN = 51                                '²όΚίf[^ ΕεΪ
Private Const MAX_KOMOKU_NUM_UNKOU_FUNOU = 1                              '²όΚίf[^ ^ss\f[^ ΕεΪ
Private Const MAX_KOMOKU_NUM_NORIKAE_TUKA = 51                            '²ό ζ· έόΚίf[^ ΕεΪ
Private Const MAX_KOMOKU_NUM_JIEKI_KYUSAI = 51                            '²ό ©wόκ~ΟΚίf[^ ΕεΪ
Private Const MAX_KOMOKU_NUM_JIKI_KAISHU_CHUSHI = 51                      '²ό ₯Cρϋ~Κίf[^ ΕεΪ

Private Const MAX_KINGAKU_SHUKEI_KOUMOKU = 11                             '²όΰzf[^ΜΕεWvΪiubNPΚj
Private Const MAX_KOMOKU_NUM_SUICA_RIYO = 11                              '²όΰzf[^ XCJpΰz@ΕεΪ
Private Const MAX_KOMOKU_NUM_SUICA_SEISAN = 32                            '²όΰzf[^ XCJοΠΤΈZf[^ ΕεΪ
Private Const MAX_KOMOKU_NUM_AUTOCHARGE = 34                              '²όΰzf[^ I[g`[Wf[^@ΕεΪ


'WvΪiΚίf[^j
Private Enum mintTukaShukeiKoumoku
    SHUKEI_KAISATU_KANSEN_TUKA = 0      'yόD€@V²όΚίf[^z
    SHUKEI_SHUSATU_KANSEN_TUKA          'yWD€@V²όΚίf[^z
    SHUKEI_IC_UNKO_FUNOU                'y^ss\f[^z
    SHUKEI_KAN_ZAI_TUKA                 'y²-έζ·Κίf[^z
    SHUKEI_ZAI_KAN_TUKA                 'yέ-²ζ·Κίf[^z
    SHUKEI_JIEKI_KYUSAI                 'y©wόκ~ΟΚίf[^z
    SHUKEI_KAISHU_CHUSHI                'y₯Cρϋ~Κίf[^z
End Enum

'WvΪiΰzf[^j
Private Enum mintKingakuShukeiKoumoku
    SHUKEI_KAI_OTONA_SUICA_RIYO         'yόD€@εl@V²όXCJpvΰzz
    SHUKEI_SHU_OTONA_SUICA_RIYO         'yWD€@εl@V²όXCJpvΰzz
    SHUKEI_KAI_SHONI_SUICA_RIYO         'yόD€@¬@V²όXCJpvΰzz
    SHUKEI_SHU_SHONI_SUICA_RIYO         'yWD€@¬@V²όXCJpvΰzz
    SHUKEI_SEISAN_SHIHARAI              'yXCJοΠΤΈZf[^@^ΐx₯zz
    SHUKEI_KAI_AUTOCHARGE               'yόD€@I[g`[Wf[^z
    SHUKEI_SHU_AUTOCHARGE               'yWD€@I[g`[Wf[^z
    SHUKEI_KAN_OTONA_SUICA_RIYO         'y²ό^ΐ@εl@XCJpvΰzz
    SHUKEI_KAN_SHONI_SUICA_RIYO         'y²ό^ΐ@¬@XCJpvΰzz
    SHUKEI_NORI_ZAI_OTONA_SUICA_RIYO    'yζ·έ^ΐ@εl@XCJpvΰzz
    SHUKEI_NORI_ZAI_SHONI_SUICA_RIYO    'yζ·έ^ΐ@¬@XCJpvΰzz
End Enum

'GAIBU_OUTPUT.INIΜL[Τ
Private Enum mintGaibuOutputKey
    GAIBU_INI_TUKA = 0                  'Κίf[^
    GAIBU_INI_ICSF_KIKAN                'ICSF­sϊΤΚpΰzf[^
    GAIBU_INI_IC_CARD_SHIHARAI          'ICJ[hοΠΤΈZf[^i^ΐx₯zj
    GAIBU_INI_AUTO_CHARGE               'I[g`[Wf[^
    GAIBU_INI_IC_UNKOU_FUNOU            'ICJ[h^ss\f[^
    GAIBU_INI_TUKA_KAN_ZAI              '²-έζ·Κίf[^
    GAIBU_INI_TUKA_ZAI_KAN              'έ-²ζ·Κίf[^
    GAIBU_INI_IC_KIKAN_KANSEN           '²ό^ΐIC­s@ΦΚpΰzf[^
    GAIBU_INI_IC_KIKAN_ZAIRAI           'ζ·έ^ΐIC­sϊΤΚpΰzf[^
    GAIBU_INI_KYUSAI                    '©wόκ~ΟΚίf[^
    GAIBU_INI_KAISHU_CHUSI              '₯Cρϋ~Κίf[^
End Enum

'Private ReadSetteiSubGate()             As JIKAIINFO_IMAGE_FILE     'EKI_DISP_SUBGATE.CSVΜ1R[iͺΜf[^    'EG20 V30.3.0.1 yHKRK_Kansi07_008_01z DEL
'EG20 V30.1.0.1 ADD END
Private ReadSetteiSubGate(0 To 191)           As JIKAIINFO_IMAGE_FILE     'EKI_DISP_SUBGATE.CSV 1`32@ @`EiΕθj     'EG20 V30.3.0.1 yHKRK_Kansi07_008_01z ADD
Private Const SUBGATE_ITEM_NUM = 6      ' SUB_GATE_KAN.INIΜ©ΠͺΜΪF6                                   'EG20 V30.3.0.1 yHKRK_Kansi07_008_01z ADD

Private Const MAX_JPR_KETA_MAX = 30 'JPR1sΕε30oCg(Όp30Ά)

Private Const MAX_KADO_PG = 6       'όD@1δ½θΙόιvOivO»θf[^)

Private Const FOOTER_STRING = "*************END**************"


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : cmdPrint_Click
'//  @\ΌΜ  : uσvtΊ
'//  @\Tv  : σόπΐs·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17  CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub cmdPrint_Click()
    Dim i       As Integer
    Dim bRet    As Boolean
    Dim intCount    As Integer
    
    'uW[iσζΚFσJnvOoΝ
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JPR_PRINT_BUTTON, 0)
    
    '{^A`FbN{bNXπρANeBuΙ·ι
    Call JPRScreenEnable(False)
    
    ' έθσΤπζΎ·ι
    Call GetPrintSettings
    
    ' R[i`FbN
    If udtJprPrintSetteingInfo.iCornerCount = 0 Then
        'R[iΙ½ΰ`FbN³κΔ’Θ’ΜΕΈs
        LstStatus.AddItem "R[iͺ`FbN³κΔ’άΉρ"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        Call JPRScreenEnable(True)
        Exit Sub
    End If
    
    '@`FbN
    If udtJprPrintSetteingInfo.iGoukiCount = 0 Then
        '@Ι½ΰ`FbN³κΔ’Θ’ΜΕΈs
        LstStatus.AddItem "@ͺ`FbN³κΔ’άΉρ"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        Call JPRScreenEnable(True)
        Exit Sub
    End If
    
    'σΪwθ`FbN
    If udtJprPrintSetteingInfo.iJprCount = 0 Then
        'R[iΙ½ΰ`FbN³κΔ’Θ’ΜΕΈs
        LstStatus.AddItem "σΪͺwθ³κΔ’άΉρ"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        Call JPRScreenEnable(True)
        Exit Sub
    End If
    
    '`FbN³κ½R[iΝέu³κΔ’ι©’Θ’©ΜξρπZbg΅Δ¨­
    Erase glngTergetCorner
    For intCount = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        '»ΜR[iͺέu³κΔ’ι©H
        If IsTaisyoCorner(udtJprPrintSetteingInfo.iCorner(intCount)) = True Then
            glngTergetCorner(udtJprPrintSetteingInfo.iCorner(intCount) - 1) = CMN_ONOFF.CMN_ON
        Else
            glngTergetCorner(udtJprPrintSetteingInfo.iCorner(intCount) - 1) = CMN_ONOFF.CMN_OFF
        End If
    Next intCount
    
    'σΪ`FbNΙΑΔ?WπΔΡo·B
    iJprIdx = 0
    Call JprOutputProc
 
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : JPREdit_EkiInfo
'//  @\ΌΜ  : wsxf[^mF(wξρ)C[Wt@Cμ¬
'//  @\Tv  : wsxf[^mF(wξρ)ΜW[iC[Wt@Cπμ¬·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-27   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG30 V32.1.0.1) 2016-06-15  REVISED BY [TCC] T.Nakajima
'//                 2016Nx{τΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function JPREdit_EkiInfo() As Boolean

    Dim strFileName          As String          't@CΌ
    Dim iResponse            As Integer         'MsgBoxίθl
    Dim lRetVal              As Long            'ίθl
    Dim sCommand             As String          'R}hΆρ
'V1.12.0.1 ADD START
    Dim sWriteDir            As String              '«έζtH_Ό
    Dim intFileNumber        As Integer             't@C|C^
    Dim strLineCount         As String              'sJE^
    Dim i                    As Integer             '[vJE^P
    Dim j                    As Integer             '[vJE^Q
    Dim k                    As Integer             '[vJE^R
    Dim l                    As Integer             '[vJE^S
    Dim ReadFileSettei()     As EKIINFO_IMAGE_FILE  't@CΗp\’Μ
    Dim fso         As New FileSystemObject         't@CVXeIuWFNg
    Dim FsoTS As TextStream

    Dim bRet                 As Boolean         'Φίθl
    Dim lErrCode             As Long            'G[R[h
    
    Dim strNowType          As String           'εͺή
    Dim strNowShoNo         As String           'ΪΤ
    Dim strNowTuban         As String           'ΪΚΤ
    Dim strNowCorner        As String           'R[i
    Dim strNowKubn          As String           'ζͺ
    
    Dim intCount            As Integer          'R[iCfbNX OFR[i1
    
    'EG30 V32.1.0.1 ADD START
    Dim strEkiSettiBefPath  As String           '»έwέθf[^iΟXOΫΆj
    Dim strGetValue         As String * 64      'DLLΙζΑΔέθ³κι½ίA64Εθ·Ι΅Δ’ι
    Dim strCompValue        As String           'έθliΟXOΫΆj
    Dim strChangeFlg        As String           'ΟXσ
    Dim intValueLen         As Integer          'ζΎ΅½έθlΜ·³
    'EG30 V32.1.0.1 ADD END
    
    On Error GoTo Err_handler
    
    
    'ζΚΕwθ³κ½R[iΝέu³κΔ’ι©H
    If pfSettingCheck(False) = False Then
        '·ΧΔ’έuΜR[iΘΜΕG[Ζ·ι
        GoTo Err_handler
    End If
    
    '////////////////////////////////////////////////
    '// R[iΌπκΚθζΎ
    gsGetCornerName
   
    'eXgΕΖθ ¦ΈAR[i1
    intCount = 0
    
    'wsxf[^mFiwξρjC[Wt@Cμ¬
    bRet = dllGetEkiIniData(0, EKI_TUDO_CHK_EKI_INFO_FILE, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        'wsxf[^mFiwξρjC[Wt@Cν
        Kill EKI_TUDO_CHK_EKI_INFO_FILE
        'ΩνOoΝ
        Call pfOutPutErrLog(lErrCode)
        JPREdit_EkiInfo = False
        Exit Function
    End If
    
    'CSVt@CΜζΎ
    intFileNumber = FreeFile
    Open EKI_TUDO_CHK_EKI_INFO_FILE For Input As #intFileNumber
    
    Do While Not EOF(intFileNumber)                     ' EG20 V3.3.0.1ΗΑ
        Line Input #intFileNumber, strLineCount
        j = j + 1
    Loop
    'CSVt@CN[Y
    Close #intFileNumber
    
    'γLͺAγΙΫ
    'Δέθ
    ReDim ReadFileSettei(j) As EKIINFO_IMAGE_FILE   't@CΗpGA
        
    'CSVt@CI[v
    Open EKI_TUDO_CHK_EKI_INFO_FILE For Input As #intFileNumber

    'Xg\¦ͺΗέέit@CI[άΕ[vπJθΤ·j
    For i = 0 To UBound(ReadFileSettei) - 1
        Input #intFileNumber, ReadFileSettei(i).sType, ReadFileSettei(i).sGoki, ReadFileSettei(i).sNo, _
        ReadFileSettei(i).sCorner, ReadFileSettei(i).sTuuban, ReadFileSettei(i).sKoumoku, ReadFileSettei(i).sKubun, _
        ReadFileSettei(i).sSettei, ReadFileSettei(i).sSyosai
    Next i

    'CSVt@CN[Y
    Close #intFileNumber
    
    'EG30 V32.1.0.1 ADD START
    '»ΜR[iΜΟXOf[^ΫΆ³κ½f[^πγΙWJ·ι
    strEkiSettiBefPath = Replace(EKI_SETTI_FILE_BEF, "#", "0")
    Call dllGetEkiIniDataBefore(strEkiSettiBefPath, lErrCode)
    'EG30 V32.1.0.1 ADD END
    
    '/////////////////////////////////////
    'W[iC[Wt@Cμ¬
    '’gpΜt@CΤζΎ
    intFileNumber = FreeFile
   
    'W[ioΝC[Wt@Cπμ¬
    Open EKI_JPR_EKIINFO_TXTFILE For Output As #intFileNumber
    
    '^Cg\¦
    'PrintHeader intFileNumber, "wsxf[^mFiwξρj"    'EG30 V32.1.0.1 DEL
    PrintHeader3 intFileNumber, "wsxf[^mFiwξρj", pfGetSaveDate(0) 'EG30 V32.1.0.1 ADD
    Print #intFileNumber, "έuwF" & Trim(pfGetEkiNameInfo(NotEkiVer))
    '`FbN³κ½R[iͺΕ[v
    For k = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        intCount = udtJprPrintSetteingInfo.iCorner(k) - 1  'ζΚΕwθ³κ½R[i-1
        If pfCornerGokiCheck(udtJprPrintSetteingInfo.iCorner(k)) = True Then
            
            ' »ΜR[iΝέu³κΔ’ιΜΕW[ioΝΦ
            '1R[iΪΎ―έuwΖέuR[iΜΤΝσsͺΘ’
            If k <> 0 Then
                Print #intFileNumber, ""
            End If
            Print #intFileNumber, "έuR[iF" & gstrCornerName(intCount)

            '////////////////////////////////
            '// eέθπoΝ
            '////////////////////////////////
            strNowType = ""
            strNowShoNo = ""
            strNowKubn = ""
            
            For i = 0 To UBound(ReadFileSettei) - 1
            
                If strNowType <> ReadFileSettei(i).sType Then
                    'V΅’εͺήζͺΙΘΑ½ΜΕ^Cgπσ
                    Print #intFileNumber, ""
                    Select Case ReadFileSettei(i).sType
                        Case "1"
                            'Print #intFileNumber, "ywξρz"     'EG30 V32.1.0.1 DEL
                            Print #intFileNumber, "@ywξρz"    'EG30 V32.1.0.1 ADD
                        Case "2"
                            'Print #intFileNumber, "yΔz"   'EG30 V32.1.0.1 DEL
                            Print #intFileNumber, "@yΔz"  'EG30 V32.1.0.1 ADD
                        Case "3"
                            'Print #intFileNumber, "ylbg[Nz"   'EG30 V32.1.0.1 DEL
                            Print #intFileNumber, "@ylbg[Nz"  'EG30 V32.1.0.1 ADD
                        Case "7"
                            'Print #intFileNumber, "yζΚz"   'EG30 V32.1.0.1 DEL
                            Print #intFileNumber, "@yζΚz"  'EG30 V32.1.0.1 ADD
                    End Select
                    strNowType = ReadFileSettei(i).sType
                End If
                
                'ΪΤͺOρΖ―ΆκΝoΝ΅Θ’
                'If strNowShoNo <> ReadFileSettei(i).sNo Then
                If strNowShoNo <> ReadFileSettei(i).sNo Or strNowKubn <> ReadFileSettei(i).sKubun Then
                    'ΪΌ+ζͺ+έθlπo·
                    If (CInt(ReadFileSettei(i).sCorner) = intCount + 1) Or (CInt(ReadFileSettei(i).sCorner) = 0) Then
                        
                        'EG30 V32.1.0.1 ADD START
                        'ΟXOf[^ΫΆ³κ½έθlΖδr·ι
                        bRet = dllGetEkiInfoValue(CInt(ReadFileSettei(i).sType), _
                                                    CInt(ReadFileSettei(i).sGoki), _
                                                    CInt(ReadFileSettei(i).sNo), _
                                                    CInt(ReadFileSettei(i).sCorner), _
                                                    strGetValue, _
                                                    intValueLen)
                        strCompValue = strGetValue
                        If (intValueLen <> 0) Then
                            strCompValue = MidByte(strGetValue, 1, intValueLen)
                            strCompValue = Trim(strCompValue)
                        ElseIf (intValueLen = 0) Then
                            strCompValue = ""
                        End If
                        
                        If (bRet = False) Or (ReadFileSettei(i).sSettei <> strCompValue) Then
                            strChangeFlg = DIFF_MARK_STRING_ON
                        Else
                            strChangeFlg = DIFF_MARK_STRING_OFF
                        End If
                        'EG30 V32.1.0.1 ADD END
                        
                        
                        '/////////////////////////////////////////
                        '//ΊLΜΪΝwsxf[^ΖW[iΜoΝf[^ͺΩΘι`?ΙΘιΜΕ­§IΙΟ··ι
                        '/////////////////////////////////////////
                        
                        'εͺήFP ͺήFO ¬ͺήFPWuͺήvΜlΝu9 9 9 9 9 9v`? ΌpXy[X2Ά¨1ΆΙΟX
                        If (ReadFileSettei(i).sType = "1") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "18") Then
                           ReadFileSettei(i).sSettei = Replace(ReadFileSettei(i).sSettei, "  ", " ")
                        End If
                        'εͺήFQ ͺήFO ¬ͺήFPuR[iΤiΞhcT[o)vΜlΝuhcvΝSp
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "1") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "hc")
                        End If
                        'εͺήFQ ͺήFO ¬ͺήFQuiΞhcT[o)vΜlΝuhcvΝSp
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "2") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "hc")
                        End If
                        'εͺήFQ ͺήFO ¬ͺήFQuiΞhcT[o)vuhcvΝSp
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "2") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "hc")
                        End If
                        'εͺήFQ ͺήFO ¬ͺήFRuiΞhcT[o)vuhcvΝSp
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "3") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "hc")
                        End If
                        'εͺήFQ ͺήFO ¬ͺήFWuiΞhcT[o)vuhcvΝSp
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "8") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "hc")
                        End If

                        'εͺήFQ ͺήFO ¬ͺήFXuiΞhcT[o)vuhcvΝSp
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "9") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "hc")
                        End If
                     
                        'εͺήFV ͺήFO ¬ͺήFQPuΫη[Uέθj[ζΚ ³l[h?μέθtv
                        'ΪΌΖζͺΜΤΙXy[XπΠΖΒ½­όκι
                        If (ReadFileSettei(i).sType = "7") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "21") Then
                           ReadFileSettei(i).sKoumoku = ReadFileSettei(i).sKoumoku & Space(1)
                        End If
                     
                        'Print #intFileNumber, ReadFileSettei(i).sKoumoku & " " & ReadFileSettei(i).sKubun & " " & ReadFileSettei(i).sSettei    'EG30 V32.1.0.1 DEL
                        Print #intFileNumber, strChangeFlg & ReadFileSettei(i).sKoumoku & " " & ReadFileSettei(i).sKubun & " " & ReadFileSettei(i).sSettei  'EG30 V32.1.0.1 ADD
                        strNowShoNo = ReadFileSettei(i).sNo
                        strNowKubn = ReadFileSettei(i).sKubun
                    End If
                End If
            
            Next i
        Else
             'έu³κΔ’Θ’R[iΘΜΕΜR[iΦ
        End If
    Next k
    
    Print #intFileNumber, ""
    Print #intFileNumber, FOOTER_STRING
    
    Close #intFileNumber
    
    JPREdit_EkiInfo = True
    Exit Function
    
Err_handler:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    Set fso = Nothing
    'ΩνOoΝ
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JPR_PRINT_OUTPUT_ERR, 0)
    
    'ΩνIΉ
    'iResponse = MsgBox("ΩνIΉ΅ά΅½B", vbOKOnly + vbCritical, "wέθeLXgoΝΚ")
    JPREdit_EkiInfo = False

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : cmdReturn_Click
'//  @\ΌΜ  : uj[ζΚΦίιvtΊ
'//  @\Tv  : ©ζΚπΑ·ιB
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-14   CODED   BY [TCC] N.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    'uW[iσζΚFΑvOoΝ
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JPR_PRINT_GAMEN_END, 0)
    
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  ΦΌΜ  : Form_Activate
'//  @\ΌΜ  : VXeϊ»(ΔΥ)ζΚ(ANeBu)
'//  @\Tv  : ΕOΚ\¦πs€B
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    pfFormActive (hwnd)
    '[σM^C}πN?·ιB
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  ΦΌΜ  : Form_Deactivate
'//  @\ΌΜ  : VXeϊ»(ΔΥ)ζΚ(fBANeBu)
'//  @\Tv  : [σMpΜ^C}β~
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
   On Error Resume Next
    '[σM^C}πβ~·ιB
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : Form_Load
'//  @\ΌΜ  : W[iσζΚ([h)
'//  @\Tv  : ϊπs€B
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.0.1.1) 2014-03-25   CODED   BY [TCC] T.Nakajima
'//                 k€V²όJΖΞ
'//     REVISIONS :
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim i As Integer    'JE^[
   
    On Error Resume Next
    
    'uW[iσζΚF\¦vOoΝ
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JPR_PRINT_GAMEN_START, 0)
    
    ' R[i`FbN{bNX
    For i = 0 To chkCorner.Count - 1
        chkCorner(i).Value = 1
    Next i

    ' @`FbN{bNX
    For i = 0 To chkGouki.Count - 1
        chkGouki(i).Value = 1
    Next i

    ' f[^Ϊ
    For i = 0 To chkJprKind.Count - 1
        chkJprKind(i).Value = 0
    Next i
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
   '[σM^C}ΜC^[oπ'PbΙZbg
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   
   'INIt@CζθAvN?^C}lπζΎ
   lngMAX_Time = GetPrivateProfileInt(APLCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      APL_INTERVAL, HOSHU_FILE)
   'ζΎlͺ0ΜκAftHglπέθ
   If lngMAX_Time = 0 Then
      lngMAX_Time = APL_INTERVAL
   End If
   
   'INIt@CζθON?^C}lπζΎ
   lngLogMAX_Time = GetPrivateProfileInt(LOGCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      LOG_INTERVAL, HOSHU_FILE)
   'ζΎlͺ0ΜκAftHglπέθ
   If lngLogMAX_Time = 0 Then
      lngLogMAX_Time = LOG_INTERVAL
   End If
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : tmrMail_Timer
'//  @\ΌΜ  : [σM^C}A^CAbv
'//  @\Tv  : [πσM·ιB
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.1.0.1) 2014-03-25   CODED   BY [TCC] T.Nakajima
'//                 k€V²όJΖΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF           '[σMGA
    Dim lngLength As Long                    'σM[oCgTCY
    Dim lngMlSts  As Long                    'σM[ΜXe[^X
    Dim bRet  As Boolean
    Dim lngDataKind As Long                 'ζΚoΝvRESΜf[^νΚ
    
    On Error Resume Next

    '[πσM·ιB
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
   'σM[ͺ κΞA[hcΜπ·ιB
        Select Case udtReadMail.udtlHeader.dwId        '[hc
            Case ML_ID_JPR_PRINT_RES
                'uW[iσόvRESσM³νvOoΝ
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, JPR_PRINT_RES_RECV, 0)
                lngMlSts = udtReadMail.lngData(0)
                If (udtJprPrintSetteingInfo.iJprKind(iJprIdx) = JPR_KIND.JPR_KIND_TUKA_DATA) Or _
                   (udtJprPrintSetteingInfo.iJprKind(iJprIdx) = JPR_KIND.JPR_KIND_RIYO_KINGAKU) Then
                    
                    'Κίf[^ά½ΝpΰzπoΝ΅Δ’ιΖ«ΝW[iσvRESπσM΅½ηA
                    'WvΙζΚoΝ?ΉΚmπM·ιB
                    If lngMlSts = 0 Then
                        bRet = SendMessageGamenOutComplete(ML_GAMEN_OUT_STS.ML_STS_OK)
                    Else
                        bRet = SendMessageGamenOutComplete(ML_GAMEN_OUT_STS.ML_STS_NG)
                    End If
                Else
                    bRet = True
                End If
                
                If (lngMlSts = 0) And (bRet = True) Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), True)
                Else
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), False)
                End If
            
            Case ML_ID_INFO_RES
                'uξρvRESσM³νvOoΝ
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GETINFO_RES_RECV, 0)
                ' ΚπmF
                lngMlSts = udtReadMail.lngData(1)
                If lngMlSts = 0 Then
                    '?Wπs€B
                    bRet = JprEdit_EkimuId()
                    If bRet = True Then
                       'W[iσvCMDπM
                        bRet = SendMessageJprPrint(EKIMUKIKI_ID_TXTFILE, ML_CUT_ARI)
                        If bRet = False Then
                            Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                            Exit Sub
                        End If
                    Else
                        Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), False)
                    End If
                    
                Else
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), False)
                End If
            
            Case ML_ID_GAMEN_OUTPUT_RES
                'uζΚoΝvRESσM³νvOoΝ
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GETINFO_RES_RECV, 0)
                'ΚmF
                lngMlSts = udtReadMail.lngData(1)
                'f[^νΚ
                lngDataKind = udtReadMail.lngData(2)
                If lngMlSts = 0 Then
                    '?Wπs€
                    bRet = JprEdit_TukaData(lngDataKind)
                    If bRet = True Then
                        'W[iσvCMDπM
                        If lngDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
                            bRet = SendMessageJprPrint(TUKA_TXTFILE, ML_CUT_ARI)
                        ElseIf lngDataKind = Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI Then
                            bRet = SendMessageJprPrint(ICRIYO_TXTFILE, ML_CUT_ARI)
                        Else
                            bRet = False
                        End If
                            
                        If bRet = False Then
                            '?WͺΈsAζΚoΝ?ΉΚmπΩνΕM
                            SendMessageGamenOutComplete (ML_GAMEN_OUT_STS.ML_STS_NG)
                            'ΩνΘΜΕAζΚoΝ?ΉΚmbZ[WΜMΙΈs΅ζ€ͺΚΝΩν
                            Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                            Exit Sub
                        End If
                    Else
                        '?WͺΈsAζΚoΝ?ΉΚmπΩνΕM
                        SendMessageGamenOutComplete (ML_GAMEN_OUT_STS.ML_STS_NG)
                        'ΩνΘΜΕAζΚoΝ?ΉΚmbZ[WΜMΙΈs΅ζ€ͺΚΝΩν
                        Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                        Exit Sub
                    End If
                Else
                    'RESͺΩνΜ½ίAIΉiζΚoΝ?ΉΚmΝM΅Θ’)
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), False)
                End If
                
            Case ML_ID_PROEND_ORD
                'uvZXIΉw¦vπσM΅½κA
                'uvZXIΉw¦σM³νvOoΝ
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
                'vZXΜIΉπs€
                pfAbortProc
            
            Case ML_ID_HOSHU_ACTIVE_REQ
                'uΫηζΚANeBu\¦vσM³νvOoΝ
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                'ΫηζΚANeBuvπσM΅½ηA©ζΚπOΚΙ\¦³ΉιB
                AppActivate frmJprPrint.Caption, False
                pfFormActive (frmJprPrint.hwnd)
                
            Case Else
                'u[IDs³vOoΝ
                Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  ΦΌΜ  : SendMessageJprPrint
'//  @\ΌΜ  : W[iσvbZ[WπM·ι
'//  @\Tv  : oΝvZXΙW[iσvπM·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : String    strFileName   oΝt@CΌ
'//              Byte      byCut         0:JbgΘ΅   1FJbg θ
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function SendMessageJprPrint(strFileName As String, byCut As Byte) As Boolean

    Dim udtMail As MAIL_JPR_PRINT_CMD   'W[iσόv[MGA
    Dim lngRet As Long                  'Φίθl
    Dim lngErrCode As Long              'G[R[h
    Dim bTmpArray() As Byte
    Dim i       As Integer
    On Error Resume Next
    
    
    'W[iσvπoΝvZXΙM·ιB
    udtMail.mlHeader.dwId = ML_ID_JPR_PRINT_REQ
    udtMail.mlHeader.dwSize = MlSize.JPR_PRINT_REQ
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    bTmpArray = StrConv(strFileName, vbFromUnicode)
    For i = 0 To UBound(bTmpArray)
        'udtMail.byOutputFilePath(i) = Chr(bTmpArray(i))
        udtMail.byOutputFilePath(i) = bTmpArray(i)
    Next
    udtMail.dwCut = byCut                                   'JbgL³
    udtMail.dwOutputDataPoint = 0                           'oΝf[^|Cg
    
    lngRet = DssSendMail(MAIL_SLOT_OUTPUT, Len(udtMail), udtMail.mlHeader)
    If lngRet = False Then
       'uW[iσζΚFW[iσόvMΩνvOoΝ
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, JPR_PRINT_REQ_SEND, lngErrCode)
       SendMessageJprPrint = False
       Exit Function
    Else
       'uW[iσζΚFW[iσόvM³νvOoΝ
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, JPR_PRINT_REQ_SEND, 0)
       SendMessageJprPrint = True
    End If
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : SendMessageInfoReq
'//  @\ΌΜ  : ξρvCMDbZ[WπM·ι
'//  @\Tv  : ID§ΙξρvvCMDπM·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function SendMessageInfoReq() As Boolean
    
    Dim bRet As Boolean                 'ίθl
    Dim lngErrCode As Long              'G[R[h
    Dim udtMail As MAIL_INFO_CMD        'ζΚ\¦v
    Dim uMail As ML_KYOTU_INF           '[
 
   'obt@tbVvπOvZXΙM·ι
   'ξρvCMD(w±@νID=0)πID§δΙM·ι
   udtMail.mlHeader.dwId = ML_ID_INFO_CMD
   udtMail.mlHeader.dwSize = MlSize.INFO_CMD
   udtMail.mlHeader.dwProid = RHOSHU_ID
   udtMail.mlHeader.dwSubArea = 0
   udtMail.dwRequestType = MailCmdType.ML_DT_EKIMU_ID
   iSendType = MailCmdType.ML_DT_EKIMU_ID
   bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
   If bRet = False Then
      'uw±@νIDmFFξρvCMDMΩνvOoΝ
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
      SendMessageInfoReq = False
      Exit Function
   Else
      'uw±@νIDmFFξρvCMDM³νvOoΝ
      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
   End If
   
   SendMessageInfoReq = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : SendMessageGamenOutReq
'//  @\ΌΜ  : ζΚoΝvCMDM
'//  @\Tv  : WvΙζΚoΝvCMDπM·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function SendMessageGamenOutReq(dwDataKind As Long) As Boolean
    
    Dim bRet As Boolean                     'ίθl
    Dim lngErrCode As Long                  'G[R[h
    Dim udtMail As MAIL_GAMEN_OUTPUT_CMD    'ζΚoΝv
    Dim uMail As ML_KYOTU_INF               '[
 
   'obt@tbVvπOvZXΙM·ι
   'ζΚoΝvCMDπWvΙM·ι
   udtMail.mlHeader.dwId = ML_ID_GAMEN_OUTPUT_REQ
   udtMail.mlHeader.dwSize = MlSize.GAMEN_OUT_REQ
   udtMail.mlHeader.dwProid = RHOSHU_ID
   udtMail.mlHeader.dwSubArea = 0
   udtMail.dwSeqence = 0                ' V[PXΤ0Εθ
   udtMail.dwDataKind = dwDataKind
   bRet = DssSendMail(MAIL_SLOT_SHUKEI, Len(udtMail), udtMail.mlHeader)
   If bRet = False Then
      'uζΚoΝvCMDMΩνvOoΝ
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
      SendMessageGamenOutReq = False
      Exit Function
   Else
      'uζΚoΝvCMDM³νvOoΝ
      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
   End If

   SendMessageGamenOutReq = True

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : SendMessageGamenOutComplete
'//  @\ΌΜ  : ζΚoΝv?ΉΚmM
'//  @\Tv  : WvΙζΚoΝ?ΉΚmπM·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Long     dwStatus   bZ[WΙZbg·ιXe[^X
'//
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function SendMessageGamenOutComplete(dwStatus As Long) As Boolean
    
    Dim bRet As Boolean                     'ίθl
    Dim lngErrCode As Long                  'G[R[h
    Dim udtMail As MAIL_GAMEN_OUTPUT_COMP   'ζΚoΝv?ΉΚm
    Dim uMail As ML_KYOTU_INF               '[
 
   'obt@tbVvπOvZXΙM·ι
   'ζΚoΝvCMDπWvΙM·ι
   udtMail.mlHeader.dwId = ML_ID_GAMEN_OUTPUT_COMP
   udtMail.mlHeader.dwSize = MlSize.GAMEN_OUT_COMP
   udtMail.mlHeader.dwProid = RHOSHU_ID
   udtMail.mlHeader.dwSubArea = 0
   udtMail.dwSequence = 0                ' V[PXΤ0Εθ
   udtMail.dwStatus = dwStatus
   bRet = DssSendMail(MAIL_SLOT_SHUKEI, Len(udtMail), udtMail.mlHeader)
   If bRet = False Then
      'uζΚoΝv?ΉΚmMΩνvOoΝ
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
      SendMessageGamenOutComplete = False
      Exit Function
   Else
      'uζΚoΝv?ΉΚmM³νvOoΝ
      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
   End If
   
   SendMessageGamenOutComplete = True

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : ResultDisp
'//  @\ΌΜ  : W[iσόΚ\¦
'//  @\Tv  : W[iΜσόΚπ\¦·ιB
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Integer    iJprKind    W[iνΚ
'//              Boolean    bResult     Κ(true/false)
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(30.3.0.1) 2014-10-01  REVISED BY [TCC] T.Nakajima
'//                 k€V²όtF[YQΞ yHKRK_Kansi07_003_01z
'//     REVISIONS :(32.1.0.1) 2016-06-10  REVISED BY [TCC] T.Nakajima
'//                 2016Nx{τΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub ResultDisp(iJprKind As Integer, bResult As Boolean)
    Dim strStatus   As String
    Dim strJprName  As String

    'ΚΆΎμ¬
    Select Case iJprKind
        Case JPR_KIND.JPR_KIND_EKI_INFO
            strJprName = "wsxf[^mFiwξρj"
            
        Case JPR_KIND.JPR_KIND_JIKAI_INFO
            strJprName = "wsxf[^mFi©όj"
            
        Case JPR_KIND.JPR_KIND_SETTING_LST
            strJprName = "έθlκ"
            
        Case JPR_KIND.JPR_KIND_TUKA_DATA
            strJprName = "Κίf[^"
            
        Case JPR_KIND.JPR_KIND_RIYO_KINGAKU
            strJprName = "pΰzf[^"
            
        Case JPR_KIND.JPR_KIND_KADO_VER
            strJprName = "?­o[Wκ"
            
        Case JPR_KIND.JPR_KIND_SIMEKIRI
            strJprName = "χΨItCoΝ"
            
        Case JPR_KIND.JPR_KIND_EKIMU_ID
            strJprName = "w±@νhc"
        'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z ADD START
        Case JPR_KIND.JPR_KIND_SUBGATE_INFO
            strJprName = "wsxf[^mFi΄έΊ°ΔήΊ°Ε@ξρθ`j"
        'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z ADD END
        'EG30 V32.1.0.1 ADD START
        Case JPR_KIND.JPR_KIND_GATE_CFG
            strJprName = "όD@Ϋηέθf[^"
        'EG30 V32.1.0.1 ADD END
    End Select
    
    If bResult = True Then
        '³ν
        LstStatus.AddItem strJprName & "    " & "³νIΉ΅ά΅½"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        'Call JPRScreenEnable(True)
    Else
        'Ων
        LstStatus.AddItem strJprName & "    " & "ΩνIΉ΅ά΅½"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        'Call JPRScreenEnable(True)
    End If

    iJprIdx = iJprIdx + 1
    If iJprIdx < udtJprPrintSetteingInfo.iJprCount Then
        '2νήΪΘ~ΜW[ioΝ
        JprOutputProc
    Else
        'SW[ioΝ?ΉΘηΞAtA`FbN{bNXμΒ\Ι·ιB
        Call JPRScreenEnable(True)
        iJprIdx = 0
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : JPRScreenEnable
'//  @\ΌΜ  : W[iσόζΚΜέθΟXΒΫ§δ
'//  @\Tv  : W[iσζΚΜΰeπΟXΜΒΫπ§δ·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Boolean   bEnable    true:ΟXΒ\  false:ΟXsΒ
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG30 V32.2.0.1) 2016-07-20  CODED   BY [TCC] T.Nakajima
'//                 2016Nx{τΞ vOXo[ρ\¦Ξ
'//                 [ΌͺΘγπgp·ιW[iͺ ι½ίAvOXo[\¦ΙΤΙνΘ­Θι½ί
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub JPRScreenEnable(bEnable As Boolean)
    Dim i   As Integer
    
    ' R[i`FbN{bNX
    For i = 0 To chkCorner.Count - 1
        chkCorner(i).Enabled = bEnable
    Next i
    
    ' @`FbN{bNX
    For i = 0 To chkGouki.Count - 1
        chkGouki(i).Enabled = bEnable
    Next i
 
    ' f[^Ϊ
    For i = 0 To chkJprKind.Count - 1
        chkJprKind(i).Enabled = bEnable
    Next i
    
    'σ{^
    cmdPrint.Enabled = bEnable
    
    'ίι{^
    cmdReturn.Enabled = bEnable
    
    If bEnable = False Then
        'Xe[^X\¦πNA·ιiσ{^ΊΕOρΜΚπNA)
         LstStatus.Clear
        'vOXo[π\¦·ι
        ' EG30 V32.2.0.1 DEL START
        'Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_JPR_OUT)
        ' EG30 V32.2.0.1 DEL END
    Else
        'vOXo[πΑ·ι
        ' EG30 V32.2.0.1 DEL START
        'Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ, PRG_JPR_OUT)
        ' EG30 V32.2.0.1 DEL END
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : GetPrintSettings
'//  @\ΌΜ  : ζΚΜ`FbNσΤπζΎ
'//  @\Tv  : wθ³κ½R[iπζΎ·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Boolean   bEnable    true:ΟXΒ\  false:ΟXsΒ
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub GetPrintSettings()
    Dim i               As Integer
    Dim k               As Integer
    Dim iCornerCount    As Integer
    Dim iGoukiCount     As Integer
    Dim iJprCount       As Integer
    
    ' W[iέθξρπNA·ι
    udtJprPrintSetteingInfo = udtInitJprSetting
    
    
    ' R[iΜ`FbNσΤ
    k = 0
    For i = 0 To chkCorner.Count - 1
        If chkCorner(i).Value = 1 Then
            iCornerCount = iCornerCount + 1
            udtJprPrintSetteingInfo.iCorner(k) = i + 1
            k = k + 1
        End If
    Next i
    udtJprPrintSetteingInfo.iCornerCount = iCornerCount
    
    '@Μ`FbNσΤ
    k = 0
    For i = 0 To chkGouki.Count - 1
        If chkGouki(i).Value = 1 Then
            iGoukiCount = iGoukiCount + 1
            udtJprPrintSetteingInfo.iGouki(k) = i + 1
            k = k + 1
        End If
    Next i
    udtJprPrintSetteingInfo.iGoukiCount = iGoukiCount
    
    'W[iνΚΜ`FbNσΤ
    k = 0
    For i = 0 To chkJprKind.Count - 1
        If chkJprKind(i).Value = 1 Then
            iJprCount = iJprCount + 1
            udtJprPrintSetteingInfo.iJprKind(k) = i
            k = k + 1
        End If
    Next i
    udtJprPrintSetteingInfo.iJprCount = iJprCount
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : JprOutputProc
'//  @\ΌΜ  : W[ioΝ
'//  @\Tv  : oΝt@Cμ¬ΖoΝvZXΙvπM
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(30.3.0.1) 2014-10-01  REVISED BY [TCC] T.Nakajima
'//                 k€V²όtF[YQΞ yHKRK_Kansi07_003_01z
'//     REVISIONS :(32.1.0.1) 2016-06-10  REVISED BY [TCC] T.Nakajima
'//                 2016Nx{τΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub JprOutputProc()
    Dim bRet        As Boolean
    
    'EG30 V32.1.0.1 ADD START
    Dim i, j            As Integer  'R[iA@JE^
    Dim intComSts       As Integer  'ΚMσΤ
    Dim blnSkipFlg      As Boolean  'Ϋηέθf[^Θ΅
    Dim intGateNo       As Integer  '@Τi1`32j
    'EG30 V32.1.0.1 ADD END
    Select Case udtJprPrintSetteingInfo.iJprKind(iJprIdx)
        Case JPR_KIND.JPR_KIND_EKI_INFO          ' wsxf[^mF(wξρ)
            bRet = JPREdit_EkiInfo
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_EKI_INFO, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(EKI_JPR_EKIINFO_TXTFILE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If
            
        Case JPR_KIND.JPR_KIND_JIKAI_INFO            ' wsxf[^mF(©ό)
            bRet = JPREdit_JikaiInfo
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_JIKAI_INFO, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(EKI_JPR_GATE_TXTFILE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If

        Case JPR_KIND.JPR_KIND_SETTING_LST           ' έθlκ
            bRet = JprEdit_SetteiList
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_SETTING_LST, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(SETTI_TXTFLE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If
        
        Case JPR_KIND.JPR_KIND_TUKA_DATA             ' Κίf[^
            ' bZ[WπM΅Δ?W³t@CΜμ¬πΛ·ιΜΕA±±ΕΝ?WΝΔΞΘ’B
            ' ?WπΔΤΜΝRES[πσM΅½Ζ«BW[iσvΝ?WͺIνΑΔ©ηΔΤB
            ' wθ³κ½R[iͺ’έuΘηΞΝ΅Θ’
            If pfSettingCheck(False) = True Then
                bRet = SendMessageGamenOutReq(Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI)
                If bRet = False Then
                    Call ResultDisp(JPR_KIND.JPR_KIND_TUKA_DATA, bRet)
                    Exit Sub
                End If
            Else
                Call ResultDisp(JPR_KIND.JPR_KIND_TUKA_DATA, False)
                Exit Sub
            End If
        
        Case JPR_KIND.JPR_KIND_RIYO_KINGAKU          ' pΰzf[^
            ' bZ[WπM΅Δ?W³t@CΜμ¬πΛ·ιΜΕA±±ΕΝ?WΝΔΞΘ’B
            ' ?WπΔΤΜΝRES[πσM΅½Ζ«BW[iσvΝ?WͺIνΑΔ©ηΔΤB
            ' wθ³κ½R[iͺ’έuΘηΞΝ΅Θ’
            If pfSettingCheck(False) = True Then
                bRet = SendMessageGamenOutReq(Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI)
                If bRet = False Then
                    Call ResultDisp(JPR_KIND.JPR_KIND_RIYO_KINGAKU, bRet)
                    Exit Sub
                End If
            Else
                Call ResultDisp(JPR_KIND.JPR_KIND_RIYO_KINGAKU, False)
                Exit Sub
            End If
        
        Case JPR_KIND.JPR_KIND_KADO_VER              ' ??o[Wκ
            bRet = JprEdit_KadoVersion
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_KADO_VER, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(KADOVER_TXTFILE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If
        
        Case JPR_KIND.JPR_KIND_SIMEKIRI              ' χΨItCoΝ
            bRet = JprEdit_SimekiriOffline
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_SIMEKIRI, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(SIMEKIRI_TXTFILE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If
        
        Case JPR_KIND.JPR_KIND_EKIMU_ID              ' w±@νID
            ' bZ[WπM΅Δ?W³t@CΜμ¬πΛ·ιΜΕA±±ΕΝ?WΝΔΞΘ’B
            ' ?WπΔΤΜΝRES[πσM΅½Ζ«BW[iσvΝ?WͺIνΑΔ©ηΔΤB
            bRet = SendMessageInfoReq
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_EKIMU_ID, bRet)
                Exit Sub
            End If
        'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z ADD START
        Case JPR_KIND.JPR_KIND_SUBGATE_INFO         ' wsxf[^mF(GR[hR[i@ξρθ`)
            bRet = JPREdit_SubGateInfo
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_SUBGATE_INFO, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(EKI_JPR_SUBGATE_TXTFILE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If
        'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z ADD END
        'EG30 V32.1.0.1 ADD START
        Case JPR_KIND.JPR_KIND_GATE_CFG             ' όD@Ϋηέθf[^
            '`FbN³κΔ’ιόD@ΜΚMσΤπζΎ·ι
            For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
                For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
                    '»ΜR[iA@Νέu³κΔ’ι©H
                    If pfCornerGokiToGateNo(udtJprPrintSetteingInfo.iCorner(i), udtJprPrintSetteingInfo.iGouki(j), intGateNo) = True Then
                    
                        'ΔΥN?L³`FbN
                        If CheckAppStart(PROC_KANRI) <> 0 Then
                            gpfGetjikaiConectSts intComSts, intGateNo
                            If intComSts <> CONECTSTS_NORMAL Then
                                Exit For
                            End If
                        End If
                    End If
                Next j
                '1δΕΰΚMΩνΜόD@ͺ κΞAxπ\¦·ιΜΕAR[iPΚΜ[vπ²―ι
                If intComSts <> CONECTSTS_NORMAL Then
                    Exit For
                End If
            Next i
            
            'Xe[^X\¦ΙΚMΩνόD@ͺ ι±Ζπ\¦·ι
            If intComSts <> CONECTSTS_NORMAL Then
                LstStatus.AddItem "Iπ΅½R[iΙΚMΩνΜόD@ͺ θά·"
                LstStatus.AddItem "ΚMΩν@ΜόD@Ϋηέθf[^ΝΕVΕ³’Β\«ͺ θά·"
                LstStatus.Selected(LstStatus.ListCount - 1) = True
            End If
            
            bRet = JprEdit_GateCfg(blnSkipFlg)
            'όD@Ϋηέθf[^’σMΜόD@ͺ Α½½ίAW[iσΕ«Θ©Α½±Ζπ\¦·ιB
            If blnSkipFlg = True Then
                LstStatus.AddItem "όD@Ϋηέθf[^πσΕ«Θ©Α½όD@ͺ θά·"
                LstStatus.Selected(LstStatus.ListCount - 1) = True
            End If
            
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_GATE_CFG, bRet)
                Exit Sub
            Else
                bRet = SendMessageJprPrint(GATE_CFG_TXTFILE, ML_CUT_ARI)
                If bRet = False Then
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                    Exit Sub
                End If
            End If
        'EG30 V32.1.0.1 ADD END
    End Select

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : JPREdit_JikaiInfo
'//  @\ΌΜ  : uσvtΊ
'//  @\Tv  : »έwέθt@Ci©ό)πeLXg\¦·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-19   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(30.1.0.1) 2014-05-01  CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(30.3.0.1) 2014-10-01  CODED   BY [TCC] T.Nakajima
'//                 k€V²όtF[YQΞ yHKRK_Kansi07_003_01z
'//                 EόD@έuπΜσΝΚW[iΦΖ§³Ήι
'//     REVISIONS :(32.1.0.1) 2016-06-16  CODED   BY [TCC] T.Nakajima
'//                 2016Nx{τΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function JPREdit_JikaiInfo() As Boolean

    Dim strFileName             As String                   't@CΌ
    Dim bRet                    As Boolean                  'Φίθl
    Dim lErrCode                As Long                     'G[R[h
    Dim strLineCount            As String                   'sJE^
    
    Dim sWriteDir               As String                   '«έζtH_Ό
    Dim intFileNumber           As Integer                  't@C|C^
    
    Dim ReadFileSettei()        As JIKAIINFO_IMAGE_FILE     '©όξρC[Wt@C
    Dim i                       As Integer
    Dim j                       As Integer
    Dim iCornerIdx              As Integer                  'R[iCfbNX(½ΤΪΜR[i)
    
    Dim fso                     As New FileSystemObject     't@CVXeIuWFNg
    Dim FsoTS                   As TextStream

    Dim strNowShobunrui         As String                   '»έ?WΜ¬ͺήR[h
    Dim strNowKubun             As String                   '»έ?WΜζͺ
    Dim strNowCorner            As String                   '»έ?WΜR[i
    
    'EG30 V32.1.0.1 ADD START
    Dim strEkiSettiBefPath      As String           '»έwέθf[^iΟXOΫΆj
    Dim strGetValue             As String * 64      'DLLΙζΑΔέθ³κι½ίA64Εθ·Ι΅Δ’ι
    Dim strCompValue            As String           'έθliΟXOΫΆj
    Dim strChangeFlg            As String           'ΟXσ
    Dim intValueLen             As Integer          'ζΎ΅½έθlΜ·³
    Dim intGateNo               As Integer          '1`32@
    'EG30 V32.1.0.1 ADD END
    
    'G[[`πιΎ
    On Error GoTo OUTPUT_ERROR
    'ζΚΕwθ³κ½R[iΝέu³κΔ’ι©H
    If pfSettingCheck(True) = False Then
        '·ΧΔ’έuΜR[iA@ΘΜΕG[Ζ·ι
        GoTo OUTPUT_ERROR
    End If
    
    'C[Wt@CΜoΝζ
    sWriteDir = EKI_JPR_GATE_TXTFILE

    'wsxf[^mFi©όjC[Wt@Cμ¬
    bRet = dllGetEkiIniDataJpr(1, EKI_TUDO_CHK_GATE_FILE_JPR, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        'wsxf[^mFi©όjC[Wt@Cν
        Kill EKI_TUDO_CHK_GATE_FILE_JPR
        'ΩνOoΝ
        Call pfOutPutErrLog(lErrCode)
        JPREdit_JikaiInfo = False
        Exit Function
    End If
    
    'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z DEL START
'    'EG20 V30.1.0.1 ADD START
'    '©όβCSVt@Cμ¬
'    bRet = dllGetEkiIniData(2, EKI_TUDO_CHK_SUBGATE_FILE, EKI_SETTI_FILE, lErrCode)
'    If bRet = False Then
'        '©όβCSVt@Cν
'        Kill EKI_TUDO_CHK_SUBGATE_FILE
'        'ΩνOoΝ
'        Call pfOutPutErrLog(lErrCode)
'        JPREdit_JikaiInfo = False
'        Exit Function
'    End If
'    'EG20 V30.1.0.1 ADD END
    'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z DEL END
    
    
    ' R[iΌΜέθ
    Call gsGetCornerName

    'ϊlέθ
    strFileName = ""

    '----------------------------------------------------
    '»έwέθt@Cυ
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    't@CͺΆέ΅Θ’κ
    If strFileName = "" Then
    
        'ΩνOoΝ
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        'ΩνIΉ
        JPREdit_JikaiInfo = False
        Exit Function
        
    End If

    'wsxf[^(©ό)C[Wt@CΜπζΎ
    't@CΤζΎ
    intFileNumber = FreeFile
    Open EKI_TUDO_CHK_GATE_FILE_JPR For Input As #intFileNumber
    
    'CSVt@CsJEgit@CI[άΕ[vπJθΤ·j
        Do While Not EOF(intFileNumber)                     ' EG20 V3.3.0.1ΗΑ
            Line Input #intFileNumber, strLineCount
            j = j + 1
        Loop
    
    'CSVt@CN[Y
    Close #intFileNumber

    't@CΤζΎ
    intFileNumber = FreeFile

    'Δέθ
    ReDim ReadFileSettei(j) As JIKAIINFO_IMAGE_FILE        't@CΗpGA
        
    'CSVt@CI[v
    Open EKI_TUDO_CHK_GATE_FILE_JPR For Input As #intFileNumber

    'Xg\¦ͺΗέέit@CI[άΕ[vπJθΤ·j
        For i = 0 To j - 1
            Input #intFileNumber, ReadFileSettei(i).strBunrui_Dai, ReadFileSettei(i).strBunrui_Tyu, _
            ReadFileSettei(i).srtBunrui_Sho, ReadFileSettei(i).strCorner, ReadFileSettei(i).strKomoku, _
            ReadFileSettei(i).strKubun, ReadFileSettei(i).strData, ReadFileSettei(i).strSetShosai
        Next

    'CSVt@CN[Y
    Close #intFileNumber
    
    'EG30 V32.1.0.1 ADD START
    '»ΜR[iΜΟXOf[^ΫΆ³κ½f[^πγΙWJ·ι(R[i0j
    strEkiSettiBefPath = Replace(EKI_SETTI_FILE_BEF, "#", "0")
    Call dllGetEkiIniDataBefore(strEkiSettiBefPath, lErrCode)
    'EG30 V32.1.0.1 ADD END
    
    '///////////////////////////////////////
    '// W[ioΝC[Wt@Cπμ¬
    '///////////////////////////////////////
    '’gpΜt@CΤζΎ
    intFileNumber = FreeFile
    
    'W[ioΝC[Wt@Cπμ¬
    Open sWriteDir For Output As #intFileNumber
    
    '^Cg\¦
    'PrintHeader intFileNumber, "wsxf[^mF"  'EG30 V32.1.0.1 DEL
    PrintHeader3 intFileNumber, "wsxf[^mF", pfGetSaveDate(0)
    Print #intFileNumber, "έuwF" & Trim(pfGetEkiNameInfo(NotEkiVer))
    
    strNowShobunrui = ""
    strNowKubun = ""
    strNowCorner = ""
    
    For i = 0 To UBound(ReadFileSettei) - 1
        'R[iͺΨθΦνΑ½©H
        If (ReadFileSettei(i).strCorner <> strNowCorner) Then
            'iCornerIdx = iCornerIdx + 1
            'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z DEL START
'            'EG20 V30.1.0.1 ADD START
'            '‘R[ioΝΕ2R[iΪΘ~ͺ ικA2R[iΪΙόιOΙ©όβπoΝ
'            If strNowCorner <> "" Then
'                pfOutPutSubGate CInt(strNowCorner), intFileNumber
'            End If
'            'EG20 V30.1.0.1 ADD END
            'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z DEL END
            'έuR[iπoΝ
            If IsTaisyoCorner(CInt(ReadFileSettei(i).strCorner)) = True Then
                
                'ΞΫR[iΕ ΑΔΰΞΫ@ͺΘ’©ΰ΅κΘ’
                For j = 0 To 15
                    If IsTaisyoGoki(CInt(ReadFileSettei(i).strCorner), j + 1) = True Then
                        If i <> 0 Then
                            Print #intFileNumber, ""
                        End If
                        Print #intFileNumber, "έuR[iF" & gstrCornerName(CInt(ReadFileSettei(i).strCorner) - 1)
                        Exit For
                    End If
                Next j
            End If
            strNowCorner = ReadFileSettei(i).strCorner
        End If
    
        '»Μ@ΝoΝΞΫ©H
        If IsTaisyoGoki(CInt(ReadFileSettei(i).strCorner), CInt(ReadFileSettei(i).strBunrui_Tyu)) = True Then
            '¬ͺήΖζͺͺκv΅Θ―κΞ^CgπoΝ·ι
            If (ReadFileSettei(i).srtBunrui_Sho <> strNowShobunrui) Or (ReadFileSettei(i).strKubun <> strNowKubun) Then
                '^CgπoΝ
                Print #intFileNumber, ""
                'Print #intFileNumber, "y" & ReadFileSettei(i).strKomoku & "z" & ReadFileSettei(i).strKubun   'EG30 V32.1.0.1 DEL
                Print #intFileNumber, "@y" & ReadFileSettei(i).strKomoku & "z" & ReadFileSettei(i).strKubun  'EG30 V32.1.0.1 ADD
                strNowShobunrui = ReadFileSettei(i).srtBunrui_Sho
                strNowKubun = ReadFileSettei(i).strKubun
            End If
            'e@ΜέθπoΝ
            
            'EG30 V32.1.0.1 ADD START
            'W[i?Wf[^t@CΜͺήΝR[i@ΤͺZbg³κA³ηΙR[iΤΰZbg³κΔ’ιͺA
            'δrθΖΘιEKI_SETTI.CSVΝͺήΝP`RQΕR[iΤΝOΖΘΑΔ’ι½ίAR[i@ΤπP`RQΙΟ··ι
            If pfCornerGokiToGateNo(CInt(ReadFileSettei(i).strCorner), CInt(ReadFileSettei(i).strBunrui_Tyu), intGateNo) = True Then
            
                'ΟXOf[^ΫΆ³κ½έθlΖδr·ι(εͺήͺόD@ΜκΝAR[iΝ0ΕθΕυj
                bRet = dllGetEkiInfoValue(CInt(ReadFileSettei(i).strBunrui_Dai), _
                                            intGateNo, _
                                            CInt(ReadFileSettei(i).srtBunrui_Sho), _
                                            0, _
                                            strGetValue, _
                                            intValueLen)
                strCompValue = strGetValue
                If (intValueLen <> 0) Then
                    strCompValue = MidByte(strGetValue, 1, intValueLen)
                    strCompValue = Trim(strCompValue)
                ElseIf (intValueLen = 0) Then
                    strCompValue = ""
                End If
                
                If (bRet = False) Or (ReadFileSettei(i).strData <> strCompValue) Then
                    strChangeFlg = DIFF_MARK_STRING_ON
                Else
                    strChangeFlg = DIFF_MARK_STRING_OFF
                End If
            'δrθΜ@ͺ’Θ©Α½ηuv
            Else
                strChangeFlg = DIFF_MARK_STRING_ON
            End If
            'EG30 V32.1.0.1 ADD END
            
            '@Τπ\¦·ιΪΝ@Τπ99`?ΙΟ··ι
            If (ReadFileSettei(i).strBunrui_Dai = 4 And ReadFileSettei(i).srtBunrui_Sho = 5) Or _
               (ReadFileSettei(i).strBunrui_Dai = 4 And ReadFileSettei(i).srtBunrui_Sho = 7) Then
                ReadFileSettei(i).strData = Format(ReadFileSettei(i).strData, "0#")
            End If
                
            'Print #intFileNumber, ReadFileSettei(i).strBunrui_Tyu & "@ " & ReadFileSettei(i).strData    'EG30 V32.1.0.1 DEL
            Print #intFileNumber, strChangeFlg & ReadFileSettei(i).strBunrui_Tyu & "@ " & ReadFileSettei(i).strData    'EG30 V32.1.0.1 ADD
            'bJprFlg = True
        End If
    Next i
    Print #intFileNumber, ""
    
    'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z DEL START
'    'EG20 V30.1.0.1 ADD START
'    '©όβπoΝ
'    pfOutPutSubGate CInt(strNowCorner), intFileNumber
'    Print #intFileNumber, ""
'    'EG20 V30.1.0.1 ADD END
    'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z DEL END
    
    Print #intFileNumber, FOOTER_STRING
    't@CπN[Y·ιB
    Close #intFileNumber
    Set fso = Nothing
    JPREdit_JikaiInfo = True
    Exit Function

OUTPUT_ERROR:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    Set fso = Nothing
    'ΩνOoΝ
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JPR_PRINT_OUTPUT_ERR, 0)
    
    JPREdit_JikaiInfo = False
End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  ΦΌΜ  : JPREdit_SubGateInfo
'//  @\ΌΜ  : uσvtΊ
'//  @\Tv  : »έwέθt@CiGR[hR[i@ξρθ`)πeLXg\¦·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(30.3.0.1) 2014-10-01  CODED   BY [TCC] T.Nakajima
'//                 k€V²όtF[YQΞ yHKRK_Kansi07_003_01z
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function JPREdit_SubGateInfo() As Boolean

    Dim strFileName             As String                   't@CΌ
    Dim bRet                    As Boolean                  'Φίθl
    Dim lErrCode                As Long                     'G[R[h
    Dim strLineCount            As String                   'sJE^
    
    Dim sWriteDir               As String                   '«έζtH_Ό
    Dim intFileNumber           As Integer                  't@C|C^
    
    Dim ReadFileSettei()        As JIKAIINFO_IMAGE_FILE     '©όξρC[Wt@C
    Dim i                       As Integer
    Dim j                       As Integer
    Dim iCornerIdx              As Integer                  'R[iCfbNX(½ΤΪΜR[i)
    
    Dim fso                     As New FileSystemObject     't@CVXeIuWFNg
    Dim FsoTS                   As TextStream

    Dim strNowShobunrui         As String                   '»έ?WΜ¬ͺήR[h
    Dim strNowKubun             As String                   '»έ?WΜζͺ
    Dim strNowCorner            As String                   '»έ?WΜR[i
    
    'G[[`πιΎ
    On Error GoTo OUTPUT_ERROR
    
    'C[Wt@CΜoΝζ
    sWriteDir = EKI_JPR_SUBGATE_TXTFILE

    '©όβCSVt@Cμ¬
    bRet = dllGetEkiIniData(2, EKI_TUDO_CHK_SUBGATE_FILE, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        '©όβCSVt@Cν
        Kill EKI_TUDO_CHK_SUBGATE_FILE
        'ΩνOoΝ
        Call pfOutPutErrLog(lErrCode)
        JPREdit_SubGateInfo = False
        Exit Function
    End If
    
    'ϊlέθ
    strFileName = ""

    '----------------------------------------------------
    '»έwέθt@Cυ
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    't@CͺΆέ΅Θ’κ
    If strFileName = "" Then
    
        'ΩνOoΝ
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        'ΩνIΉ
        JPREdit_SubGateInfo = False
        Exit Function
        
    End If

    '///////////////////////////////////////
    '// W[ioΝC[Wt@Cπμ¬
    '///////////////////////////////////////
    '’gpΜt@CΤζΎ
    intFileNumber = FreeFile
    
    'W[ioΝC[Wt@Cπμ¬
    Open sWriteDir For Output As #intFileNumber
    
    '^Cg\¦
    PrintHeader2 intFileNumber, "wsxf[^mF", "(GR[hR[i@ξρθ`)"
    Print #intFileNumber, "έuwF" & Trim(pfGetEkiNameInfo(NotEkiVer))
    
    strNowShobunrui = ""
    strNowKubun = ""
    strNowCorner = ""
    
    If pfOutPutSubGate(0, intFileNumber) = False Then
        GoTo OUTPUT_ERROR
    End If
    Print #intFileNumber, ""
    
    Print #intFileNumber, FOOTER_STRING
    't@CπN[Y·ιB
    Close #intFileNumber
    Set fso = Nothing
    JPREdit_SubGateInfo = True
    Exit Function

OUTPUT_ERROR:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    Set fso = Nothing
    'ΩνOoΝ
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JPR_PRINT_OUTPUT_ERR, 0)
    
    JPREdit_SubGateInfo = False
End Function
'EG30 V32.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  ΦΌΜ  : JprEdit_GateCfg
'//  @\ΌΜ  : όD@Ϋηέθf[^W[iC[Wt@Cμ¬
'//  @\Tv  : όD@Ϋηέθf[^W[iC[Wt@Cπμ¬·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Boolean   bSkipFlg  όD@Ϋηf[^ͺ³’½ίAW[i?WπXLbv΅½@ͺ ιB
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-10   CODED   BY [TCC] T.Nakajima
'//             2016Nx{τΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_GateCfg(ByRef bSkipFlg As Boolean) As Boolean

    Dim strOutputFile As String         'oΝt@C
    Dim lngRet As Long                  'ΦΤθl
    Dim lngErrCode As Long              'G[R[h
    Dim iOutFile    As Integer          't@CΤ
    Dim ReadFileGateCfg()    As GATE_CFG_DATA_FILE  'όD@Ϋηέθf[^
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim l           As Integer
    Dim strJpCfgPath            As String  '@ΚέθRtBOt@C(JPp)
    Dim strSetteiBefFolder      As String  'ΔΥΟXOΫΆΜζ
    Dim strJpCfgPathBef         As String  'ΟXOΫΆ΅½@ΚέθRtBOt@CΌ
    Dim strDispImageFileName    As String  '?Wf[^t@CΌ
    Dim objFs                   As New FileSystemObject
    Dim intFileNo               As Integer
    
    Dim blnRet                  As Boolean  '?Wf[^μ¬Φίθl
    
    Dim strMutexName    As String       '~[ebNXΌ
    
    Dim strNowInfoName  As String       '»έoΝΜξρΌ
    Dim strNowDai       As String       '»έoΝΜεΪ
    Dim strNowChu       As String       '»έoΝΜΪ
    
    Dim iKoumokuByte          As Integer 'ΪΌΜoCg
    Dim iValueByte          As Integer 'έθlΜoCg
    Dim iSpaceByte          As Integer 'ΤΙ}ό·ιXy[XΜoCg
    Dim strChangeFlg        As String  'ΟXtO
    Dim strSyoName          As String  '¬ΪΌ
    Dim strValue            As String  'έθl
    Dim blnInfoNameFlg      As Boolean 'όstOiξρΌΌγΜεΪ|ΪΌΜΌOΝόsΘ΅)
    Dim intOutCount         As Integer  'oΝΒ\@
    Dim intOutCountbyCorner(0 To 5) As Integer 'oΝΒ\@iR[ij
    Dim intGateNo           As Integer  '1`32@
    Dim bResult(0 To 5, 0 To 15) As Boolean  'W[i?Wf[^t@CoΝΚ
    Dim strExistsCheckFileName As String     '0101.CSV`0616.CSVάΕΜt@CΜΆέπ`FbN
    Dim bDelFlg                 As Boolean  'νtO
    Const COLON_LEN = 2                 'uFvΜoCg
    
    On Error GoTo Err_handler
    bDelFlg = False
    intOutCount = 0
    
    'έuw
    gsGetStationName
    '©όξρ
    gsGetGateInfo
    'R[iΌ
    gsGetCornerName
    'R[i^Cv
    gsGetCornerType
    
    'ζΚΕwθ³κ½R[iΝέu³κΔ’ι©H
    If pfSettingCheck(True) = False Then
        '·ΧΔ’έuΜR[iA@ΘΜΕG[Ζ·ι
        GoTo Err_handler
    End If
    
    'oΝt@CΌ?W
    strOutputFile = GATE_CFG_TXTFILE
    
    'W[i?Wf[^t@Cπ·ΧΔν·ι
    'νt@CͺΆέ΅Θ’κΝErr_HandlerΙ’ΑΔ΅ά€½ίAΆέ`FbNπs€B
    't@CͺκΒΕΰ©Β©κΞAChJ[hΙζιt@CνͺΕ«ιΜΕA[vπ²―ι
    strDispImageFileName = Replace(EDIT_DATA_GATECFG, "####", "*") & ".csv"
    For i = 1 To 6
        For j = 1 To 16
            strExistsCheckFileName = Replace(EDIT_DATA_GATECFG, "####", Format(i, "0#") & Format(j, "0#")) & ".csv"
            If objFs.FileExists(strExistsCheckFileName) Then
                objFs.DeleteFile strDispImageFileName
                bDelFlg = True
                Exit For
            End If
        Next j
        'JP_CFG*.CSVΕνΟέ
        If bDelFlg = True Then
            Exit For
        End If
    Next i
    
    bSkipFlg = False
    
    't@CoΝΦπCall
    '`FbN³κΔ’ιR[iA@ͺΙΒ’Δ
    For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        intOutCountbyCorner(i) = 0
        For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
            '»ΜR[iA@Νέu³κΔ’ι©H
            If pfCornerGokiToGateNo(udtJprPrintSetteingInfo.iCorner(i), udtJprPrintSetteingInfo.iGouki(j), intGateNo) = True Then
                
                strDispImageFileName = Replace(EDIT_DATA_GATECFG, "####", _
                    Format(udtJprPrintSetteingInfo.iCorner(i), "0#") & Format(udtJprPrintSetteingInfo.iGouki(j), "0#")) & ".csv"
                
                '~[ebNXΌπμ¬
                strMutexName = Replace(MU_N_CFG, "##", Format(intGateNo, "0#"))

                strJpCfgPath = PATH_DATA & Replace(JP_CFG, "##", Format(intGateNo, "0#"))
                strSetteiBefFolder = PATH_OPERATE & "CORNER" & udtJprPrintSetteingInfo.iCorner(i) & "\\SETTEI_BEF\\"
                strJpCfgPathBef = strSetteiBefFolder & Replace(JP_CFG, "##", Format(intGateNo, "0#"))
                
                '³f[^(JP_CFGnn.GAT)ͺΆέ΅½κΝAW[if[^t@Cπμ¬
                If objFs.FileExists(strJpCfgPath) = True Then
                    bResult(i, j) = dllCreateGateCfgData(gintCornerType(udtJprPrintSetteingInfo.iCorner(i) - 1), _
                                                strDispImageFileName, strJpCfgPath, strJpCfgPathBef, strMutexName, lngErrCode)
                    If bResult(i, j) <> False Then
                        'W[i?Wf[^t@CͺκΒΘγμκΔ’κΞAΌΜ@ΕΈs΅ΔΰσόΒ\Μ½ίB
                        intOutCount = intOutCount + 1
                        intOutCountbyCorner(i) = intOutCountbyCorner(i) + 1
                    Else
                        'eLXgμ¬ΈsΙζθXLbv
                        bSkipFlg = True
                    End If
                Else
                    ' uόD@Ϋηέθf[^πσΕ«Θ©Α½όD@ͺ θά·vπ\¦·ι½ίΙON
                    bSkipFlg = True
                End If
                
            End If
        Next j
    Next i
    
    'W[i?Wf[^t@CͺμκΔ’κΞAW[ioΝΒ\
    If intOutCount > 0 Then
        'όD@Ϋηέθf[^ W[iC[Wt@Cπμ¬
        iOutFile = FreeFile
        Open strOutputFile For Output As #iOutFile
        
        'wb_[
        PrintHeader iOutFile, "όD@Ϋηέθf[^mF"
        
        'έuw
        Print #iOutFile, "έuwF" & gstrStationName(0)
        
        '`FbN³κ½R[iͺ[v
        For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
            Erase ReadFileGateCfg
            
            '»ΜR[iΜόD@ͺ·ΧΔόD@Ϋηέθf[^πΑΔ’Θ’κΝσ΅Θ’
            If intOutCountbyCorner(i) > 0 Then
                'R[iΌ
                Print #iOutFile, "έuR[iF" & gstrCornerName(udtJprPrintSetteingInfo.iCorner(i) - 1)
                'ΫΆϊ
                Print #iOutFile, "ΫΆϊF" & pfGetSaveDate(udtJprPrintSetteingInfo.iCorner(i))
    
                For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
                    ' W[i?Wf[^t@Cμ¬ͺ³νΜκΝ
                    If bResult(i, j) <> False Then
                        '»Μ@ͺέu³κΔ’ι©H
                        intFileNo = FreeFile
                        strDispImageFileName = Replace(EDIT_DATA_GATECFG, "####", _
                            Format(udtJprPrintSetteingInfo.iCorner(i), "0#") & Format(udtJprPrintSetteingInfo.iGouki(j), "0#")) & ".csv"
                    
                        'W[i?Wf[^t@CπI[vi±Μt@CͺΆέ΅Θ’κΝ±±ΙΝΘ’j
                        Open strDispImageFileName For Input As #intFileNo
                
                        'ζΚ\¦pf[^(csv)πGAΙΗέή
                        k = 0
                        Do While Not EOF(intFileNo)
                            ReDim Preserve ReadFileGateCfg(k)
                            Input #intFileNo, _
                                    ReadFileGateCfg(k).strInfoName, _
                                    ReadFileGateCfg(k).strBunrui_Dai, ReadFileGateCfg(k).strBunrui_Chu, _
                                    ReadFileGateCfg(k).strBunrui_Syo, ReadFileGateCfg(k).strValue, ReadFileGateCfg(k).strChangeFlg
                            k = k + 1
                        Loop
                        't@CN[Y
                        Close #intFileNo
                        
                        '@Τ
                        Print #iOutFile, "@ΤF" & Format(udtJprPrintSetteingInfo.iGouki(j), "00") & "@"
                        
                        '±±©η1δͺΜόD@Ϋηέθf[^Μΰe({Ά)πσ·ι[v
                        strNowInfoName = ""
                        strNowDai = ""
                        strNowChu = ""
                        blnInfoNameFlg = False
                        For l = 0 To UBound(ReadFileGateCfg)
                            'ξρΌͺΩΘκΞAζΨθ^CgπoΝ·ι αξρΌβ
                            If strNowInfoName <> ReadFileGateCfg(l).strInfoName Then
                                Print #iOutFile, ""
                                Print #iOutFile, "@" & ReadFileGateCfg(l).strInfoName
                                strNowInfoName = ReadFileGateCfg(l).strInfoName
                                blnInfoNameFlg = True
                            End If
                            'εͺήAͺήͺΩΘικΝAζΨθ^CgπoΝ·ι yεͺή-ͺήz
                            If strNowDai <> ReadFileGateCfg(l).strBunrui_Dai Or strNowChu <> ReadFileGateCfg(l).strBunrui_Chu Then
                                'ζΨθ^Cg(εͺή|ͺή)ΜΌOΝ1sόs³ΉιͺAξρΌΜΌγΝόsΘ΅
                                If blnInfoNameFlg = False Then
                                    Print #iOutFile, ""
                                Else
                                    blnInfoNameFlg = False
                                End If
                                Print #iOutFile, "@y" & ReadFileGateCfg(l).strBunrui_Dai & "|" & ReadFileGateCfg(l).strBunrui_Chu & "z"
                                strNowDai = ReadFileGateCfg(l).strBunrui_Dai
                                strNowChu = ReadFileGateCfg(l).strBunrui_Chu
                            End If
                            
                            'ΟXtO { ΪΌ { ":" { έθlπoΝ
                            '½ΆXy[XπΤΙόκι©H
                            strSyoName = RTrim(ReadFileGateCfg(l).strBunrui_Syo)
                            strValue = RTrim(ReadFileGateCfg(l).strValue)
                            iKoumokuByte = LenB(StrConv(strSyoName, vbFromUnicode))
                            iValueByte = LenB(StrConv(strValue, vbFromUnicode))
                            'W[i1sͺΕε30oCg
                            iSpaceByte = MAX_JPR_KETA_MAX - DIFF_MARK_LEN - iKoumokuByte - COLON_LEN - iValueByte
                            If iSpaceByte <= 0 Then
                                iSpaceByte = 0
                            End If
                            If ReadFileGateCfg(l).strChangeFlg = "" Then
                                strChangeFlg = DIFF_MARK_STRING_OFF
                            Else
                                strChangeFlg = DIFF_MARK_STRING_ON
                            End If
                            
                            Print #iOutFile, strChangeFlg & strSyoName & Space(iSpaceByte) & "F" & strValue
                        
                        Next l
                        
                        Print #iOutFile, ""
                    End If
                Next j
            End If
        Next i
        
        Print #iOutFile, FOOTER_STRING
        Close #iOutFile
      
        JprEdit_GateCfg = True
    Else
        'oΝΞΫΖΘιόD@ͺ1δΰΆέ΅Θ’ΜΕAXLbvΝ³΅Ζ·ιB
        bSkipFlg = False
        JprEdit_GateCfg = False
    End If
    Exit Function

Err_handler:

    If intFileNo > 0 Then
        Close #intFileNo
    End If
    
    If iOutFile > 0 Then
        Close #iOutFile
    End If
    
    Set objFs = Nothing

    'MsgBox "ΩνIΉ΅ά΅½B", vbCritical, "oΝΚ"
    'uW[iσζΚiόD@Ϋηέθf[^jFW[iC[Wt@Cμ¬ΩνvOoΝ
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, lngErrCode)
    JprEdit_GateCfg = False

End Function
'EG30 V32.1.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : IsTaisyoGoki
'//  @\ΌΜ  : wθ@mF
'//  @\Tv  : C[Wt@CΙoΝ·ιΪΝoΝΞΫ©mF·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Integer   iCorner   R[iΤ
'//              Integer   iGouki    @Τ
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-19   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function IsTaisyoGoki(iCorner As Integer, iGouki As Integer) As Boolean
    Dim bRet        As Boolean
    Dim i           As Integer
    Dim j           As Integer
    
    bRet = False
    
    If pfCornerGokiCheck(iCorner, iGouki) = False Then
        '’έuΜ@ΘΜΕAζΚγ`FbN³κΔ’Δΰfalse
        IsTaisyoGoki = False
        Exit Function
    End If
    
    
    For j = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        For i = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
            If udtJprPrintSetteingInfo.iCorner(j) = iCorner Then
                If udtJprPrintSetteingInfo.iGouki(i) = iGouki Then
                    bRet = True
                    Exit For
                End If
            End If
        Next i
    Next j
    
    IsTaisyoGoki = bRet
   
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : IsTaisyoCorner
'//  @\ΌΜ  : wθR[imF
'//  @\Tv  : C[Wt@CΙoΝ·ιΪΝoΝΞΫ©mF·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Integer   iCorner   R[iΤ
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-19   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function IsTaisyoCorner(iCorner As Integer) As Boolean
    Dim bRet        As Boolean
    Dim i           As Integer
    
    bRet = False
    '»ΜR[iΝέu³κΔ’ι©H
    If pfCornerGokiCheck(iCorner) = True Then
        
        For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
            If udtJprPrintSetteingInfo.iCorner(i) = iCorner Then
                bRet = True
                Exit For
            End If
        Next i
    End If
    
    IsTaisyoCorner = bRet
   
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : JprEdit_SetteiList
'//  @\ΌΜ  : έθlκoΝ
'//  @\Tv  : έθlκW[iΜC[Wt@Cπ?W·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Integer   iCorner   R[iΤ
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-19   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(7.4.0.1) 2013-07-22   REVISED BY [TCC] T.Nakajima
'//                ϊά½ͺθoκt[έθζΚΞ
'//     REVISIONS :(EG30 V32.1.0.1) 2016-06-17   REVISED BY [TCC] T.Nakajima
'//                2016Nx{τΞ
'//     REVISIONS :(EG30 V35.3.0.1) 2019-07-03   REVISED BY [TCC] H.Kondoh
'//                2019Nx{τΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_SetteiList() As Boolean
    Dim strFilePath As String           'oΝt@CpX
    Dim intCount As Integer             'JE^
    Dim intOutFile As Integer           'oΝt@CΤ
    Dim intTgtFileNo As Integer         'oΝΞΫέθt@CΤ
    Dim strTgtFileName As String        'oΝΞΫέθt@C
    Dim strTargetFile() As String       'oΝΞΫt@C
    Dim intFileNum As Integer
    Dim objFileObj As FileSystemObject  't@CVXeIuWFNg
    Dim ReadFileSettei()    As SETTEI_OUTPUT_IMAGE_FILE   't@CΗp\’Μ
    Dim strCsvBuffer        As String
    Dim strCammaArray()     As String
    Dim i As Integer
    Dim strNowDaikomoku     As String
    Dim strNowKomoku        As String
    Dim FsoTS   As TextStream
    Dim iKomkuByte          As Integer 'ΪΌΜoCg
    Dim iValueByte          As Integer 'έθlΜoCg
    Dim iSpaceByte          As Integer 'ΤΙ}ό·ιXy[XΜoCg
    Dim intJprFile            As Integer
    Dim strNyujoFree(3)       As String
    Dim iSeparatePos          As Integer    'ζΚΌΜͺu:vΕζΨηκΔ’½κΜζΨθΚu
    'EG30 V32.1.0.1 ADD START
    Dim strChangeFlg        As String  'ΟXtO
    'EG30 V32.1.0.1 ADD END

    Set objFileObj = New FileSystemObject
    
    On Error GoTo Err_handler
    
    'EG20 V30.1.0.1 ADD START
    'έuw
    gsGetStationName
    '©όξρ
    gsGetGateInfo
    'R[iΌ
    gsGetCornerName
    'R[i^Cv
    gsGetCornerType
    'EG20 V30.1.0.1 ADD END
    'ζΚΕwθ³κ½R[iΝέu³κΔ’ι©H
    If pfSettingCheck(False) = False Then
        '·ΧΔ’έuΘΜΕG[
        GoTo Err_handler
    End If
    
    'oΝΞΫέθt@CπI[v·ιB
    intTgtFileNo = FreeFile
    strTgtFileName = OUTPUT_TARGET_FILE
    
    'oΝΞΫέθt@CͺΆέ΅Θ’κΝΩνIΉ
    If objFileObj.FileExists(strTgtFileName) = False Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_SEARCH_ERROR & ":" & strTgtFileName, 0)
        GoTo Err_handler
    End If
    
    Open strTgtFileName For Input As #intTgtFileNo
    
    'oΝΞΫt@CπζΎ
    Input #intTgtFileNo, intFileNum
    
    'oΝΞΫt@CπζΎ
    ReDim strTargetFile(intFileNum - 1)
    For intCount = 0 To UBound(strTargetFile)
        Input #intTgtFileNo, strTargetFile(intCount)
    Next
    
    Close #intTgtFileNo
    
    'EG20 V30.1.0.1 ADD START
    '²όR[i[ΙΞ·ιoΝΞΫt@CΜΰeπmΫ·ι
    intTgtFileNo = FreeFile
    strTgtFileName = OUTPUT_TARGET_FILE_KAN
    
    'oΝΞΫέθt@CͺΆέ΅Θ’κΝΩνIΉ
    If objFileObj.FileExists(strTgtFileName) = False Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_SEARCH_ERROR & ":" & strTgtFileName, 0)
        GoTo Err_handler
    End If
    
    Open strTgtFileName For Input As #intTgtFileNo
    
    'oΝΞΫt@CπζΎ
    Input #intTgtFileNo, intFileNum
    
    'oΝΞΫt@CπζΎ
    ReDim strTargetFileKan(intFileNum - 1)
    For intCount = 0 To UBound(strTargetFileKan)
        Input #intTgtFileNo, strTargetFileKan(intCount)
    Next
    
    Close #intTgtFileNo
    'EG20 V30.1.0.1 ADD END
    
    '////////////////////////////////
    'W[iC[Wt@Cμ¬
    'OρΜoΝΟέΜW[iC[Wt@CΝΑ΅Δ¨­(R[iPΚΕΗL΅Δ’­½ί)
    If Dir(SETTI_TXTFLE) <> "" Then
        Kill SETTI_TXTFLE
    End If
    
    'R[iPΚΕέθlκΜCSVt@Cπμ¬·ι
    'wb_μ¬
    intJprFile = FreeFile
    Open SETTI_TXTFLE For Output As #intJprFile
    PrintHeader intJprFile, "έθlκ"
    
    'έuw
    'gsGetStationName   'EG20 V30.1.0.1 DEL
    Print #intJprFile, "έuwF" & gstrStationName(0)
    'R[iΌ
    'gsGetCornerName    'EG20 V30.1.0.1 DEL

    For intCount = 0 To UBound(glngTergetCorner)
        
        If glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON Then
            'R[iPΚΕέθt@Cκ(?WpCSV)μ¬ OPERATE_SETTI99.csv
            strFilePath = Replace(EDIT_DATA_SETTEI, "##", Format(intCount + 1, "0#"))
            
            '---- έθκeLXgμ¬ Jn
            't@Cμ¬
            If objFileObj.FileExists(strFilePath) = True Then
                objFileObj.DeleteFile (strFilePath)
            End If
            Call objFileObj.CreateTextFile(strFilePath)
            
            'oΝt@CπI[v·ιB
            intOutFile = FreeFile
            Open strFilePath For Output As #intOutFile
    
            'IDέθlπoΝ
            'If gsubOutput_Id(intCount + 1, intOutFile, True) = False Then      'EG30 V32.1.0.1 DEL
            If gsubOutput_Id_JPR(intCount + 1, intOutFile, True) = False Then   'EG30 V32.1.0.1 ADD
                GoTo Err_handler
            End If
            
            'EG20 V30.1.0.1 DEL START
            'όoκt[t@CπoΝ
'            If gsubOutput_Free_InOut(intCount + 1, intOutFile) = False Then
'                GoTo Err_Handler
'            End If
'
'            'jΥϊt@CπoΝ
'            If gsubOutput_Shukusai(intCount + 1, intOutFile) = False Then
'                GoTo Err_Handler
'            End If
            'EG20 V30.1.0.1 DEL END
            
            'EG20 V30.1.0.1 ADD START
            If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                '²όR[i[Μκ
                'V²όs³p[^πoΝ
                'If gsubOutput_ParaKan(FILE_PR_FSK, OUTPUT_PRFSK_FILE, intCount + 1, intOutFile) = False Then   'EG30 V32.1.0.1 DEL
                If gsubOutput_ParaKan_JPR(FILE_PR_FSK, OUTPUT_PRFSK_FILE, intCount + 1, intOutFile) = False Then    'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
                'έόIC»θp[^πoΝ
                'If gsubOutput_ParaKan(FILE_PR_ICZ, OUTPUT_PRICZ_FILE, intCount + 1, intOutFile) = False Then   'EG30 V32.1.0.1 DEL
                If gsubOutput_ParaKan_JPR(FILE_PR_ICZ, OUTPUT_PRICZ_FILE, intCount + 1, intOutFile) = False Then    'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
                
                'έόICΚίp[^πoΝ
                'If gsubOutput_ParaKan(FILE_PR_ITZ, OUTPUT_PRITZ_FILE, intCount + 1, intOutFile) = False Then   'EG30 V32.1.0.1 DEL
                If gsubOutput_ParaKan_JPR(FILE_PR_ITZ, OUTPUT_PRITZ_FILE, intCount + 1, intOutFile) = False Then    'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
            Else
                'έR[i[Μκ
                'όoκt[t@CπoΝ
                'If gsubOutput_Free_InOut(intCount + 1, intOutFile) = False Then    'EG30 V32.1.0.1 DEL
                If gsubOutput_Free_InOut_JPR(intCount + 1, intOutFile) = False Then     'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
                
                'jΥϊt@CπoΝ
                'If gsubOutput_Shukusai(intCount + 1, intOutFile) = False Then  'EG30 V32.1.0.1 V32.1.0.1 DEL
                If gsubOutput_Shukusai_JPR(intCount + 1, intOutFile) = False Then   'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
            End If
            'EG20 V30.1.0.1 ADD END

            Close #intOutFile
            '---- έθκeLXgμ¬ IΉ
            
            'oΝ΅½?W³f[^πGAΙZbg·ι
            Set FsoTS = objFileObj.OpenTextFile(strFilePath, ForReading)
            i = 0
            Do Until FsoTS.AtEndOfStream = True
                ReDim Preserve ReadFileSettei(i)
                strCsvBuffer = FsoTS.ReadLine
                'J}πL[[hΙeΪπΨθo·B
                strCammaArray = Split(strCsvBuffer, ",")
                ReadFileSettei(i).strDaiKomoku = strCammaArray(0)   'εΪ
                ReadFileSettei(i).strKomoku = strCammaArray(1)      'ΪΌ
                ReadFileSettei(i).strValue = strCammaArray(2)       'έθl
                ReadFileSettei(i).strChangeFlg = strCammaArray(3)   'ΟXtO
                
                i = i + 1
            Loop
            FsoTS.Close
            
            'ΗέρΎGA©ηW[iC[Wt@Cπμ¬·ι
            
            'R[iΌ
            Print #intJprFile, "έuR[iF" & gstrCornerName(intCount)
            'ΫΆϊ
            Print #intJprFile, "ΫΆϊF" & pfGetSaveDate(intCount + 1)

            strNowDaikomoku = ""
            strNowKomoku = ""
            
            For i = 0 To UBound(ReadFileSettei)
                'εΪπoΝ·ι©H
                If strNowDaikomoku <> ReadFileSettei(i).strDaiKomoku Then
                    '½Ύ΅ANULLΜκΝ@Θ~ΘΜΕp±
                    If ReadFileSettei(i).strDaiKomoku <> "" Then
                        'Print #intJprFile, ""  'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
                        If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                            'ΪxΜΨΦΝόs΅Θ’
                            iSeparatePos = InStr(ReadFileSettei(i).strDaiKomoku, ":")
                            If iSeparatePos > 0 Then
                                If Left(strNowDaikomoku, iSeparatePos - 1) <> Left(ReadFileSettei(i).strDaiKomoku, iSeparatePos - 1) Then
                                    'εͺήxΕΩΘΑΔ’ιΜΕόs
                                    Print #intJprFile, ""
                                Else
                                End If
                            Else
                                Print #intJprFile, ""
                            End If
                        Else
                            Print #intJprFile, ""
                        End If
                        
                        'Print #intJprFile, "y" & ReadFileSettei(i).strDaiKomoku & "z"    'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
                        If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                            '²όR[iΜκ
                            'εΪ(ζΚΌΜͺ":"ΕζΨηκΔ’½η»±Εͺ―ι)
                            iSeparatePos = InStr(ReadFileSettei(i).strDaiKomoku, ":")
                            If iSeparatePos > 0 Then
                                'εͺήͺ―ΆΎΑ½ηoΝ΅Θ’
                                If Left(strNowDaikomoku, iSeparatePos - 1) <> Left(ReadFileSettei(i).strDaiKomoku, iSeparatePos - 1) Then
                                    'Print #intJprFile, "y" & Left(ReadFileSettei(i).strDaiKomoku, iSeparatePos - 1) & "z"    'EG30 V32.1.0.1 DEL
                                    Print #intJprFile, "@y" & Left(ReadFileSettei(i).strDaiKomoku, iSeparatePos - 1) & "z"   'EG30 V32.1.0.1 ADD
                                    'ΪπoΝ·ι
                                    'Print #intJprFile, Mid(ReadFileSettei(i).strDaiKomoku, iSeparatePos + 1)   'EG30 V32.1.0.1 DEL
                                    Print #intJprFile, "@" & Mid(ReadFileSettei(i).strDaiKomoku, iSeparatePos + 1) 'EG30 V32.1.0.1 ADD
                                Else
                                    'εͺήάΕΝ―ΆΘΜΕAΪΎ―πoΝ·ιB
                                    'Print #intJprFile, Mid(ReadFileSettei(i).strDaiKomoku, iSeparatePos + 1)   'EG30 V32.1.0.1 DEL
                                    Print #intJprFile, "@" & Mid(ReadFileSettei(i).strDaiKomoku, iSeparatePos + 1) 'EG30 V32.1.0.1 ADD
                                End If
                            Else
                                'Print #intJprFile, "y" & ReadFileSettei(i).strDaiKomoku & "z"    'EG30 V32.1.0.1 DEL
                                Print #intJprFile, "@y" & ReadFileSettei(i).strDaiKomoku & "z"   'EG30 V32.1.0.1 ADD
                            End If
                        Else
                            'Print #intJprFile, "y" & ReadFileSettei(i).strDaiKomoku & "z"    'EG30 V32.1.0.1 DEL
                            Print #intJprFile, "@y" & ReadFileSettei(i).strDaiKomoku & "z"   'EG30 V32.1.0.1 ADD
                        End If
                        'EG20 V30.1.0.1 ADD END
                        strNowDaikomoku = ReadFileSettei(i).strDaiKomoku
                    End If
                End If
                
                'όκt[έθζΚΝέθlπόs³ΉιKvͺ ιB
                If ReadFileSettei(i).strDaiKomoku = "όκt[έθζΚ" Then
                    'όκt[1`6ΜΝSpΙΟX·ιB(dlΙ νΉι½ί)
                    Select Case ReadFileSettei(i).strKomoku
                        Case "όκt[1"
                            'strNyujoFree(0) = "όκt[P"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@όκt[P"    'EG30 V32.1.0.1 ADD
                        Case "όκt[2"
                            'strNyujoFree(0) = "όκt[Q"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@όκt[Q"    'EG30 V32.1.0.1 ADD
                        Case "όκt[3"
                            'strNyujoFree(0) = "όκt[R"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@όκt[R"    'EG30 V32.1.0.1 ADD
                        Case "όκt[4"
                            'strNyujoFree(0) = "όκt[S"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@όκt[S"    'EG30 V32.1.0.1 ADD
                        Case "όκt[5"
                            'strNyujoFree(0) = "όκt[T"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@όκt[T"    'EG30 V32.1.0.1 ADD
                        Case "όκt[6"
                            'strNyujoFree(0) = "όκt[U"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@όκt[U"    'EG30 V32.1.0.1 ADD
                        Case "όκt[7"
                            'strNyujoFree(0) = "όκt[V"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@όκt[V"    'EG30 V32.1.0.1 ADD
                        Case "όκt[8"
                            'strNyujoFree(0) = "όκt[W"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@όκt[W"    'EG30 V32.1.0.1 ADD
                        Case "όκt[9"
                            'strNyujoFree(0) = "όκt[X"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@όκt[X"    'EG30 V32.1.0.1 ADD
                    End Select
                    'EG30 V32.1.0.1 DEL START
'                    strNyujoFree(1) = MidByte(ReadFileSettei(i).strValue, 1, 8) & Space(6) & MidByte(ReadFileSettei(i).strValue, 11, 16) 'Jnϊ
'                    strNyujoFree(2) = MidByte(ReadFileSettei(i).strValue, 28, 8) & Space(6) & MidByte(ReadFileSettei(i).strValue, 38, 16) 'IΉϊ
'                    strNyujoFree(3) = MidByte(ReadFileSettei(i).strValue, 55, 4) & Space(22) & MidByte(ReadFileSettei(i).strValue, 61, 4) 'ν
                    'EG30 V32.1.0.1 DEL END
                    'EG30 V32.1.0.1 ADD START
                    If ReadFileSettei(i).strChangeFlg = "@" Then
                        strChangeFlg = DIFF_MARK_STRING_OFF
                    Else
                        strChangeFlg = DIFF_MARK_STRING_ON
                    End If
                    strNyujoFree(1) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 1, 8) & Space(4) & MidByte(ReadFileSettei(i).strValue, 11, 16)  'Jnϊ
                    strNyujoFree(2) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 28, 8) & Space(4) & MidByte(ReadFileSettei(i).strValue, 38, 16) 'IΉϊ
                    strNyujoFree(3) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 55, 4) & Space(20) & MidByte(ReadFileSettei(i).strValue, 61, 4) 'ν
                    'EG30 V32.1.0.1 ADD END
                    '1soΝ
                    Print #intJprFile, strNyujoFree(0)
                    Print #intJprFile, strNyujoFree(1)
                    Print #intJprFile, strNyujoFree(2)
                    Print #intJprFile, strNyujoFree(3)
                'EG20 V7.4.0.1 ADD START
                'ϊά½ͺθoκt[έθζΚΝέθπόs³ΉιKvͺ ι
                'ElseIf ReadFileSettei(i).strDaiKomoku = "ϊά½ͺθoκt[έθζΚ" Then    'EG30 V32.1.0.1 DEL
                ElseIf ReadFileSettei(i).strDaiKomoku = "ϊΧθoκt[έθζΚ" Then         'EG30 V32.1.0.1 ADD
                    'όκt[1`6ΜΝSpΙΟX·ιB(dlΙ νΉι½ί)
                    Select Case ReadFileSettei(i).strKomoku
                        'Case "ϊά½ͺθoκt[1"   'EG30 V32.1.0.1 DEL
                        Case "ϊΧθoκt[1"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "ϊά½ͺθoκt[P" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@ϊΧθoκt[P"  'EG30 V32.1.0.1 ADD
                        'Case "ϊά½ͺθoκt[2"   'EG30 V32.1.0.1 DEL
                        Case "ϊΧθoκt[2"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "ϊά½ͺθoκt[Q" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@ϊΧθoκt[Q"  'EG30 V32.1.0.1 ADD
                        'Case "ϊά½ͺθoκt[3"   'EG30 V32.1.0.1 DEL
                        Case "ϊΧθoκt[3"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "ϊά½ͺθoκt[R" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@ϊΧθoκt[R"  'EG30 V32.1.0.1 ADD
                        'Case "ϊά½ͺθoκt[4"   'EG30 V32.1.0.1 DEL
                        Case "ϊΧθoκt[4"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "ϊά½ͺθoκt[S" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@ϊΧθoκt[S"  'EG30 V32.1.0.1 ADD
                        'Case "ϊά½ͺθoκt[5"   'EG30 V32.1.0.1 DEL
                        Case "ϊΧθoκt[5"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "ϊά½ͺθoκt[T" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@ϊΧθoκt[T"  'EG30 V32.1.0.1 ADD
                        'Case "ϊά½ͺθoκt[6"   'EG30 V32.1.0.1 DEL
                        Case "ϊΧθoκt[6"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "ϊά½ͺθoκt[U" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@ϊΧθoκt[U"  'EG30 V32.1.0.1 ADD
                        'Case "ϊά½ͺθoκt[7"   'EG30 V32.1.0.1 DEL
                        Case "ϊΧθoκt[7"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "ϊά½ͺθoκt[V" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@ϊΧθoκt[V"  'EG30 V32.1.0.1 ADD
                        'Case "ϊά½ͺθoκt[8"   'EG30 V32.1.0.1 DEL
                        Case "ϊΧθoκt[8"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "ϊά½ͺθoκt[W" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@ϊΧθoκt[W"  'EG30 V32.1.0.1 ADD
                        'Case "ϊά½ͺθoκt[9"   'EG30 V32.1.0.1 DEL
                        Case "ϊΧθoκt[9"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "ϊά½ͺθoκt[X" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "@ϊΧθoκt[X"  'EG30 V32.1.0.1 ADD
                    End Select
                    'EG30 V32.1.0.1 DEL START
                    'strNyujoFree(1) = MidByte(ReadFileSettei(i).strValue, 1, 8) & Space(6) & MidByte(ReadFileSettei(i).strValue, 11, 16) 'Jnϊ
                    'strNyujoFree(2) = MidByte(ReadFileSettei(i).strValue, 28, 8) & Space(6) & MidByte(ReadFileSettei(i).strValue, 38, 16) 'IΉϊ
                    'strNyujoFree(3) = MidByte(ReadFileSettei(i).strValue, 55, 4) & Space(22) & MidByte(ReadFileSettei(i).strValue, 61, 4) 'ν
                    'EG30 V32.1.0.1 DEL END
                    'EG30 V32.1.0.1 ADD START
                    If ReadFileSettei(i).strChangeFlg = "@" Then
                        strChangeFlg = DIFF_MARK_STRING_OFF
                    Else
                        strChangeFlg = DIFF_MARK_STRING_ON
                    End If
                    strNyujoFree(1) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 1, 8) & Space(4) & MidByte(ReadFileSettei(i).strValue, 11, 16)  'Jnϊ
                    strNyujoFree(2) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 28, 8) & Space(4) & MidByte(ReadFileSettei(i).strValue, 38, 16)  'IΉϊ
                    strNyujoFree(3) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 55, 4) & Space(20) & MidByte(ReadFileSettei(i).strValue, 61, 4) 'ν
                    'EG30 V32.1.0.1 ADD END
                    '1soΝ
                    Print #intJprFile, strNyujoFree(0)
                    Print #intJprFile, strNyujoFree(1)
                    Print #intJprFile, strNyujoFree(2)
                    Print #intJprFile, strNyujoFree(3)
                'EG20 V7.4.0.1 ADD END
                Else
                
                    If ReadFileSettei(i).strKomoku = "" Then
                        'OρΜΪΌπg€
                        ReadFileSettei(i).strKomoku = strNowKomoku
                    End If
                    strNowKomoku = ReadFileSettei(i).strKomoku
                    'αOΪ
                    'ξ{ΝeLXgoΝΜCSVΜΆΎπg€ͺAΊLΜΪΎ―ΝW[idlΙ νΉιKvͺ ιB
                    Select Case ReadFileSettei(i).strKomoku
                        Case "ΤΡJn", _
                             "ΤΡIΉ", _
                             "LψΤΡJn", _
                             "LψΤΡIΉ"
                            'uS@F9999ͺv¨ u S@ 9999ͺvΙΟ··ι
                            ReadFileSettei(i).strValue = Space(1) & Replace(ReadFileSettei(i).strValue, "F", " ")
                        
                        'Case "ΚίT[rXs³Ϋ―"    'EG30 V32.1.0.1 DEL
                        'EG30 V32.1.0.1 ADD START
'EG30 V35.3.0.1 DEL Start
'                        Case "ΚίT[rXs³Ϋ―", _
'                             "ICοΠΤoHA±«", _
'                             "I[g`[W@\", _
'                             "θϊtF[Z[t", _
'                             "ΚtF[Z[t", _
'                             "ICJ[hϊΐO\", _
'                             "ICJ[hϊΐγΔΰ", _
'                             "³l[hΉΊΔΰ", _
'                             "ICΔΰ\¦ζΚ"
'                        'EG30 V32.1.0.1 ADD END
'EG30 V35.3.0.1 ADD End
'EG30 V35.3.0.1 ADD Start
                        Case "ΚίT[rXs³Ϋ―", _
                             "ICοΠΤoHA±«", _
                             "I[g`[W@\", _
                             "θϊtF[Z[t", _
                             "ΚtF[Z[t", _
                             "ICJ[hϊΐO\", _
                             "ICJ[hϊΐγΔΰ", _
                             "³l[hΉΊΔΰ", _
                             "ICΔΰ\¦ζΚ", _
                             "JnNϊ", _
                             "IΉNϊ"
'EG30 V35.3.0.1 ADD End
                            'uΚίT[rXs³Ϋ―S@Fxxv¨ uΚίT[rXs³Ϋ― S@FxxvΙΟ··ι
                            ReadFileSettei(i).strValue = Space(1) & ReadFileSettei(i).strValue
                        
                        Case "³l[h?μέθ"
                            'έθlπΆlΙ·ι
                            ReadFileSettei(i).strValue = ReadFileSettei(i).strValue & Space(8 - LenB(ReadFileSettei(i).strValue))
                    End Select
                
                    'ΪΌoΝ
                    '½ΆXy[XπΤΙόκι©H
                    iKomkuByte = LenB(StrConv(ReadFileSettei(i).strKomoku, vbFromUnicode))
                    iValueByte = LenB(StrConv(ReadFileSettei(i).strValue, vbFromUnicode))
                    'W[i1sͺΕε30oCg
                    'iSpaceByte = MAX_JPR_KETA_MAX - iKomkuByte - iValueByte    'EG30 V32.1.0.1 DEL
                    iSpaceByte = MAX_JPR_KETA_MAX - DIFF_MARK_LEN - iKomkuByte - iValueByte  'EG30 V32.1.0.1 ADD
                    If iSpaceByte < 0 Then
                        'iSpaceByte = 0    'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
                        If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                            iSpaceByte = 1
                        Else
                            iSpaceByte = 0
                        End If
                        'EG20 V30.1.0.1 ADD END
                    ElseIf iSpaceByte = 0 Then
                        'iSpaceByte = 0     'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
                        If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                            iSpaceByte = 1
                        Else
                            iSpaceByte = 0
                        End If
                        'EG20 V30.1.0.1 DEL END
                    End If
                    
                    Space (iSpaceByte)
                    '1soΝ
                    'Print #intJprFile, ReadFileSettei(i).strKomoku & Space(iSpaceByte) & ReadFileSettei(i).strValue    'EG30 V32.1.0.1 DEL
                    'EG30 V32.1.0.1 ADD START
                    If ReadFileSettei(i).strChangeFlg = "@" Then
                        strChangeFlg = DIFF_MARK_STRING_OFF
                    Else
                        strChangeFlg = DIFF_MARK_STRING_ON
                    End If
                    Print #intJprFile, strChangeFlg & ReadFileSettei(i).strKomoku & Space(iSpaceByte) & ReadFileSettei(i).strValue
                    'EG30 V32.1.0.1 ADD END
                End If
              
            Next i
            Print #intJprFile, ""
        End If
    Next intCount
    
    Print #intJprFile, FOOTER_STRING
    
    Close #intJprFile
    Set objFileObj = Nothing
    
    JprEdit_SetteiList = True
    Exit Function
    
'G[
Err_handler:

    If intTgtFileNo > 0 Then
        Close #intTgtFileNo
    End If
    If intOutFile > 0 Then
        Close #intOutFile
    End If
    If intJprFile > 0 Then
        Close #intJprFile
    End If
    Set objFileObj = Nothing
    'G[OΜoΝ
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
    JprEdit_SetteiList = False
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : JprEdit_SimekiriOffline
'//  @\ΌΜ  : χΨItCoΝW[i?W
'//  @\Tv  : χΨItCoΝΜC[Wt@Cπ?W·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      :
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.1.0.1) 2014-03-25   CODED   BY [TCC] T.Nakajima
'//                 k€V²όJΖΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_SimekiriOffline() As Boolean
    
    Dim objFso As New FileSystemObject                  ' t@CVXeIuWFNg
    Dim objTs   As TextStream
    Dim bProceed As Boolean                             ' χΨJntO
    Dim nListCnt As Integer                             ' t@Ci[
    Dim szSaveFolder As String                          ' ΫΆζtH_
    Dim szFileName As String                            ' t@CΌ
    Dim iResponse As Integer
    Dim Index       As Integer                          'CfbNX
    Dim iOutFile    As Integer
    
    On Error GoTo ErrorHandler                          ' G[nhΜo^
    
    'EG20 V30.1.0.1 ADD START
    ' R[iΌζΎ
    gsGetCornerName
    ' R[i^CvζΎ
    gsGetCornerType
    
    ' wΌζΎ
    gsGetStationName
    ' EG20 V30.1.0.1 ADD END
    
    '`FbN³κ½R[iΝέu³κΔ’ι©HiΗκ©ΠΖΒΕΰ κΞOK)
    If pfSettingCheck(False) = False Then
        '·ΧΔ’έuΜR[iΘΜΕG[
        GoTo ErrorHandler
    End If
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // ϊ»
    Index = 0
    Erase gOfflineFileList

    ReDim Preserve gOfflineFileList(0)
    bProceed = False
    nListCnt = 0
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // W[iC[Wt@Cμ¬
    'EG20 V30.1.0.1 DEL START
    ' R[iΌζΎ
    'gsGetCornerName
    
    ' wΌζΎ
    'gsGetStationName
    'EG20 V30.1.0.1 DELEND
    
    'W[iC[Wt@CπI[v
    iOutFile = FreeFile
    Open SIMEKIRI_TXTFILE For Output As #iOutFile
    
    'wb_πoΝ
    PrintHeader iOutFile, "χΨItCoΝ"
    
    'έuw/έuR[i
    Print #iOutFile, "έuwF" & gstrStationName(0)
    
    For Index = 0 To UBound(glngTergetCorner)
    
        If glngTergetCorner(Index) = CMN_ONOFF.CMN_ON Then
            
            ' /////////////////////////////////////////////////////////////////////////
            ' // χΨoΝf[^ΝΆέ·ι©HiD:\KANSI\SHUKEI\SEND_DATA\SIMEKIRI##.DATj
            szFileName = Replace(FILENAME_SIMEKIRIDAT, "##", Format(Index + 1, "0#"))
            If objFso.FileExists(szFileName) = True Then              ' t@CΌΜζΎ`FbN
                nListCnt = nListCnt + 1                             ' t@CΜJE^πAbv·ι
                ReDim Preserve gOfflineFileList(nListCnt)           ' t@CΌi[GAπg£·ι
                gOfflineFileList(nListCnt - 1) = szFileName         ' t@CpXπi[
                bProceed = True
            End If
            
                
            If bProceed = False Then
                GoTo ErrorHandler
            End If
            
            ' /////////////////////////////////////////////////////////////////////////
            ' // ?Wf[^t@Cπμ¬
            ' // R[i²ΖΜχΨeLXgt@Cπμ¬
            bProceed = sOutPutOfflineData(Index)
            If bProceed = False Then
                GoTo ErrorHandler
            End If
            
            Print #iOutFile, "έuR[iF" & gstrCornerName(Index)
            Print #iOutFile, ""
            
            '1R[iͺΜχΨf[^πΗέή
            szFileName = Replace(EDIT_DATA_SIMEKIRI, "##", Format(Index + 1, "0#"))
            Set objTs = objFso.OpenTextFile(szFileName, ForReading)
            Print #iOutFile, objTs.ReadAll
            objTs.Close
            Set objFso = Nothing
        End If
    Next Index
    
    'tb^oΝ
    Print #iOutFile, FOOTER_STRING
    
    Close #iOutFile
    Set objFso = Nothing

    JprEdit_SimekiriOffline = True
    Exit Function

' /////////////////////////////////////////////////////////
' // G[
ErrorHandler:
    'Call MsgBox("ΩνIΉ΅ά΅½B", vbOKOnly, "ItCoΝΚ")
    If iOutFile > 0 Then
        Close #iOutFile
    End If

    Set objFso = Nothing

    'G[OΜoΝ
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 ALL Rights Reserved
'//
'//  ΦΌΜ  : sOutPutOfflineData
'//  @\ΌΜ  : ItCf[^}ΜoΝ
'//  @\Tv  : R[i²ΖΙχΨt@C(eLXgt@C)πμ¬·ιB
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.1.0.1) 2014-03-25  CODED   BY [TCC] T.Nakajima
'//                 k€V²όJΖΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function sOutPutOfflineData(dwCornerIdx As Integer) As Boolean
            
    Dim szFileName As String                            ' t@CΌ
    Dim lResult As Long                                 ' Κ
    Dim dwSequense As Long                              ' V[PXΤ

    ' //////////////////////////////////////////////////////////////
    ' // t@Cμ¬
    ' // QΖ³t@CSIMEKIRI##.DATΜt@CΌπμ¬
    szFileName = Replace(FILENAME_SIMEKIRIDAT, "##", Format(dwCornerIdx + 1, "0#"))
    
    ' //////////////////////////////////////////////////////////////
    ' // R[i²ΖΜχΨf[^(eLXg)πμ¬
    dwSequense = 0                              ' V[PXΤ:0Εθ
    'EG20 V30.1.0.1 DEL START
'    lResult = dllCreateShimekiriFileJpr(dwCornerIdx + 1, dwSequense, _
'                                        PATH_WORK, _
'                                        szFileName)
    'EG20 V30.1.0.1 ADD START
    If gintCornerType(dwCornerIdx) = CORNER_TYPE_KANSEN Then
        '²όR[iΘηΞ²όR[ipΜΦπΔΡo·
        lResult = dllCreateShimekiriFileJprKan(dwCornerIdx + 1, dwSequense, _
                                                PATH_WORK, _
                                                szFileName)
    Else
        'έR[iΘηΞέR[ipΜΦπΔΡo·
        lResult = dllCreateShimekiriFileJpr(dwCornerIdx + 1, dwSequense, _
                                            PATH_WORK, _
                                            szFileName)
    End If
    'EG20 V30.1.0.1 ADD END
    If lResult = False Then
        sOutPutOfflineData = False
        Exit Function
    End If

    sOutPutOfflineData = True
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : JprEdit_KadoVersion
'//  @\ΌΜ  : ??o[WW[iC[Wt@Cμ¬
'//  @\Tv  : ??o[WW[iC[Wt@Cπμ¬·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Θ΅
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-05-07   CODED   BY [TCC] T.Nakajima
'//             k€V²όJΖΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_KadoVersion() As Boolean

    Dim strOutputFile As String         'oΝt@C
    Dim lngRet As Long                  'ΦΤθl
    Dim lngErrCode As Long              'G[R[h
    Dim iOutFile    As Integer          't@CΤ
    Dim ReadFileKado()    As KADO_VER_DISP_IMAGE_FILE '?­o[Wκ³f[^
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim l           As Integer
    Dim strDispImageFileName As String
    Dim objFs       As New FileSystemObject
    Dim intFileNo   As Integer
    Dim iHeadFlg    As Integer
    
    
    On Error GoTo Err_handler
    
    'ζΚΕwθ³κ½R[iΝέu³κΔ’ι©H
    If pfSettingCheck(True) = False Then
        '·ΧΔ’έuΜR[iA@ΘΜΕG[Ζ·ι
        GoTo Err_handler
    End If
    
    'oΝt@CΌ?W
    strOutputFile = KADOVER_TXTFILE
    
    '// R[iΌπκΚθζΎ ζΎΚΝgstrCornerName(0 to 5)ΙόΑΔ’ι
    gsGetCornerName
    'EG20 V30.0.1.1 ADD START
    ' R[i^CvζΎ
    gsGetCornerType
    'EG20 V30.0.1.1 ADD END

    
    'wΌπζΎ   ζΎΚΝgstrStationName(0 to 5)ΙόΑΔ’ι
    gsGetStationName
    
    iHeadFlg = 0
    
    't@CoΝΦπCall
    '`FbN³κΔ’ιR[iA@ͺΜo[Wt@Cπ’Α½ρoΝ
    For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
            '»ΜR[iA@Νέu³κΔ’ι©H
            If pfCornerGokiCheck(udtJprPrintSetteingInfo.iCorner(i), udtJprPrintSetteingInfo.iGouki(j)) = True Then
        
                strDispImageFileName = Replace(EDIT_DATA_KADOVERSION, "####", _
                    Format(udtJprPrintSetteingInfo.iCorner(i), "0#") & Format(udtJprPrintSetteingInfo.iGouki(j), "0#")) & ".csv"
                
                'EG20 V30.1.0.1 DEL START
'                lngRet = dllCreateKadoVersionFile(mintDispDiv.KADOVER_FILE_DISP, udtJprPrintSetteingInfo.iCorner(i), _
'                                                  udtJprPrintSetteingInfo.iGouki(j), strDispImageFileName, PATH_IDU_APP, PATH_LDU_APP)
                'EG20 V30.1.0.1 DEL END
                'EG20 V30.1.0.1 ADD START
                If gintCornerType(udtJprPrintSetteingInfo.iCorner(i) - 1) = CORNER_TYPE_KANSEN Then
                
                    lngRet = dllCreateKadoVersionFileKan(mintDispDiv.KADOVER_FILE_DISP, udtJprPrintSetteingInfo.iCorner(i), _
                                                      udtJprPrintSetteingInfo.iGouki(j), strDispImageFileName, PATH_IDU_APP, PATH_LDU_APP)
                
                Else
                    lngRet = dllCreateKadoVersionFile(mintDispDiv.KADOVER_FILE_DISP, udtJprPrintSetteingInfo.iCorner(i), _
                                                      udtJprPrintSetteingInfo.iGouki(j), strDispImageFileName, PATH_IDU_APP, PATH_LDU_APP)
                End If
                'V30.1.0.1 ADD END
                
                'ΩνIΉΝG[Φ
                If lngRet = 0 Then
                    GoTo Err_handler
                    Exit Function
                End If
                
                't@CͺΆέ΅Θ’κΝG[Φ
                If objFs.FileExists(strDispImageFileName) = False Then
                    GoTo Err_handler
                    Exit Function
                End If
            End If
        Next j
    Next i
    
    '?­o[Wκ W[iC[Wt@Cπμ¬
    iOutFile = FreeFile
    Open strOutputFile For Output As #iOutFile
    
    'wb_[
    PrintHeader iOutFile, "?­o[Wκ"
    
    'έuw
    Print #iOutFile, "έuwF" & gstrStationName(0)
    Print #iOutFile, ""
    
    'ζΚ\¦pt@CπI[v
    For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        Erase ReadFileKado
        If i > 0 Then
            '1R[iΪΝSΜo[Wπ\¦΅Δ©ηR[iΌπoΝ
            '»ΜR[iΝέu³κΔ’ι©H
            'If pfCornerGokiCheck(udtJprPrintSetteingInfo.iCorner(i)) = True Then
            If IsTaisyoCorner(udtJprPrintSetteingInfo.iCorner(i)) = True Then
                'ΞΫR[iΕ ΑΔΰΞΫ@ͺΘ’©ΰ΅κΘ’
                For j = 0 To 15
                    If IsTaisyoGoki(udtJprPrintSetteingInfo.iCorner(i), j + 1) = True Then
                        Print #iOutFile, "R[iΌF" & gstrCornerName(udtJprPrintSetteingInfo.iCorner(i) - 1)
                        Exit For
                    End If
                Next j
                        
            End If
        End If
    
        For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
            '»Μ@ͺέu³κΔ’ι©H
            If pfCornerGokiCheck(udtJprPrintSetteingInfo.iCorner(i), udtJprPrintSetteingInfo.iGouki(j)) = True Then
    
                intFileNo = FreeFile
                strDispImageFileName = Replace(EDIT_DATA_KADOVERSION, "####", _
                    Format(udtJprPrintSetteingInfo.iCorner(i), "0#") & Format(udtJprPrintSetteingInfo.iGouki(j), "0#")) & ".csv"
                
                Open strDispImageFileName For Input As #intFileNo
        
                'ζΚ\¦pf[^(csv)πGAΙΗέή
                k = 0
                Do While Not EOF(intFileNo)
                    ReDim Preserve ReadFileKado(k)
                    'intKishu, intCorner, intGokiDiv, strName, strMaker, strVer, strDate
                    Input #intFileNo, _
                            ReadFileKado(k).strKishu, ReadFileKado(k).strCorner, ReadFileKado(k).strGokiDiv, _
                            ReadFileKado(k).strName, ReadFileKado(k).strMaker, ReadFileKado(k).strVer, ReadFileKado(k).strDate
                    k = k + 1
                Loop
                't@CN[Y
                Close #intFileNo
                
                'ΕΜ[vΎ―SΜξρπ\¦
                'If i = 0 And j = 0 Then
                If iHeadFlg = 0 Then
                    
                    'ΔΥSΜo[W
                    Print #iOutFile, "ΔΥSΜo[W"
                    Print #iOutFile, ReadFileKado(0).strVer
                    
                    'ΔΥ
                    Print #iOutFile, "ΔΥ"
                    Print #iOutFile, ReadFileKado(1).strVer
                    
                    'hcto[W
                    Print #iOutFile, "hct"
                    Print #iOutFile, ReadFileKado(2).strVer
                    
                    'kcto[W
                    Print #iOutFile, "kct"
                    Print #iOutFile, ReadFileKado(3).strVer
                    Print #iOutFile, ""
                    
                    'μμ
                    Print #iOutFile, "μμ"
                    Print #iOutFile, ReadFileKado(4).strVer
                    Print #iOutFile, ""
                    
                    'R[iΌ
                    Print #iOutFile, "R[iΌF" & gstrCornerName(udtJprPrintSetteingInfo.iCorner(i) - 1)
                    
                    iHeadFlg = 1
                End If
    
                '@Τ
                Print #iOutFile, "@ΤF" & Format(udtJprPrintSetteingInfo.iGouki(j), "00") & "@"
                'evOo[W(6sΪ©ηevOo[W)
                For l = 0 To k - 1
                    If ReadFileKado(l).strKishu = "06" Then
                        '\υΜκΝo[Wπo³Θ’
                        If ReadFileKado(l).strName = "\υP" Or ReadFileKado(l).strName = "\υQ" Then
                            Print #iOutFile, ReadFileKado(l).strName
                        Else
                            Print #iOutFile, ReadFileKado(l).strName & Space(11 - LenB(StrConv(ReadFileKado(l).strName, vbFromUnicode))) & ReadFileKado(l).strVer
                        End If
                    End If
                Next l
                Print #iOutFile, ""
            End If
        Next j
    Next i
    
    Print #iOutFile, FOOTER_STRING
    Close #iOutFile
    
  
    JprEdit_KadoVersion = True
    Exit Function

Err_handler:

    If intFileNo > 0 Then
        Close #intFileNo
    End If
    
    If iOutFile > 0 Then
        Close #iOutFile
    End If
    
    Set objFs = Nothing

    'MsgBox "ΩνIΉ΅ά΅½B", vbCritical, "oΝΚ"
    'u?­o[WΗζΚF?­o[Wξρ}ΜoΝΩνvOoΝ
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, lngErrCode)
    JprEdit_KadoVersion = False

End Function


'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2013 All Right Reserved
'//
'//  ΦΌΜ : JprEdit_EkimuId
'//  Tv     : w±@νIDW[iC[Wt@Cμ¬
'//  ΰΎ     : w±@νIDW[iC[Wt@Cπμ¬·ι
'//  ΚίΧ?°ΐ   :
'//           :
'//
'//  ORIGINAL  F(EG20 V7.2.0.1) 2013-06-26  CODED BY  [TCC] T.Nakajima
'//  REVISIONS F(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_EkimuId() As Boolean
    
    Dim sEkimuIDFile    As String   'w±@νIDt@CpX
    Dim iRet            As Integer  'INIζΎίθl
    Dim sFolder         As String * MAX_PATH_SIZE  'tH_Ό
    Dim sFile           As String   't@CΌ
    Dim MyName          As String   't@CυΚ
    Dim bRet            As Boolean  'ίθl
    Dim lngErrCode      As Long     'G[R[h
    Dim intFileNo       As Integer  't@CΤ
    Dim strWork         As String   'μΖGA
    Dim dwErrsts        As Long
    Dim sFolderName     As String
    Dim objFso          As New FileSystemObject
    Dim objTs           As TextStream
    
        
    On Error GoTo Err_handler
    sFolder = ""
    
    'ΚF³νΝζΚ\¦
    iRet = GetPrivateProfileString(IDU_SECTION_NAME, _
                                   IDU_EKIMUID_KEY, _
                                   EKIMU_DEFU, sFolder, Len(sFolder), _
                                   PATH_IDU_INI_FILE)
    If iRet = 0 Then
      sFolder = EKIMU_DEFU
    End If
    sEkimuIDFile = ""
    'vνΚlζθt@CΌμ¬
    sFile = Replace(EKIMU_ID_FILE, "##", Format(iSendType, "0#"))
    If iRet = 0 Then
       sFolderName = RTrim(sFolder)
    Else
       sFolderName = Mid(sFolder, 1, iRet)
    End If
    'pXΟ·
    sFolderName = pfChangeFolderName(sFolderName)
    'w±@νIDt@CpXμ¬
    sEkimuIDFile = sFolderName & "\" & sFile
    't@CL³`FbN
    If Dir(sEkimuIDFile, vbNormal) = "" Then
       Exit Function
    End If
    
    '/////////////////////////////////////////////////////////////////////
    '//ΫηκpΦFw±@νIDζΚ\¦pt@Cμ¬
    '////////////////////////////////////////////////////////////////////
    bRet = dllEKIMUKIKI(sEkimuIDFile, dwErrsts, MN_VERSI_FILE, PATH_IDU_APP, 1) 'V1.8.0.1 ADD
    
    If bRet = False Then
        GoTo Err_handler
        Exit Function
    End If
    
    
    '/////////////////////////////////////////////////////////////////////
    '//W[iC[Wt@Cπμ¬
    '////////////////////////////////////////////////////////////////////
    intFileNo = FreeFile
    Open EKIMUKIKI_ID_TXTFILE For Output As #intFileNo
    
    'wb_oΝ
    PrintHeader intFileNo, "w±@νhcoΝ"
    
    'έuwΌ
    gsGetStationName
    Print #intFileNo, "έuwF" & gstrStationName(0)
    Print #intFileNo, ""
    
    'f[^πΒΘ°ι
    Set objTs = objFso.OpenTextFile(MN_VERSI_FILE, ForReading)
    Print #intFileNo, objTs.ReadAll
    objTs.Close
    Set objFso = Nothing
    
    'tb^μ¬
    Print #intFileNo, FOOTER_STRING
    
    Close #intFileNo
    
    JprEdit_EkimuId = True
    
    Exit Function

Err_handler:

    If intFileNo > 0 Then
        Close #intFileNo
    End If
    
    
    Set objFso = Nothing

    'MsgBox "ΩνIΉ΅ά΅½B", vbCritical, "oΝΚ"
    'u?­o[WΗζΚF?­o[Wξρ}ΜoΝΩνvOoΝ
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, lngErrCode)
    JprEdit_EkimuId = False
    
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  ΦΌΜ  : pfChangeFolderName
'//  @\ΌΜ  : tH_pXΟ·
'//  @\Tv  : INIt@CζθζΎ΅½tH_θ`ΜΟ·πs€B
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : String sFolderName    [IN]INIθ`
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υl F
'///////////////////////////////////////////////////////////////////
Private Function pfChangeFolderName(sFolderName As String) As String
   Dim iPath As Integer
   Dim sRootPath As String
   Dim sFolder As String
      
   'uvΚuπζΎ
   iPath = InStr(sFolderName, "\")
   If iPath = 0 Then
     sRootPath = Mid(sFolderName, 1)
   Else
     'uvOΆρπζΎ
     sRootPath = Mid(sFolderName, 1, iPath - 1)
     'uvγΆρπζΎ
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


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : JprEdit_TukaData
'//  @\ΌΜ  : Κίf[^/pΰzW[iC[Wt@Cμ¬
'//  @\Tv  : Κίf[^/pΰzW[iC[Wt@Cπμ¬·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : long      dwDataKind f[^νΚ    Κί}ΜF306010
'//                                                 p}ΜF306020
'//
'//              ^        l        Σ‘
'//  ίθl    : Boolean@@@@@@[OUT]ίθl
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-01   CODED   BY [TCC] T.Nakajima
'//                 k€V²όJΖΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_TukaData(dwDataKind As Long) As Boolean
    
    Dim strFilePath As String           'oΝt@CpX
    Dim intCount As Integer             'JE^
    'EG20 V30.1.0.1 DEL START
'    Dim intOutFile As Integer           'oΝt@CΤ
'    Dim strBaitaiFileName As String     ' }ΜoΝt@C TUKAR[iΌYYYYMMDDhhmmss.csv ICRIYOR[iΌYYYYMMDDhhmmss.csv
'    Dim ReadFileBaitai()  As BAITAI_OUTPUT_IMAGE_FILE '}ΜoΝt@C
'    Dim strLineCount()  As String
'    Dim i As Integer
'    Dim j As Integer
'    Dim k As Integer
'    Dim l As Integer
'    Dim strCammaArray() As String   'J}ζΨθΕ1ΪΈΒζθo΅½f[^

'    Dim fso As New FileSystemObject
'    Dim FsoTS As TextStream
    
'    Dim iKomokuMaxCnt       As Integer      ' Wvf[^ΪΜΕε
'    Dim iStartLineKaisatu   As Integer      ' όD€f[^ΜJnsibrut@CΜj
'    Dim iStartLineShusatu   As Integer      ' WD€f[^ΜJnsibrut@CΜj
    'EG20 V30.1.0.1 DEL END
    
    'Dim intJprFile        As Integer
    
    On Error GoTo Err_handler
    'ζΚΕwθ³κ½R[iΝέu³κΔ’ι©H
    If pfSettingCheck(False) = False Then
        '·ΧΔ’έuΘΜΕG[
        GoTo Err_handler
    End If
  
    '////////////////////////////////////////////////
    '// έuwER[iΌπκΚθζΎ
    gsGetStationName
    gsGetCornerName
    gsGetCornerType
    gsGetShukeiKoumoku     'WvΪΜoΝL³πζΎ    EG20 V30.1.0.1 ADD

   
    'R[iPΚΕ
    
    '/////////////////////////////////////////////
    '// W[iC[Wt@Cμ¬
    
    'oΝt@CπI[v·ιB
    intJprFile = FreeFile
    If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
        Open TUKA_TXTFILE For Output As #intJprFile
    ElseIf dwDataKind = Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI Then
        Open ICRIYO_TXTFILE For Output As #intJprFile
    Else
        JprEdit_TukaData = False
        Exit Function
    End If

   '^Cg\¦
   If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
        'EG20 V30.1.0.1 DEL START iέΖ²όΙζΑΔέθlͺΩΘιΜΕC[Wt@CΦΪ?j
'        iKomokuMaxCnt = MAX_KOMOKU_NUM_TUKA
'        iStartLineKaisatu = 6   'όD€ΜΎΧΝ³t@C(CSV)zρΜ(6)©η
'        iStartLineShusatu = 60  'WD€ΜΎΧΝ³t@C(CSV)zρΜ(60)©η
        'EG20 V30.1.0.1 DEL END
        PrintHeader intJprFile, "Κίf[^oΝ"
    Else
        'EG20 V30.1.0.1 DEL START iέΖ²όΙζΑΔέθlͺΩΘιΜΕC[Wt@CΦΪ?j
'        iKomokuMaxCnt = MAX_KOMOKU_NUM_KINGAKU
'        iStartLineKaisatu = 6   'όD€ΜΎΧΝ³t@C(CSV)zρΜ(6)©η
'        iStartLineShusatu = 25  'WD€ΜΎΧΝ³t@C(CSV)zρΜ(60)©η
        'EG20 V30.1.0.1 DEL END
        PrintHeader intJprFile, "pΰzf[^oΝ"
    End If

    'έuwER[iΌoΝ
    Print #intJprFile, "έuwF" & gstrStationName(0)
    
    For intCount = 0 To UBound(glngTergetCorner)
    
        If glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON Then
            'EG20 V30.1.0.1 ADD START
            If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
                    psMakeTukaImageFileKan intCount
                Else
                    psMakeRiyoImageFileKan intCount
                End If
            Else
                psMakeTukaRiyoImageFile intCount, dwDataKind
            End If
            
        
            'EG20 V30.1.0.1 DEL START
'            If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then       'Κίf[^
'                strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "TUKA" & gstrStationName(intCount) & gstrCornerName(intCount) & "*.csv")
'            ElseIf dwDataKind = Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI Then    'pΰzf[^
'                strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "ICRIYO" & gstrStationName(intCount) & gstrCornerName(intCount) & "*.csv")
'            Else
'                JprEdit_TukaData = False
'                Exit Function
'            End If
'
'            '////////////////////////////////////////////////
'            '// Κίf[^/pΰzΜ}ΜoΝt@CπζΎ
'            't@CΤζΎ
'            'wΌΜ{R[iΌΜyyyymmddhhmmss.csv
'            Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForAppending)
'            j = FsoTS.Line
'            FsoTS.Close
'
'            ReDim strLineCount(j) As String         'CSVt@Cπ1sΈΒόκΔ¨­
'
'            i = 0
'            Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForReading)
'            Do Until FsoTS.AtEndOfStream = True
'                strLineCount(i) = FsoTS.ReadLine
'                i = i + 1
'            Loop
'            FsoTS.Close
'            Set fso = Nothing
'
'            '}ΜoΝt@CC[W\’ΜΙZbg·ι
'            ReDim ReadFileBaitai(j) As BAITAI_OUTPUT_IMAGE_FILE         't@CΗpGA
'            l = 0
'            If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
'
'                For i = 0 To j - 1
'                    Select Case i
'                        Case 0, 1, 2, 3, 4, 57, 58    'TUKAxxxx.csvΜ1`4sΪάΕΝ^CgΘΜΕAΪΌΙZbg
'                            ReadFileBaitai(i).strKomokuName = strLineCount(i)
'                        Case Else
'                            'J}ζΨθπ1ΪΈΒζθo·B
'                            strCammaArray = Split(strLineCount(i), ",")
'                            For k = 0 To UBound(strCammaArray())
'                                If k = 0 Then
'                                    ReadFileBaitai(i).strKomokuName = strCammaArray(k)
'                                ElseIf k = 1 Then
'                                    ReadFileBaitai(i).strGoukei = strCammaArray(k)
'                                Else
'                                    ReadFileBaitai(i).srtGoukiValue(l) = strCammaArray(k)
'                                    l = l + 1
'                                End If
'                            Next k
'                    End Select
'                    l = 0
'                Next i
'            Else
'                For i = 0 To j - 1
'                    Select Case i
'                        Case 0, 1, 2, 3, 4, 22, 23    'ICRIYOxxxx.csvΜ1`4sΪάΕΝ^CgΘΜΕAΪΌΙZbg
'                            ReadFileBaitai(i).strKomokuName = strLineCount(i)
'                        Case Else
'                            'J}ζΨθπ1ΪΈΒζθo·B
'                            strCammaArray = Split(strLineCount(i), ",")
'                            For k = 0 To UBound(strCammaArray())
'                                If k = 0 Then
'                                    ReadFileBaitai(i).strKomokuName = strCammaArray(k)
'                                ElseIf k = 1 Then
'                                    ReadFileBaitai(i).strGoukei = strCammaArray(k)
'                                Else
'                                    ReadFileBaitai(i).srtGoukiValue(l) = strCammaArray(k)
'                                    l = l + 1
'                                End If
'                            Next k
'                    End Select
'                    l = 0
'                Next i
'            End If
'
'            Print #intJprFile, "έuR[iF" & gstrCornerName(intCount)
'            Print #intJprFile, ""
'
'            If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
'                Print #intJprFile, "yΚίf[^z"
'            Else
'                Print #intJprFile, "yhbJ[hpΰzf[^z"
'            End If
'            '/////////////////////
'            'όD€f[^ΜoΝ
'            Print #intJprFile, "όD€Κίv"
'
'            For i = 0 To MAX_KOMOKU_NUM_TUKA - 1
'                If RTrim(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName) = "0" Then
'                    'ΪΌΙ0ͺZbg³κΔ’½ηΘ~ΝoΝ΅Θ’
'                    Exit For
'                Else
'                    'χΨItCW[iΖ νΉι½ίΜαO
'                    If RTrim(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName) = "»ΜΌIC (¬)" Then
'                        ReadFileBaitai(i + iStartLineKaisatu).strKomokuName = "»ΜΌIC(¬)" & Space(38)   'Xy[Xπ­
'                    End If
'
'                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName, vbFromUnicode), 20), vbUnicode) _
'                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLineKaisatu).strGoukei), "#,0"), 10)
'                End If
'            Next i
'            Print #intJprFile, ""
'
'            '/////////////////////
'            'WD€f[^ΜoΝ
'            Print #intJprFile, "WD€Κίv"
'
'            For i = 0 To MAX_KOMOKU_NUM_TUKA - 1
'                If RTrim(ReadFileBaitai(i + iStartLineShusatu).strKomokuName) = "0" Then
'                    'ΪΌΙ0ͺZbg³κΔ’½ηΘ~ΝoΝ΅Θ’
'                    Exit For
'                Else
'                    'χΨItCW[iΖ νΉι½ίΜαO
'                    If RTrim(ReadFileBaitai(i + iStartLineShusatu).strKomokuName) = "»ΜΌIC (¬)" Then
'                        ReadFileBaitai(i + iStartLineShusatu).strKomokuName = "»ΜΌIC(¬)" & Space(38)    'Xy[Xπ­
'                    End If
'
'                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLineShusatu).strKomokuName, vbFromUnicode), 20), vbUnicode) _
'                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLineShusatu).strGoukei), "#,0"), 10)
'                End If
'            Next i
'            Print #intJprFile, ""
        'EG20 V30.1.0.1 DEL END
            
        End If
    Next intCount
    
    Print #intJprFile, FOOTER_STRING
    Close #intJprFile
    
    JprEdit_TukaData = True
    Exit Function
    
'G[
Err_handler:

    'EG20 V30.1.0.1 DEL START
'    If intOutFile > 0 Then
'        Close #intOutFile
'    End If
    'EG20 V30.1.0.1 DEL END
    
    If intJprFile > 0 Then
        Close #intJprFile
    End If

'    Set fso = Nothing      'EG20 V30.1.0.1 DEL
    'G[OΜoΝ
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
    JprEdit_TukaData = False
                                      
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : PadLeft
'//  @\ΌΜ  : EρΉ
'//  @\Tv  : wθΜΆΙΘιάΕζͺπΆΕίιB
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : string    strTarget    ΞΫΆρ
'//              Integer   iLength      ΆΜ·³
'//              string    chOne        ίιΆ(ΘͺΝΌpXy[X)
'//
'//              ^        l        Σ‘
'//  ίθl    : string    EρΉ³κ½Άρ
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function PadLeft(ByVal strTarget As String, ByVal iLength As Integer, Optional ByVal chOne As String = " ") As String
    
    Do While (Len(strTarget) < iLength)
        strTarget = chOne & strTarget
    Loop

    PadLeft = Right$(strTarget, iLength)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : PadRight
'//  @\ΌΜ  : ΆρΉiφπXy[XΕίι)
'//  @\Tv  : wθΜΆΙΘιάΕζͺπΆΕίιB
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : string    strTarget    ΞΫΆρ
'//              Integer   iLength      ΆΜ·³
'//              string    chOne        ίιΆ(ΘͺΝΌpXy[X)
'//
'//              ^        l        Σ‘
'//  ίθl    : string    ΆρΉ³κ½Άρ
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Public Function PadRight(ByVal strTarget As String, ByVal iLength As Integer, Optional ByVal chOne As String = " ") As String
    Do While (Len(strTarget) < iLength)
        strTarget = strTarget & chOne
    Loop

    PadRight = Left$(strTarget, iLength)
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : PrintHeader
'//  @\ΌΜ  : wb_μ¬
'//  @\Tv  : wb_πμ¬·ιBiW[iΜP`SsΪ)
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Integer   iFileNum     t@CΤ
'//              string    strTitle     W[i^Cg
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG30 V32.1.0.1 2016-06-14   REVISED BY [TCC] T.Nakajima
'//                 2016Nx{τΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub PrintHeader(iFileNum As Integer, strTitle As String)
    Dim lpSystemTime            As SYSTEMTIME               '[JπζΎ
    'EG30 V32.1.0.1 DEL START
    'Print #iFileNum, "*************EG20*************"
    'EG30 V32.1.0.1 DEL END
    Print #iFileNum, strTitle
    
    '[JπζΎ
    Call GetLocalTime(lpSystemTime)
    Print #iFileNum, "σϊF" & lpSystemTime.wYear & "N" & Format(lpSystemTime.wMonth, "00") & "" & Format(lpSystemTime.wDay, "00") & "ϊ" _
                            & Format(lpSystemTime.wHour, "00") & ":" & Format(lpSystemTime.wMinute, "00")
    Print #iFileNum, ""
End Sub
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  ΦΌΜ  : PrintHeader2
'//  @\ΌΜ  : wb_μ¬
'//  @\Tv  : wb_πμ¬·ιBiW[iΜP`SsΪ)
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Integer   iFileNum     t@CΤ
'//              string    strTitle     W[i^Cg
'//              string    strTitle2    W[i^CgQsΌ
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(EG20 V30.3.0.1) 2014-10-01   CODED   BY [TCC] T.Nakajima
'//                 k€V²όtF[YQΞ yHKRK_Kansi07_003_01z
'//     REVISIONS :(EG30 V32.1.0.1 2016-06-14   REVISED BY [TCC] T.Nakajima
'//                 2016Nx{τΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub PrintHeader2(iFileNum As Integer, strTitle As String, strTitle2 As String)
    Dim lpSystemTime            As SYSTEMTIME               '[JπζΎ
    
    'EG30 V32.1.0.1 DEL START
    'Print #iFileNum, "*************EG20*************"
    'EG30 V32.1.0.1 DEL END
    Print #iFileNum, strTitle
    Print #iFileNum, strTitle2
    
    '[JπζΎ
    Call GetLocalTime(lpSystemTime)
    Print #iFileNum, "σϊF" & lpSystemTime.wYear & "N" & Format(lpSystemTime.wMonth, "00") & "" & Format(lpSystemTime.wDay, "00") & "ϊ" _
                            & Format(lpSystemTime.wHour, "00") & ":" & Format(lpSystemTime.wMinute, "00")
    Print #iFileNum, "ΫΆϊF" & pfGetSaveDate(0)    'R[i0ΜΫΆϊ  'EG30 V32.1.0.1 ADD
    Print #iFileNum, ""
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  ΦΌΜ  : PrintHeader3
'//  @\ΌΜ  : wb_μ¬
'//  @\Tv  : wb_πμ¬·ιB
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Integer   iFileNum     t@CΤ
'//              string    strTitle     W[i^Cg
'//              string    strSaveDate  ΫΆϊ
'//
'//              ^        l        Σ‘
'//  ίθl    : Θ΅
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-22   CODED   BY [TCC] T.Nakajima
'//                 2016Nx{τΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub PrintHeader3(iFileNum As Integer, strTitle As String, strSaveDate As String)
    Dim lpSystemTime            As SYSTEMTIME               '[JπζΎ
    
    Print #iFileNum, strTitle
    '[JπζΎ
    Call GetLocalTime(lpSystemTime)
    Print #iFileNum, "σϊF" & lpSystemTime.wYear & "N" & Format(lpSystemTime.wMonth, "00") & "" & Format(lpSystemTime.wDay, "00") & "ϊ" _
                            & Format(lpSystemTime.wHour, "00") & ":" & Format(lpSystemTime.wMinute, "00")
    Print #iFileNum, "ΫΆϊF" & strSaveDate
    Print #iFileNum, ""
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : pfCornerGokiCheck
'//  @\ΌΜ  : R[i@`FbN
'//  @\Tv  : ζΚΕ`FbN³κ½R[i@ͺΆέ·ι©mF·ι
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Integer   iCorner      R[i(1`6)
'//              Integer   iGoki        @iΘͺΒ\FΘͺΝ@Ν`FbN΅Θ’) 1`16
'//
'//              ^        l           Σ‘
'//  ίθl    : Boolean   true/false   true:έu³κΔ’ι   false:έu³κΔ’Θ’
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function pfCornerGokiCheck(iCornerNo As Integer, Optional iGoki As Integer = 0) As Boolean
    Dim i       As Integer
    Dim bRet    As Boolean
    bRet = False
    
    If gudtSettiCorner(iCornerNo - 1).intGokiNum > 0 Then
        ' wθ΅½R[iΝέu³κΔ’ι
        ' p[^Εwθ³κ½@Νέu³κΔ’ι©H
        If iGoki <> 0 Then
            For i = 0 To 15
                If iGoki = gudtSettiCorner(iCornerNo - 1).intGokiNo(i) Then
                    bRet = True
                    Exit For
                End If
            Next i
        Else
            bRet = True
        End If
    Else
        'wθ³κ½R[iΝέu³κΔ’Θ’
    End If
    
    pfCornerGokiCheck = bRet
    Exit Function
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  ΦΌΜ  : pfCornerGokiToGateNo
'//  @\ΌΜ  : R[i@¨_@ΤΙΟ·
'//  @\Tv  : ζΚΕ`FbN³κ½R[i@ͺΆέ·ι©mF΅A_@ΤπΤ·B
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Integer   iCorner      R[i(1`6)
'//              Integer   iGoki        @iΘͺΒ\FΘͺΝ@Ν`FbN΅Θ’) 1`16
'//              Integer   iGateNo      _@(1`32)
'//
'//              ^        l           Σ‘
'//  ίθl    : Boolean   true/false   true:έu³κΔ’ι   false:έu³κΔ’Θ’
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-28   CODED   BY [TCC] T.Nakajima
'//                 2016Nx{τΞ
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function pfCornerGokiToGateNo(iCornerNo As Integer, iGoki As Integer, ByRef iGateNo As Integer) As Boolean
    Dim i       As Integer
    Dim bRet    As Boolean
    bRet = False
    iGateNo = 0
    If gudtSettiCorner(iCornerNo - 1).intGokiNum > 0 Then
        ' wθ΅½R[iΝέu³κΔ’ι
        ' p[^Εwθ³κ½@Νέu³κΔ’ι©H
        If iGoki <> 0 Then
            For i = 0 To 15
                If iGoki = gudtSettiCorner(iCornerNo - 1).intGokiNo(i) Then
                    iGateNo = gudtSettiCorner(iCornerNo - 1).intGateNo(i)
                    bRet = True
                    Exit For
                End If
            Next i
        Else
            bRet = True
        End If
    Else
        'wθ³κ½R[iΝέu³κΔ’Θ’
    End If
    
    pfCornerGokiToGateNo = bRet
    Exit Function
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : pfSettingaCheck
'//  @\ΌΜ  : R[i@ΜέumF
'//  @\Tv  : W[iΙoΝ·ιR[i@ͺέu³κΔ’ι©mF·ιB
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Boolean   bGokiCheck   @`FbNL³
'//
'//
'//              ^        l           Σ‘
'//  ίθl    : Boolean   true/false   true:έu³κΔ’ι   false:έu³κΔ’Θ’
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function pfSettingCheck(Optional bGokiCheck As Boolean = True) As Boolean
    Dim i   As Integer
    Dim j   As Integer
    Dim k   As Integer
    
    'ζΚΕέθ³κ½€ΏΗκ©ΠΖΒΕΰέu³κΔ’ιR[i@ͺ κΞOKΖ·ι
    For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        If gudtSettiCorner(udtJprPrintSetteingInfo.iCorner(i) - 1).intGokiNum > 0 Then
            '»ΜR[iΝέu³κΔ’ι
            ' `FbN³κ½@Ν»ΜR[iΙΆέ΅Δ’ι©H(@`FbN θΜκ)
            If bGokiCheck = True Then
                For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
                    For k = 0 To 15
                        If udtJprPrintSetteingInfo.iGouki(j) = gudtSettiCorner(udtJprPrintSetteingInfo.iCorner(i) - 1).intGokiNo(k) Then
                            pfSettingCheck = True
                            Exit Function
                        End If
                    Next k
                Next j
            Else
                pfSettingCheck = True
                Exit Function
            End If
        End If
    Next i
                
    pfSettingCheck = False
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  ΦΌΜ  : MidByte
'//  @\ΌΜ  : R[i@ΜέumF
'//  @\Tv  : W[iΙoΝ·ιR[i@ͺέu³κΔ’ι©mF·ιB
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : String    strTarget     ΞΫΆρ
'//              long      iStart       JnΚu(1oCg`)
'//              Variant   ibyteCount   ·³
'//
'//
'//              ^        l           Σ‘
'//  ίθl    :String                  o³κ½Άρ
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function MidByte(ByVal strTarget As String, ByVal iStart As Long, Optional ByVal iByteCount As Variant) As String
    If IsMissing(iByteCount) = False Then
        MidByte = StrConv(MidB$(StrConv(strTarget, vbFromUnicode), iStart, iByteCount), vbUnicode)
    Else
        MidByte = StrConv(MidB$(StrConv(strTarget, vbFromUnicode), iStart), vbUnicode)
    End If
End Function


'EG20 V30.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  ΦΌΜ  : psMakeTukaRiyoImageFile
'//  @\ΌΜ  : Κίf[^/pΰzf[^W[iΜC[Wt@Cμ¬iέpj
'//  @\Tv  : Κίf[^¨ζΡpΰzf[^W[iΜC[Wt@Cπμ¬·ιB
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Integer   iCornerIdx   R[iCfbNX
'//              Long      dwDataKind   f[^νΚiΚίf[^Apΰzf[^j
'//
'//
'//              ^        l           Σ‘
'//  ίθl    : ³΅
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF      ²όR[ipΜC[Wt@Cμ¬ͺΚrKvΖΘΑ½½ίA
'//              JprEdit_TukaData()©ηTu[`»
'///////////////////////////////////////////////////////////////////
Private Sub psMakeTukaRiyoImageFile(iCornerIdx As Integer, dwDataKind As Long)
    
    Dim strBaitaiFileName   As String                       '}ΜoΝt@C TUKAR[iΌYYYYMMDDhhmmss.csv ICRIYOR[iΌYYYYMMDDhhmmss.csv
    Dim ReadFileBaitai()    As BAITAI_OUTPUT_IMAGE_FILE     '}ΜoΝt@C
    Dim intOutFile          As Integer                      'oΝt@CΤ
    Dim strLineCount()      As String
    Dim fso                 As New FileSystemObject
    Dim FsoTS               As TextStream
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    Dim l                   As Integer
    Dim strCammaArray()     As String                       'J}ζΨθΕ1ΪΈΒζθo΅½f[^
    Dim iKomokuMaxCnt       As Integer                      ' Wvf[^ΪΜΕε
    Dim iStartLineKaisatu   As Integer                      ' όD€f[^ΜJnsibrut@CΜj
    Dim iStartLineShusatu   As Integer                      ' WD€f[^ΜJnsibrut@CΜj
            
    On Error GoTo Err_handler
    
    If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then       'Κίf[^
        
        iKomokuMaxCnt = MAX_KOMOKU_NUM_TUKA
        iStartLineKaisatu = 6   'όD€ΜΎΧΝ³t@C(CSV)zρΜ(6)©η
        iStartLineShusatu = 60  'WD€ΜΎΧΝ³t@C(CSV)zρΜ(60)©η
        
        strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "TUKA" & gstrStationName(iCornerIdx) & gstrCornerName(iCornerIdx) & "*.csv")
    ElseIf dwDataKind = Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI Then    'pΰzf[^
        
        iKomokuMaxCnt = MAX_KOMOKU_NUM_KINGAKU
        iStartLineKaisatu = 6   'όD€ΜΎΧΝ³t@C(CSV)zρΜ(6)©η
        iStartLineShusatu = 25  'WD€ΜΎΧΝ³t@C(CSV)zρΜ(60)©η
        
        strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "ICRIYO" & gstrStationName(iCornerIdx) & gstrCornerName(iCornerIdx) & "*.csv")
    End If
           
    '////////////////////////////////////////////////
    '// Κίf[^/pΰzΜ}ΜoΝt@CπζΎ
    't@CΤζΎ
    'wΌΜ{R[iΌΜyyyymmddhhmmss.csv
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForAppending)
    j = FsoTS.Line
    FsoTS.Close
           
    ReDim strLineCount(j) As String         'CSVt@Cπ1sΈΒόκΔ¨­
           
    i = 0
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForReading)
    Do Until FsoTS.AtEndOfStream = True
        strLineCount(i) = FsoTS.ReadLine
        i = i + 1
    Loop
    FsoTS.Close
    Set fso = Nothing
    
    '}ΜoΝt@CC[W\’ΜΙZbg·ι
    ReDim ReadFileBaitai(j) As BAITAI_OUTPUT_IMAGE_FILE         't@CΗpGA
    l = 0
    If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
    
        For i = 0 To j - 1
            Select Case i
                Case 0, 1, 2, 3, 4, 57, 58    'TUKAxxxx.csvΜ1`4sΪάΕΝ^CgΘΜΕAΪΌΙZbg
                    ReadFileBaitai(i).strKomokuName = strLineCount(i)
                Case Else
                    'J}ζΨθπ1ΪΈΒζθo·B
                    strCammaArray = Split(strLineCount(i), ",")
                    For k = 0 To UBound(strCammaArray())
                        If k = 0 Then
                            ReadFileBaitai(i).strKomokuName = strCammaArray(k)
                        ElseIf k = 1 Then
                            ReadFileBaitai(i).strGoukei = strCammaArray(k)
                        Else
                            ReadFileBaitai(i).srtGoukiValue(l) = strCammaArray(k)
                            l = l + 1
                        End If
                    Next k
            End Select
            l = 0
        Next i
    Else
        For i = 0 To j - 1
            Select Case i
                Case 0, 1, 2, 3, 4, 22, 23    'ICRIYOxxxx.csvΜ1`4sΪάΕΝ^CgΘΜΕAΪΌΙZbg
                    ReadFileBaitai(i).strKomokuName = strLineCount(i)
                Case Else
                    'J}ζΨθπ1ΪΈΒζθo·B
                    strCammaArray = Split(strLineCount(i), ",")
                    For k = 0 To UBound(strCammaArray())
                        If k = 0 Then
                            ReadFileBaitai(i).strKomokuName = strCammaArray(k)
                        ElseIf k = 1 Then
                            ReadFileBaitai(i).strGoukei = strCammaArray(k)
                        Else
                            ReadFileBaitai(i).srtGoukiValue(l) = strCammaArray(k)
                            l = l + 1
                        End If
                    Next k
            End Select
            l = 0
        Next i
    End If

    Print #intJprFile, "έuR[iF" & gstrCornerName(iCornerIdx)
    Print #intJprFile, ""
    
    If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
        Print #intJprFile, "yΚίf[^z"
    Else
        Print #intJprFile, "yhbJ[hpΰzf[^z"
    End If
    '/////////////////////
    'όD€f[^ΜoΝ
    Print #intJprFile, "όD€Κίv"
    
    For i = 0 To MAX_KOMOKU_NUM_TUKA - 1
        If RTrim(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName) = "0" Then
            'ΪΌΙ0ͺZbg³κΔ’½ηΘ~ΝoΝ΅Θ’
            Exit For
        Else
            'χΨItCW[iΖ νΉι½ίΜαO
            If RTrim(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName) = "»ΜΌIC (¬)" Then
                ReadFileBaitai(i + iStartLineKaisatu).strKomokuName = "»ΜΌIC(¬)" & Space(38)   'Xy[Xπ­
            End If
        
            Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName, vbFromUnicode), 20), vbUnicode) _
            & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLineKaisatu).strGoukei), "#,0"), 10)
        End If
    Next i
    Print #intJprFile, ""
    
    '/////////////////////
    'WD€f[^ΜoΝ
    Print #intJprFile, "WD€Κίv"
    
    For i = 0 To MAX_KOMOKU_NUM_TUKA - 1
        If RTrim(ReadFileBaitai(i + iStartLineShusatu).strKomokuName) = "0" Then
            'ΪΌΙ0ͺZbg³κΔ’½ηΘ~ΝoΝ΅Θ’
            Exit For
        Else
            'χΨItCW[iΖ νΉι½ίΜαO
            If RTrim(ReadFileBaitai(i + iStartLineShusatu).strKomokuName) = "»ΜΌIC (¬)" Then
                ReadFileBaitai(i + iStartLineShusatu).strKomokuName = "»ΜΌIC(¬)" & Space(38)    'Xy[Xπ­
            End If
        
            Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLineShusatu).strKomokuName, vbFromUnicode), 20), vbUnicode) _
            & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLineShusatu).strGoukei), "#,0"), 10)
        End If
    Next i
    Print #intJprFile, ""
        
    
    'Print #intJprFile, FOOTER_STRING
    'Close #intJprFile
    
    Exit Sub
    
'G[
Err_handler:

    If intOutFile > 0 Then
        Close #intOutFile
    End If
    
    If intJprFile > 0 Then
        Close #intJprFile
    End If

    Set fso = Nothing
    'G[OΜoΝ
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  ΦΌΜ  : psMakeTukaImageFileKan
'//  @\ΌΜ  : Κίf[^W[iΜC[Wt@Cμ¬i²όpj
'//  @\Tv  : Κίf[^W[iΜC[Wt@Cπμ¬·ιB
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Integer   iCornerIdx   R[iCfbNX
'//
'//
'//              ^        l           Σ‘
'//  ίθl    : ³΅
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub psMakeTukaImageFileKan(iCornerIdx As Integer)
    
    Dim strBaitaiFileName   As String                       '}ΜoΝt@C TUKAR[iΌYYYYMMDDhhmmss.csv ICRIYOR[iΌYYYYMMDDhhmmss.csv
    Dim ReadFileBaitai()    As BAITAI_OUTPUT_IMAGE_FILE_KAN '}ΜoΝt@C
    Dim intOutFile          As Integer                      'oΝt@CΤ
    Dim strLineCount()      As String
    Dim fso                 As New FileSystemObject
    Dim FsoTS               As TextStream
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    Dim l                   As Integer
    Dim strCammaArray()     As String                       'J}ζΨθΕ1ΪΈΒζθo΅½f[^
    Dim iKomokuMaxCnt       As Integer                      ' Wvf[^ΪΜΕε
    Dim iStartLine          As Integer                      'eWvubNΜJns
                                                                
    On Error GoTo Err_handler
    
    'eWvΪΜoΝJnΚuπζΎiINIt@CΙζθoΝL³ͺwθΕ«ι½ίAJnΚuΝΒΟΙΘιj
    
    strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "TUKA" & gstrStationName(iCornerIdx) & gstrCornerName(iCornerIdx) & "*.csv")
           
    '////////////////////////////////////////////////
    '// Κίf[^/pΰzΜ}ΜoΝt@CπζΎ
    't@CΤζΎ
    'wΌΜ{R[iΌΜyyyymmddhhmmss.csv
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForAppending)
    j = FsoTS.Line
    FsoTS.Close
           
    ReDim strLineCount(j) As String         'CSVt@Cπ1sΈΒόκΔ¨­
           
    i = 0
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForReading)
    Do Until FsoTS.AtEndOfStream = True
        strLineCount(i) = FsoTS.ReadLine
        i = i + 1
    Loop
    FsoTS.Close
    Set fso = Nothing
    
    '}ΜoΝt@CC[W\’ΜΙZbg·ι
    ReDim ReadFileBaitai(j) As BAITAI_OUTPUT_IMAGE_FILE_KAN     't@CΗpGA
    l = 0
    
    For i = 0 To j - 1
        If InStr(strLineCount(i), ",") = 0 Then
            'J}ζΨθΙΘΑΔ’Θ’sΝΪΌΙΖθ ¦Έf[^πZbg
            ReadFileBaitai(i).strKomokuName = strLineCount(i)
        Else
            'J}ζΨθπ1ΪΈΒζθo·B
            strCammaArray = Split(strLineCount(i), ",")
            For k = 0 To UBound(strCammaArray())
                If k = 0 Then
                    ReadFileBaitai(i).strKomokuName = strCammaArray(k)
                ElseIf k = 1 Then
                    ReadFileBaitai(i).strGoukei = strCammaArray(k)
                ElseIf k = 2 Then
                    ReadFileBaitai(i).strNorikae = strCammaArray(k)
                ElseIf k = 3 Then
                    ReadFileBaitai(i).strTukaChoku = strCammaArray(k)
                Else
                    ReadFileBaitai(i).srtGoukiValue(l) = strCammaArray(k)
                    l = l + 1
                End If
            Next k
        End If
        l = 0
    Next i

    Print #intJprFile, "έuR[iF" & gstrCornerName(iCornerIdx)
    Print #intJprFile, ""
    
    Print #intJprFile, "yiqV²όΚίf[^z"
    
    '//////////////////////////////////////////////////////////
    'όD€ V²όΚίf[^ΜoΝ
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_KAISATU_KANSEN_TUKA)
        Print #intJprFile, "όD€@V²όΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_TUKA_KAN - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    'WD€@V²όΚίf[^ΜoΝ
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_SHUSATU_KANSEN_TUKA)
        
        Print #intJprFile, "WD€@V²όΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_TUKA_KAN - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '^ss\Κίf[^ΜoΝ
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_UNKOU_FUNOU) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_IC_UNKO_FUNOU)
    
        Print #intJprFile, "^ss\Κίv"
        
        For i = 0 To MAX_KOMOKU_NUM_UNKOU_FUNOU - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '²\έζ·@έόΚίf[^ΜoΝ
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA_KAN_ZAI) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_KAN_ZAI_TUKA)

        Print #intJprFile, "²|έζ·@έόΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_NORIKAE_TUKA - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    'έ\²ζ·@έόΚίf[^ΜoΝ
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA_ZAI_KAN) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_ZAI_KAN_TUKA)
        
        Print #intJprFile, "έ|²ζ·@έόΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_NORIKAE_TUKA - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '©wόκ~Οf[^ΜoΝ
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_KYUSAI) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_JIEKI_KYUSAI)
    
        Print #intJprFile, "©wόκ~ΟΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_JIEKI_KYUSAI - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    '₯Cρϋ~Κίf[^ΜoΝ
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_KAISHU_CHUSI) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_KAISHU_CHUSHI)
    
        Print #intJprFile, "₯Cρϋ~Κίv"
        
        For i = 0 To MAX_KOMOKU_NUM_JIKI_KAISHU_CHUSHI - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    'Print #intJprFile, FOOTER_STRING
    'Close #intJprFile
    
    Exit Sub
    
'G[
Err_handler:

    If intOutFile > 0 Then
        Close #intOutFile
    End If
    
    If intJprFile > 0 Then
        Close #intJprFile
    End If

    Set fso = Nothing
    'G[OΜoΝ
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  ΦΌΜ  : psMakeRiyoImageFileKan
'//  @\ΌΜ  : pΰzf[^W[iΜC[Wt@Cμ¬i²όpj
'//  @\Tv  : pΰzf[^W[iΜC[Wt@Cπμ¬·ιB
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Integer   iCornerIdx   R[iCfbNX
'//
'//
'//              ^        l           Σ‘
'//  ίθl    : ³΅
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Sub psMakeRiyoImageFileKan(iCornerIdx As Integer)
    
    Dim strBaitaiFileName   As String                       '}ΜoΝt@C TUKAR[iΌYYYYMMDDhhmmss.csv ICRIYOR[iΌYYYYMMDDhhmmss.csv
    Dim ReadFileBaitai()    As BAITAI_OUTPUT_IMAGE_FILE_KAN '}ΜoΝt@C
    Dim intOutFile          As Integer                      'oΝt@CΤ
    Dim strLineCount()      As String
    Dim fso                 As New FileSystemObject
    Dim FsoTS               As TextStream
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    Dim l                   As Integer
    Dim strCammaArray()     As String                       'J}ζΨθΕ1ΪΈΒζθo΅½f[^
    Dim iKomokuMaxCnt       As Integer                      ' Wvf[^ΪΜΕε
    Dim iStartLine          As Integer                      'eWvubNΜJns
                                                                
    On Error GoTo Err_handler
    
    'eWvΪΜoΝJnΚuπζΎiINIt@CΙζθoΝL³ͺwθΕ«ι½ίAJnΚuΝΒΟΙΘιj
    
    strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "ICRIYO" & gstrStationName(iCornerIdx) & gstrCornerName(iCornerIdx) & "*.csv")
           
    '////////////////////////////////////////////////
    '// Κίf[^/pΰzΜ}ΜoΝt@CπζΎ
    't@CΤζΎ
    'wΌΜ{R[iΌΜyyyymmddhhmmss.csv
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForAppending)
    j = FsoTS.Line
    FsoTS.Close
           
    ReDim strLineCount(j) As String         'CSVt@Cπ1sΈΒόκΔ¨­
           
    i = 0
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForReading)
    Do Until FsoTS.AtEndOfStream = True
        strLineCount(i) = FsoTS.ReadLine
        i = i + 1
    Loop
    FsoTS.Close
    Set fso = Nothing
    
    '}ΜoΝt@CC[W\’ΜΙZbg·ι
    ReDim ReadFileBaitai(j) As BAITAI_OUTPUT_IMAGE_FILE_KAN     't@CΗpGA
    l = 0
    
    For i = 0 To j - 1
        If InStr(strLineCount(i), ",") = 0 Then
            'J}ζΨθΙΘΑΔ’Θ’sΝΪΌΙΖθ ¦Έf[^πZbg
            ReadFileBaitai(i).strKomokuName = strLineCount(i)
        Else
            'J}ζΨθπ1ΪΈΒζθo·B
            strCammaArray = Split(strLineCount(i), ",")
            For k = 0 To UBound(strCammaArray())
                If k = 0 Then
                    ReadFileBaitai(i).strKomokuName = strCammaArray(k)
                ElseIf k = 1 Then
                    ReadFileBaitai(i).strGoukei = strCammaArray(k)
                ElseIf k = 2 Then
                    ReadFileBaitai(i).strNorikae = strCammaArray(k)
                ElseIf k = 3 Then
                    ReadFileBaitai(i).strTukaChoku = strCammaArray(k)
                Else
                    ReadFileBaitai(i).srtGoukiValue(l) = strCammaArray(k)
                    l = l + 1
                End If
            Next k
        End If
        l = 0
    Next i

    Print #intJprFile, "έuR[iF" & gstrCornerName(iCornerIdx)
    Print #intJprFile, ""
    
    Print #intJprFile, "yiqV²όΰzf[^z"
    
    '//////////////////////////////////////////////////////////
    'όD€ εl V²όXCJΚίvΜoΝ
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAI_OTONA_SUICA_RIYO)
        Print #intJprFile, "όD€ εl V²ό½²ΆΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    'WD€ εl V²όXCJΚίvΜoΝ
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_SHU_OTONA_SUICA_RIYO)
        Print #intJprFile, "WD€ εl V²ό½²ΆΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    'όD€ ¬ V²όXCJΚίvΜoΝ
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAI_SHONI_SUICA_RIYO)
        Print #intJprFile, "όD€ ¬ V²ό½²ΆΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    'WD€ ¬ V²όXCJΚίvΜoΝ
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_SHU_SHONI_SUICA_RIYO)
        Print #intJprFile, "WD€ ¬ V²ό½²ΆΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    
    '//////////////////////////////////////////////////////////
    'XCJοΠΤΈZ^ΐx₯’Κίv
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_CARD_SHIHARAI) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_SEISAN_SHIHARAI)
        Print #intJprFile, "½²ΆοΠΤΈZ^ΐx₯Κίv"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_SEISAN - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If

    '//////////////////////////////////////////////////////////
    'όD€I[g`[WΚίv
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_AUTO_CHARGE) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAI_AUTOCHARGE)
        Print #intJprFile, "όD€ I[g`[WΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_AUTOCHARGE - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    'WD€I[g`[WΚίv
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_AUTO_CHARGE) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_SHU_AUTOCHARGE)
        Print #intJprFile, "WD€ I[g`[WΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_AUTOCHARGE - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    'V²ό^ΐ εl@XCJΚίv
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_KIKAN_KANSEN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAN_OTONA_SUICA_RIYO)
        Print #intJprFile, "²ό^ΐ εl ½²ΆΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    'V²ό^ΐ ¬@XCJΚίv
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_KIKAN_KANSEN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAN_SHONI_SUICA_RIYO)
        Print #intJprFile, "²ό^ΐ ¬ ½²ΆΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    'ζ·έ^ΐ εl@XCJΚίv
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_KIKAN_ZAIRAI) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_NORI_ZAI_OTONA_SUICA_RIYO)
        Print #intJprFile, "ζ·έ^ΐ εl ½²ΆΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    '//////////////////////////////////////////////////////////
    'ζ·έ^ΐ ¬@XCJΚίv
    '//////////////////////////////////////////////////////////
    'INIΕoΝLΙέθ³κΔ’κΞoΝ·ι
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_KIKAN_ZAIRAI) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_NORI_ZAI_SHONI_SUICA_RIYO)
        Print #intJprFile, "ζ·έ^ΐ ¬ ½²ΆΚίv"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                'ΪΌΙ0ͺZbg³κΔ’½ηoΝ΅Θ’
            Else
                'ΪΌͺ20ΙϋάηΘ’κΝΌpXy[XκΒόκΔlπoΝiΚuΝ»λ¦Θ’)
                If LenB(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode)) > 19 Then
                    Print #intJprFile, StrConv(StrConv(RTrim(ReadFileBaitai(i + iStartLine).strKomokuName), vbFromUnicode), vbUnicode) & " " _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                Else
                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLine).strKomokuName, vbFromUnicode), 20), vbUnicode) _
                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLine).strGoukei), "#,0"), 10)
                End If
            End If
        Next i
        Print #intJprFile, ""
    End If
    
    'Print #intJprFile, FOOTER_STRING
    'Close #intJprFile
    
    Exit Sub
    
'G[
Err_handler:

    If intOutFile > 0 Then
        Close #intOutFile
    End If
    
    If intJprFile > 0 Then
        Close #intJprFile
    End If

    Set fso = Nothing
    'G[OΜoΝ
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  ΦΌΜ  : pfGetStartLineTuka
'//  @\ΌΜ  : wθ΅½WvΪΜσJnΚuζΎ
'//  @\Tv  : wθ΅½WvΪΜσJnΚuπGAIBU_OUTPUT.INIΙ]ΑΔίιB
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Integer   intShukeiKoumoku     WvΪ
'//
'//
'//              ^        l           Σ‘
'//  ίθl    : Integer                JnΚu
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function pfGetStartLineTuka(intShukeiKoumoku As Integer) As Integer

    Dim intCount            As Integer
    Dim intIniIdx           As Integer  'GAIBU_OUTPUT.INIΜL[ΙΞ·ιCfbNX
    
    Dim intStartLine        As Integer  'Κίf[^ΜJnsiCSVγj
    
    Dim intNextBlockLine    As Integer  'ΜWvubNΜf[^ͺ ιΚuiCSVγj
    
    Dim intNowLine           As Integer  'INIt@CΜoΝL³Ι]ΑΔACSVt@Cπγ©ηΙ©Δ’Α½Ζ«Μ»έs
    
    intNowLine = 0
    intNextBlockLine = 6
    intIniIdx = 0
    
    For intCount = 0 To MAX_TUKA_SHUKEI_KOUMOKU - 1
    
        Select Case intCount
            Case mintTukaShukeiKoumoku.SHUKEI_KAISATU_KANSEN_TUKA              'yόD€ V²όΚίf[^z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_TUKA_KAN + 2
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_SHUSATU_KANSEN_TUKA              'yWD€@V²όΚίf[^z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_TUKA_KAN + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_IC_UNKO_FUNOU                    'y^ss\f[^z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_UNKOU_FUNOU) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_UNKOU_FUNOU + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_KAN_ZAI_TUKA                    'y²-έ ζ·Κίf[^z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA_KAN_ZAI) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_NORIKAE_TUKA + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_ZAI_KAN_TUKA                    'yέ-² ζ·Κίf[^z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA_ZAI_KAN) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_NORIKAE_TUKA + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_JIEKI_KYUSAI                    'y©wόκ~ΟΚίf[^z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_KYUSAI) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_JIEKI_KYUSAI + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_KAISHU_CHUSHI                  'y₯Cρϋ~Κίf[^z
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_KAISHU_CHUSI) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_JIKI_KAISHU_CHUSHI + 3
                End If
            Case Else   'γLΘOΝΰzf[^ΙΦ·ιέθΜ½ίXLbv
                
        End Select
        If intCount <> 0 Then
            intIniIdx = intIniIdx + 1
        End If
        
        'ί½’JnΚuΎΑ½ηA»ΜsπΤ·
        If intShukeiKoumoku = intCount Then
            pfGetStartLineTuka = intNowLine
            Exit Function
        End If
    Next
    
    'γLΜForΆπΕεράΕρΑΔIΉ΅½Ζ’€±ΖΝAί½’JnΚuͺίηκΘ©Α½B
    pfGetStartLineTuka = intNowLine

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  ΦΌΜ  : pfGetStartLineKingaku
'//  @\ΌΜ  : wθ΅½WvΪΜσJnΚuζΎ
'//  @\Tv  : wθ΅½WvΪΜσJnΚuπGAIBU_OUTPUT.INIΙ]ΑΔίιB
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Integer   intShukeiKoumoku     WvΪ
'//
'//
'//              ^        l           Σ‘
'//  ίθl    : Integer                JnΚu
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function pfGetStartLineKingaku(intShukeiKoumoku As Integer) As Integer

    Dim intCount            As Integer
    Dim intIniIdx           As Integer  'GAIBU_OUTPUT.INIΜL[ΙΞ·ιCfbNX
    
    Dim intStartLine        As Integer  'Κίf[^ΜJnsiCSVγj
    
    Dim intNextBlockLine    As Integer  'ΜWvubNΜf[^ͺ ιΚuiCSVγj
    
    Dim intNowLine           As Integer  'INIt@CΜoΝL³Ι]ΑΔACSVt@Cπγ©ηΙ©Δ’Α½Ζ«Μ»έs
    
    intNowLine = 0
    intNextBlockLine = 6
    intIniIdx = 0
    
    For intCount = 0 To MAX_KINGAKU_SHUKEI_KOUMOKU - 1
    
        Select Case intCount
            'yόD€@εl@V²όXCJpvΰzz
            'yWD€@εl@V²όXCJpvΰzz
            'yόD€@¬@V²όXCJpvΰzz
            
            'y²ό^ΐ@εl@XCJpvΰzz
            'yζ·έ^ΐ@εl@XCJpvΰzz
            Case mintKingakuShukeiKoumoku.SHUKEI_KAI_OTONA_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_SHU_OTONA_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_KAI_SHONI_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_KAN_OTONA_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_NORI_ZAI_OTONA_SUICA_RIYO
                
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_SUICA_RIYO + 2
                End If
            'yWD€@¬@V²όXCJpvΰzz
            'y²ό^ΐ@¬@XCJpvΰzz
            'yζ·έ^ΐ@¬@XCJpvΰzz
            Case mintKingakuShukeiKoumoku.SHUKEI_SHU_SHONI_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_KAN_SHONI_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_NORI_ZAI_SHONI_SUICA_RIYO
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_SUICA_RIYO + 3
                End If
            'yXCJοΠΤΈZf[^@^ΐx₯zz
            Case mintKingakuShukeiKoumoku.SHUKEI_SEISAN_SHIHARAI
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_CARD_SHIHARAI) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_SUICA_SEISAN + 3
                End If
            'yόD€@I[g`[Wf[^z
            Case mintKingakuShukeiKoumoku.SHUKEI_KAI_AUTOCHARGE
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_AUTO_CHARGE) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_AUTOCHARGE + 2
                End If
            'yWD€@I[g`[Wf[^z
            Case mintKingakuShukeiKoumoku.SHUKEI_SHU_AUTOCHARGE
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_AUTO_CHARGE) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_AUTOCHARGE + 3
                End If
            Case Else   'γLΘOΝΰzf[^ΙΦ·ιέθΜ½ίXLbv
                
        End Select
        If intCount <> 0 Then
            intIniIdx = intIniIdx + 1
        End If
        
        'ί½’JnΚuΎΑ½ηA»ΜsπΤ·
        If intShukeiKoumoku = intCount Then
            pfGetStartLineKingaku = intNowLine
            Exit Function
        End If
    Next
    
    'γLΜForΆπΕεράΕρΑΔIΉ΅½Ζ’€±ΖΝAί½’JnΚuͺίηκΘ©Α½B
    pfGetStartLineKingaku = intNowLine

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  ΦΌΜ  : pfGetSubGateCsv
'//  @\ΌΜ  : ©όβξρζΎ
'//  @\Tv  : wθ΅½R[iΜ©όβCSVt@CπζΎ·ιB
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Integer   intCornerNo   R[iΤ
'//
'//
'//              ^        l           Σ‘
'//  ίθl    : Integer                ζΎR[h
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(EG20 V30.3.0.1) 2014-09-19   CODED   BY [TCC] T.Nakajima
'//                 k€V²όtF[YQΞ yHKRK_Kansi07_008_01z
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
'Private Function pfGetSubGateCsv(intCornerNo As Integer) As Integer                                            ' EG20 V30.3.0.1 yHKRK_Kansi07_008_01z DEL
Private Function pfGetSubGateCsv(intCornerNo As Integer, intGokiNo As Integer, intKomoku As Integer) As Integer 'EG20 V30.3.0.1 yHKRK_Kansi07_008_01z ADD

    Dim intFileNumber            As Integer
    Dim i                        As Integer
    Dim ReadBuf                  As JIKAIINFO_IMAGE_FILE    'Ηέέobt@
        
    'Erase ReadSetteiSubGate        'EG20 V30.3.0.1 yHKRK_Kansi07_008_01z DEL
    
    'G[[`πιΎ
    On Error GoTo Err_handler      'EG20 V30.3.0.1 ADD
    
    't@CΤζΎ
    intFileNumber = FreeFile
    
    'CSVt@CI[v
    Open EKI_TUDO_CHK_SUBGATE_FILE For Input As #intFileNumber
    
    'κv·ιR[iΤΜR[hπGAΙΫΆ΅Δ’­
    i = 0
    Do While Not EOF(intFileNumber)
                
        Input #intFileNumber, ReadBuf.strBunrui_Dai, ReadBuf.strBunrui_Tyu, _
            ReadBuf.srtBunrui_Sho, ReadBuf.strCorner, ReadBuf.strKomoku, _
            ReadBuf.strKubun, ReadBuf.strData, ReadBuf.strSetShosai
        
        If CInt(ReadBuf.strCorner) = intCornerNo Then
            If CInt(ReadBuf.strBunrui_Tyu) = intGokiNo Then     'EG20 V30.3.0.1 yHKRK_Kansi07_008_01z ADD
                If CInt(ReadBuf.srtBunrui_Sho) = intKomoku Then     'EG20 V30.3.0.1 yHKRK_Kansi07_008_01z ADD
                    'ReDim Preserve ReadSetteiSubGate(i) As JIKAIINFO_IMAGE_FILE    'EG20 V30.3.0.1 yHKRK_Kansi07_008_01z DEL
                    'ReadSetteiSubGate(i) = ReadBuf                                 'EG20 V30.3.0.1 yHKRK_Kansi07_008_01z ADD
                    ReadSetteiSubGate((intGokiNo - 1) * SUBGATE_ITEM_NUM + (intKomoku - 1)) = ReadBuf
                    i = i + 1
                    Exit Do     '@AΪΤΕiθήζ€Ι΅½ΜΕAίθlΖΘιR[hΝO©PΗΏη©ΙΘιB EG20 V30.3.0.1 yHKRK_Kansi07_008_01z ADD
                End If      'EG20 V30.3.0.1 yHKRK_Kansi07_008_01z ADD
            End If      'EG20 V30.3.0.1 yHKRK_Kansi07_008_01z ADD
        End If
    Loop
    
    'CSVt@CN[Y
    Close #intFileNumber
    pfGetSubGateCsv = i
    
'EG20 V30.3.0.1 ADD START
    Exit Function
Err_handler:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    'ΩνOoΝ
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JPR_PRINT_OUTPUT_ERR, 0)
    
    pfGetSubGateCsv = 0
'EG20 V30.3.0.1 ADD END


End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  ΦΌΜ  : pfOutPutSubGate
'//  @\ΌΜ  : ©όβξρoΝ
'//  @\Tv  : wθ΅½R[iΜ©όβΰeπW[i`?ΕoΝ·ι
'//
'//              ^        ΌΜ         Σ‘
'//  ψ      : Integer   intCornerNo   R[iΤ
'//              Integer   intFileNumber t@CΤ
'//
'//
'//              ^        l           Σ‘
'//  ίθl    : Integer                ζΎR[h
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(EG20 V30.3.0.1) 2014-09-19   CODED   BY [TCC] T.Nakajima
'//                 k€V²όtF[YQΞ yHKRK_Kansi07_003_01zAyHKRK_Kansi07_008_01z
'//  REVISIONS :(EG30 V32.1.0.1) 2016-06-16   CODED   BY [TCC] T.Nakajima
'//                 2016Nx{τΞ
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
'Private Sub pfOutPutSubGate(intCornerNo As Integer, intFileNumber As Integer)  'EG20 V30.3.0.1 DEL
Private Function pfOutPutSubGate(intCornerNo As Integer, intFileNumber As Integer) As Boolean   'EG20 V30.3.0.1 ADD
    Dim intTitleFlg             As Integer                  '©όβΜε©o΅ΜoΝtO
    Dim intSubGateCnt           As Integer                  '©όβ1R[iͺΜR[h
    Dim i                       As Integer
    Dim intGokiLoop             As Integer                  '@1`32 EG20 V30.3.0.1       yHKRK_Kansi07_008_01z ADD
    Dim intKomokuLoop           As Integer                  '¬Ϊ@`E EG20 V30.3.0.1    yHKRK_Kansi07_008_01z ADD
    Dim intRet                  As Integer                  ' EG20 V30.3.0.1 ADD
    
    'EG30 V32.1.0.1 ADD START
    Dim bRet                    As Boolean
    Dim lErrCode                As Long
    Dim strEkiSettiBefPath      As String           '»έwέθf[^iΟXOΫΆj
    Dim strGetValue             As String * 64
    Dim strCompValue            As String           'έθliΟXOΫΆj
    Dim strChangeFlg            As String           'ΟXσ
    Dim intValueLen             As Integer          'ζΎ΅½έθlΜ·³
    'EG30 V32.1.0.1 ADD END


    '»ΜR[iΜ©όβf[^πζΎ
    intTitleFlg = 0
    'intSubGateCnt = pfGetSubGateCsv(intCornerNo)    'EG20 V30.3.0.1 yHKRK_Kansi07_003_01zDEL
    'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z ADD START
    'SUB_GATE_KAN.INI©ηR[iͺΘ­ΘΑ½½ίAR[iΝ0ΕθΕ@AΪ@`EΜΕEKI_DATA.CSV©ηυ
    intSubGateCnt = 0                       'EG20 V30.3.0.1 yHKRK_Kansi07_008_01z ADD
    For intGokiLoop = 0 To 31
        For intKomokuLoop = 0 To 5
            intRet = pfGetSubGateCsv(0, intGokiLoop + 1, intKomokuLoop + 1)
            If intRet = 0 Then
                ' CSV©ηΜζΎͺ0ΜκΝG[Ζ·ιB
                pfOutPutSubGate = False
                Exit Function
            Else
                intSubGateCnt = intSubGateCnt + intRet
            End If
        Next
    Next
    
    'EG30 V32.1.0.1 ADD START
    'R[iOΜΟXOΫΆ³κ½wsxf[^Ζδr·ιB
    '»ΜR[iΜΟXOf[^ΫΆ³κ½f[^πγΙWJ·ι
    strEkiSettiBefPath = Replace(EKI_SETTI_FILE_BEF, "#", 0)
    Call dllGetEkiIniDataBefore(strEkiSettiBefPath, lErrCode)
    'EG30 V32.1.0.1 ADD END
    
    'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z ADD END
    For i = 0 To intSubGateCnt - 1
        ' EG20 V30.3.0.1 yHKRK_Kansi07_003_01z DEL START
        ' wθ΅½R[iA@ΙΞ·ιR[hπoΝ·ιKvͺΘ­ΘθA1`32@ΕθΙΘΑ½½ίIfΆπν
        'If IsTaisyoGoki(CInt(ReadSetteiSubGate(i).strCorner), CInt(ReadSetteiSubGate(i).strBunrui_Tyu)) = True Then
        ' EG20 V30.3.0.1 yHKRK_Kansi07_003_01z DEL END
        If intTitleFlg = 0 Then
            Print #intFileNumber, ""
            'Print #intFileNumber, "yόD@@έuπ@©Πz" 'EG30 V32.1.0.1 DEL
            Print #intFileNumber, "@yόD@@έuπ@©Πz"    'EG30 V32.1.0.1 ADD
            intTitleFlg = 1
        End If
        
        'EG30 V32.1.0.1 ADD START
        'ΟXOf[^ΫΆ³κ½έθlΖδr·ι
        bRet = dllGetEkiInfoValue(CInt(ReadSetteiSubGate(i).strBunrui_Dai), _
                                    CInt(ReadSetteiSubGate(i).strBunrui_Tyu), _
                                    CInt(ReadSetteiSubGate(i).srtBunrui_Sho), _
                                    0, _
                                    strGetValue, _
                                    intValueLen)
        strCompValue = strGetValue
        If (intValueLen <> 0) Then
            strCompValue = MidByte(strGetValue, 1, intValueLen)
            strCompValue = Trim(strCompValue)
        ElseIf (intValueLen = 0) Then
            strCompValue = "0"
        End If
        
        If (bRet = False) Or (CInt(ReadSetteiSubGate(i).strData) <> CInt(strCompValue)) Then
            strChangeFlg = DIFF_MARK_STRING_ON
        Else
            strChangeFlg = DIFF_MARK_STRING_OFF
        End If
        'EG30 V32.1.0.1 ADD END
        
        'ReadSetteiSubGate(i).strData = Format(ReadSetteiSubGate(i).strData, "0#")      'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z DEL
        ReadSetteiSubGate(i).strData = Format(ReadSetteiSubGate(i).strData, "00#")      'EG20 V30.3.0.1 yHKRK_Kansi07_003_01z ADD
        Select Case CInt(ReadSetteiSubGate(i).srtBunrui_Sho)
            Case 1
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "@ " & "FM Ί°Ε°Τ" & " " & ReadSetteiSubGate(i).strData  'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "@ " & "FM Ί°Ε°Τ" & " " & ReadSetteiSubGate(i).strData    'EG30 V32.1.0.1 ADD
            Case 2
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "@ " & "FM @Τ" & " " & ReadSetteiSubGate(i).strData  'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "@ " & "FM @Τ" & " " & ReadSetteiSubGate(i).strData    'EG30 V32.1.0.1 ADD
            Case 3
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "@ " & "V²όIC Ί°Ε°Τ" & " " & ReadSetteiSubGate(i).strData  'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "@ " & "V²όIC Ί°Ε°Τ" & " " & ReadSetteiSubGate(i).strData    'EG30 V32.1.0.1 ADD
            Case 4
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "@ " & "V²όIC @Τ" & " " & ReadSetteiSubGate(i).strData  'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "@ " & "V²όIC @Τ" & " " & ReadSetteiSubGate(i).strData    'EG30 V32.1.0.1 ADD
            Case 5
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "@ " & "NRZ Ί°Ε°Τ" & " " & ReadSetteiSubGate(i).strData 'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "@ " & "NRZ Ί°Ε°Τ" & " " & ReadSetteiSubGate(i).strData   'EG30 V32.1.0.1 ADD
            Case 6
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "@ " & "NRZ @Τ" & " " & ReadSetteiSubGate(i).strData 'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "@ " & "NRZ @Τ" & " " & ReadSetteiSubGate(i).strData   'EG30 V32.1.0.1 ADD
            Case Else
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & ReadSetteiSubGate(i).strKomoku & " " & ReadSetteiSubGate(i).strData 'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & ReadSetteiSubGate(i).strKomoku & " " & ReadSetteiSubGate(i).strData 'EG30 V32.1.0.1 ADD
        End Select
            
    Next i
    pfOutPutSubGate = True
End Function
'EG20 V30.1.0.1 ADD END
'EG30 V32.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  ΦΌΜ  : pfGetSaveDate
'//  @\ΌΜ  : ΟXOf[^ΫΆϊtζΎ
'//  @\Tv  : R[i²ΖΙΫΆ³κΔ’ιSaveDate.datΜXVϊtπζΎ·ι
'//
'//              ^        ΌΜ      Σ‘
'//  ψ      : Integer    intCorner   ζΎ·ιR[iΤ
'//
'//              ^        l        Σ‘
'//  ίθl    : String     XVϊt    YYYYNMMDDϊHH:MM
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-14   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  υlF
'///////////////////////////////////////////////////////////////////
Private Function pfGetSaveDate(intCorner As Integer) As String
    Dim strFileName(0 To 1)     As String           'μ¬ϊ
    Dim intCnt                  As Integer          'JE^
    Dim lngHandle               As Long             'nh

    Dim lpCreatTime             As FILETIME         'μ¬ϊ
    Dim lpAccessTime            As FILETIME         'ΕIANZXϊ
    Dim lpLastwTime             As FILETIME         'XVϊ
    Dim lpLocalTime             As FILETIME         '[Jϊ
    Dim lpSystemTime            As SYSTEMTIME       'VXe
    Dim bRet                    As Boolean          'ίθl
    
    Dim strSaveFile             As String
    
    On Error Resume Next

           
    'ΫΆt@CΜϊtπζΎ
    strSaveFile = PATH_OPERATE & "CORNER" & CStr(intCorner) & "\\SETTEI_BEF\\" & SET_BEF_DATE_FILE
    If Dir(strSaveFile) = "" Then
        pfGetSaveDate = "    N    ϊ  :  "
        Exit Function
    Else
        't@CπI[v
        lngHandle = CreateFile(strSaveFile, GENERIC_READ, FILE_SHARE_READ, _
                                    0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)

        't@CI[vͺ³νΙsνκ½©H
        If lngHandle = INVALID_HANDLE_VALUE Then GoTo ErrorHandler
            't@C^CπGET
            bRet = GetFileTime(lngHandle, lpCreatTime, lpAccessTime, lpLastwTime)
            If bRet = False Then GoTo APIError                          'ζΎͺ³νΙsνκ½©H
        
            't@C^Cπ[J^CΙΟ·
            bRet = FileTimeToLocalFileTime(lpLastwTime, lpLocalTime)    'EG20 V2.1.0.1 ADD yMainte_03_01z
            If bRet = False Then GoTo APIError                          'Ο·ͺ³νΙsνκ½©H
        
            '[J^CπVXe^CΙΟ·
            bRet = FileTimeToSystemTime(lpLocalTime, lpSystemTime)
            If bRet = False Then GoTo APIError                          'Ο·ͺ³νΙsνκ½©H
                
            'nhΜN[Y
            Call CloseHandle(lngHandle)
        
            'μ¬ϊtπ\¦·ι (YYYYNMMDDϊhh:mm)
            pfGetSaveDate = lpSystemTime.wYear & "N" & _
                                Format(lpSystemTime.wMonth, "00") & "" & _
                                Format(lpSystemTime.wDay, "00") & "ϊ" & _
                                Format(lpSystemTime.wHour, "00") & ":" & _
                                Format(lpSystemTime.wMinute, "00")
    End If
            
    Exit Function

APIError:
    Call CloseHandle(lngHandle)             'nhΜN[Y

ErrorHandler:
    pfGetSaveDate = "    N    ϊ  :  "
    
End Function
'EG30 V32.1.0.1 ADD END
