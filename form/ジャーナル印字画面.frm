VERSION 5.00
Begin VB.Form frmJprPrint 
   BorderStyle     =   0  'なし
   Caption         =   "ジャーナル印字"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
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
      Caption         =   "設定値一覧"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   32
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Frame FraPrintKind 
      Caption         =   "印字項目指定"
      Height          =   1815
      Left            =   120
      TabIndex        =   29
      Top             =   2760
      Width           =   11655
      Begin VB.CheckBox chkJprKind 
         Caption         =   "改札機保守設定データ"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   39
         Top             =   1440
         Width           =   5055
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "駅都度データ確認(ｴﾝｺｰﾄﾞｺｰﾅ号機情報定義)"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   38
         Top             =   1080
         Width           =   5055
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "駅務機器ＩＤ"
         Height          =   255
         Index           =   7
         Left            =   8520
         TabIndex        =   37
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "締切オフライン出力"
         Height          =   255
         Index           =   6
         Left            =   8520
         TabIndex        =   36
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "稼働バージョン一覧"
         Height          =   255
         Index           =   5
         Left            =   8520
         TabIndex        =   35
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "利用金額データ"
         Height          =   255
         Index           =   4
         Left            =   5400
         TabIndex        =   34
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "通過データ"
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   33
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "駅都度データ確認(自改)"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   31
         Top             =   720
         Width           =   3255
      End
      Begin VB.CheckBox chkJprKind 
         Caption         =   "駅都度データ確認(駅情報)"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.CheckBox chkGouki 
      Caption         =   "１５号機"
      Height          =   375
      Index           =   14
      Left            =   10080
      TabIndex        =   27
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox chkGouki 
      Caption         =   "１４号機"
      Height          =   375
      Index           =   13
      Left            =   10080
      TabIndex        =   26
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chkGouki 
      Caption         =   "７号機"
      Height          =   375
      Index           =   6
      Left            =   6960
      TabIndex        =   19
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame FraGouki 
      Caption         =   "号機指定"
      Height          =   1935
      Left            =   4800
      TabIndex        =   12
      Top             =   600
      Width           =   6975
      Begin VB.CheckBox chkGouki 
         Caption         =   "１６号機"
         Height          =   375
         Index           =   15
         Left            =   5280
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "１３号機"
         Height          =   375
         Index           =   12
         Left            =   5280
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "１２号機"
         Height          =   375
         Index           =   11
         Left            =   3600
         TabIndex        =   24
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "１１号機"
         Height          =   375
         Index           =   10
         Left            =   3600
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "１０号機"
         Height          =   375
         Index           =   9
         Left            =   3600
         TabIndex        =   22
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "９号機"
         Height          =   375
         Index           =   8
         Left            =   3600
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "８号機"
         Height          =   375
         Index           =   7
         Left            =   2160
         TabIndex        =   20
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "６号機"
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "５号機"
         Height          =   375
         Index           =   4
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "４号機"
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "３号機"
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "２号機"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkGouki 
         Caption         =   "１号機"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FraCorner 
      Caption         =   "コーナ指定"
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4455
      Begin VB.CheckBox chkCorner 
         Caption         =   "コーナ６"
         Height          =   225
         Index           =   5
         Left            =   2640
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "コーナ５"
         Height          =   225
         Index           =   4
         Left            =   2640
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "コーナ４"
         Height          =   225
         Index           =   3
         Left            =   2640
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "コーナ３"
         Height          =   225
         Index           =   2
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "コーナ２"
         Height          =   225
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkCorner 
         Caption         =   "コーナ１"
         Height          =   225
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "印字"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "データ収集・出力  画面へ戻る"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      Alignment       =   2  '中央揃え
      BackColor       =   &H00800000&
      Caption         =   "ジャーナル印字"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
'//  ファイル名  ：frmJprPrint.frm
'//  パッケージ名：ジャーナル印字画面
'/
'//  概要：システム初期化(監視盤)画面
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-18   CODED   BY [TCC] T.Takajima
'//     REVISIONS :(EG20 V7.4.0.1) 2013-07-22   CODED   BY [TCC] T.Nakajima
'//                 日またがり出場フリー設定画面対応
'//     REVISIONS :(EG20 V30.3.0.1) 2014-09-19  CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応
'//                 【HKRK_Kansi07_003_01】 SUB_GATE_KAN.INIフォーマット見直し対応
'//     REVISIONS :(EG30 V32.1.0.1) 2016-06-10  CODED   BY [TCC] T.Nakajima
'//                 2016年度施策対応
'//     REVISIONS :(EG30 V32.2.0.1) 2016-07-20  CODED   BY [TCC] T.Nakajima
'//                 2016年度施策対応 プログレスバー非表示対応
'//     REVISIONS :(EG30 V35.3.0.1) 2019-07-03   REVISED BY [TCC] H.Kondoh
'//                2019年度施策対応
'//     REVISIONS :
'//
'//  備考：
'///////////////////////////////////////////////////////////////////
Option Explicit

'初期化実行フラグ
Private bSysFormat As Boolean

Private Const APL_INTERVAL = 390000     'アプリ起動タイマデフォルト値
Public glbFilePath  As String             'ファイルパス     'V1.12.0.1 ADD
Dim lngMAX_Time As Long                    'INI取得設定値
Dim lngtime     As Long                    '現在タイマ値
Private iSendType As Integer            '要求種別値
'V1.5.0.1 ADD END
'V1.3.0.1 ADD START
Private Const MN_MAIL_INTERVAL = 1000   'メールタイマのインターバル値
'V1.20.0.1 ADD START
Private Const LOG_INTERVAL = 30000        'ログ起動タイマデフォルト値(30秒)
Dim lngLogMAX_Time As Long                'INI取得設定値(ログ）
'V1.20.0.1 ADD END
Dim intJprFile        As Integer        'EG20 V30.1.0.1 ADD


' ジャーナル出力設定情報
Private Type JPR_PRINT_SETTING_INFO
    iCornerCount        As Integer          ' チェックされたコーナ数
    iCorner(5)          As Integer          ' チェックされたコーナ一覧
    iGoukiCount         As Integer          ' チェックされた号機数
    iGouki(15)          As Integer          ' チェックされた号機一覧
    iJprCount           As Integer          ' チェックされたジャーナル種類数
'    iJprKind(7)         As Integer          ' チェックされたジャーナル一覧      'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL
'    iJprKind(8)         As Integer          ' チェックされたジャーナル一覧      'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD    'EG30 V32.1.0.1 DEL
    iJprKind(9)         As Integer          ' チェックされたジャーナル一覧      'EG30 V32.1.0.1 ADD
End Type
Private Enum JPR_KIND
    JPR_KIND_EKI_INFO = 0           ' 駅都度データ確認(駅情報)
    JPR_KIND_JIKAI_INFO = 1         ' 駅都度データ確認(自改)
    JPR_KIND_SETTING_LST = 2        ' 設定値一覧
    JPR_KIND_TUKA_DATA = 3          ' 通過データ
    JPR_KIND_RIYO_KINGAKU = 4       ' 利用金額データ
    JPR_KIND_KADO_VER = 5           ' 稼動バージョン一覧
    JPR_KIND_SIMEKIRI = 6           ' 締切オフライン出力
    JPR_KIND_EKIMU_ID = 7           ' 駅務機器ID
    JPR_KIND_SUBGATE_INFO = 8       ' 駅都度データ確認(エンコードコーナ号機情報定義)    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
    JPR_KIND_GATE_CFG = 9          ' 改札機保守設定データ  'EG30 V32.1.0.1 ADD
End Enum
Dim udtJprPrintSetteingInfo    As JPR_PRINT_SETTING_INFO
Dim udtInitJprSetting           As JPR_PRINT_SETTING_INFO
Dim iJprIdx                     As Integer          '処理中のジャーナル

'機器構成データ（駅情報）イメージファイル読取用の構造体
Private Type EKIINFO_IMAGE_FILE
    sType       As String                '種別
    sGoki       As String                '号機
    sNo         As String                '種別毎通番
    sCorner     As String                'コーナ        ' EG20 V2.1.0.1[Mainte_03_01 駅都度対応]追加
    sTuuban     As String                '通番
    sKoumoku    As String                '項目
    sKubun      As String                '区分
    sSettei     As String                '設定値
    sSyosai     As String                '設定値詳細
End Type

'駅都度データ（改札機)イメージファイル読み取り用の構造体
Private Type JIKAIINFO_IMAGE_FILE
    strBunrui_Dai  As String               '大分類
    strBunrui_Tyu  As String               '中分類
    srtBunrui_Sho   As String               '小分類
    strCorner       As String               'コーナ
    strKomoku       As String               '項目
    strKubun        As String               '区分
    strData         As String               'データ
    strSetShosai    As String               '詳細
    
End Type

'駅都度データ確認(自改)ジャーナル出力ファイル作成テーブル
Private Type JIKAI_JPREDIT_TBL
    strKomoku       As String               '項目名（＋区分)
    strBunrui_Sho   As String               'その項目を指す小分類コード
    strKubun        As String               'その項目を指す区分
End Type

'稼動バージョン出力区分
Private Enum mintDispDiv
    KADOVER_FILE_DISP = 0
    KADOVER_FILE_OUTPUT
End Enum

'媒体出力ファイル読み取り用の構造体(通過/利用金額)
Private Type BAITAI_OUTPUT_IMAGE_FILE
    strKomokuName       As String          '項目名
    strGoukei           As String          '通過合計
    srtGoukiValue(15)   As String          '号機別の値(未使用)
End Type

'EG20 V30.1.0.1 ADD START
'媒体出力ファイル読み取り用の構造体(通過/利用金額)【幹線用】
Private Type BAITAI_OUTPUT_IMAGE_FILE_KAN
    strKomokuName       As String          '項目名
    strGoukei           As String          '通過合計
    strNorikae          As String          '通過乗換(未使用)
    strTukaChoku        As String          '通過直接(未使用)
    srtGoukiValue(31)   As String          '号機別の値(未使用)
End Type
'EG20 V30.1.0.1 ADD END

'設定一覧ファイル読み取り用構造体(OPERATE_SET##.CSV）
Private Type SETTEI_OUTPUT_IMAGE_FILE
    strDaiKomoku        As String           '大項目名
    strKomoku           As String           '項目名
    strValue            As String           '設定値
    strChangeFlg        As String           '変更フラグ 'EG30 V32.1.0.1 ADD
End Type

'稼働バージョンファイル読み取り用構造体(KadoVerDisp.csv)
Private Type KADO_VER_DISP_IMAGE_FILE
    strKishu            As String           '機種分類（ファイル読み込み用）
    strCorner           As String           'コーナ分類（ファイル読み込み用）
    strGokiDiv          As String           '号機分類（ファイル読み込み用）
    strName             As String           '機種名（ファイル読み込み用）
    strMaker            As String           'メーカ名（ファイル読み込み用）
    strVer              As String           'バージョン（ファイル読み込み用）
    strDate             As String           '作成日付（ファイル読み込み用）
End Type

'EG30 V32.1.0.1 ADD START
'改札機保守設定データ 読み取り用の構造体(JP_CFGコーナ号機番号.csv)
Private Type GATE_CFG_DATA_FILE
    strInfoName         As String           '情報部名
    strBunrui_Dai       As String           '大項目
    strBunrui_Chu       As String           '中分類
    strBunrui_Syo       As String           '小分類
    strValue            As String           '設定値
    strChangeFlg        As String           '変更有無フラグ
End Type
'EG30 V32.1.0.1 ADD END


'ジャーナル編集中間ファイル
Private Const EKIMU_DEFU = "APL\APL_WORK"
Private Const EDIT_DATA_EKIINFO = PATH_WORK & "EKI_DISP_EKIINFO.csv"    '駅都度データ確認(駅情報)"
Private Const EDIT_DATA_JIKAIINFO = PATH_WORK & "EKI_DISP_GATE_JPR.csv" '駅都度データ確認(自改)"
Private Const EDIT_DATA_SETTEI = PATH_WORK & "OPERATE_SET##.csv"        '設定値一覧
Private Const EDIT_DATA_KADOVERSION = PATH_WORK & "KadoVerDisp####"     '稼動バージョン一覧（KadoVerDispコーナ番号、号機番号）
Private Const EDIT_DATA_SIMEKIRI = PATH_WORK & "SIME##.txt"             '締切オフライン出力
Private Const EDIT_DATA_EKIMUID = PATH_WORK & "MN_VERSI.txt"            '駅務機器ID
Private Const EDIT_DATA_TUKA = PATH_SHUKEI_SEND & "TUKA*.csv"           '通過データ
Private Const EDIT_DATA_RIYO = PATH_SHUKEI_SEND & "ICRIYO*.csv"         '利用金額データ
Private Const EDIT_DATA_GATECFG = PATH_WORK & "JP_CFG####"              '改札機保守設定データ  'EG30 V32.1.0.1 ADD
Private Const APL = "APL"
Private Const LOG = "LOG"
Private Const Data = "DATA"
Private Const BACKUP = "BACKUP"

Private Const MAX_KOMOKU_NUM_TUKA = 51                      '通過外部媒体最大項目数
Private Const MAX_KOMOKU_NUM_KINGAKU = 16                   '金額外部媒体最大項目数
'EG20 V30.1.0.1 ADD START
Private Const MAX_TUKA_SHUKEI_KOUMOKU = 7                                 '幹線通過データの最大集計項目数（ブロック単位）
Private Const MAX_KOMOKU_NUM_TUKA_KAN = 51                                '幹線通過データ 最大項目数
Private Const MAX_KOMOKU_NUM_UNKOU_FUNOU = 1                              '幹線通過データ 運行不能データ 最大項目数
Private Const MAX_KOMOKU_NUM_NORIKAE_TUKA = 51                            '幹線 乗換 在来線通過データ 最大項目数
Private Const MAX_KOMOKU_NUM_JIEKI_KYUSAI = 51                            '幹線 自駅入場救済通過データ 最大項目数
Private Const MAX_KOMOKU_NUM_JIKI_KAISHU_CHUSHI = 51                      '幹線 磁気券回収中止通過データ 最大項目数

Private Const MAX_KINGAKU_SHUKEI_KOUMOKU = 11                             '幹線金額データの最大集計項目数（ブロック単位）
Private Const MAX_KOMOKU_NUM_SUICA_RIYO = 11                              '幹線金額データ スイカ利用金額　最大項目数
Private Const MAX_KOMOKU_NUM_SUICA_SEISAN = 32                            '幹線金額データ スイカ会社間精算データ 最大項目数
Private Const MAX_KOMOKU_NUM_AUTOCHARGE = 34                              '幹線金額データ オートチャージデータ　最大項目数


'集計項目（通過データ）
Private Enum mintTukaShukeiKoumoku
    SHUKEI_KAISATU_KANSEN_TUKA = 0      '【改札側　新幹線通過データ】
    SHUKEI_SHUSATU_KANSEN_TUKA          '【集札側　新幹線通過データ】
    SHUKEI_IC_UNKO_FUNOU                '【運行不能データ】
    SHUKEI_KAN_ZAI_TUKA                 '【幹-在乗換通過データ】
    SHUKEI_ZAI_KAN_TUKA                 '【在-幹乗換通過データ】
    SHUKEI_JIEKI_KYUSAI                 '【自駅入場救済通過データ】
    SHUKEI_KAISHU_CHUSHI                '【磁気券回収中止通過データ】
End Enum

'集計項目（金額データ）
Private Enum mintKingakuShukeiKoumoku
    SHUKEI_KAI_OTONA_SUICA_RIYO         '【改札側　大人　新幹線スイカ利用合計金額】
    SHUKEI_SHU_OTONA_SUICA_RIYO         '【集札側　大人　新幹線スイカ利用合計金額】
    SHUKEI_KAI_SHONI_SUICA_RIYO         '【改札側　小児　新幹線スイカ利用合計金額】
    SHUKEI_SHU_SHONI_SUICA_RIYO         '【集札側　小児　新幹線スイカ利用合計金額】
    SHUKEI_SEISAN_SHIHARAI              '【スイカ会社間精算データ　運賃支払額】
    SHUKEI_KAI_AUTOCHARGE               '【改札側　オートチャージデータ】
    SHUKEI_SHU_AUTOCHARGE               '【集札側　オートチャージデータ】
    SHUKEI_KAN_OTONA_SUICA_RIYO         '【幹線運賃　大人　スイカ利用合計金額】
    SHUKEI_KAN_SHONI_SUICA_RIYO         '【幹線運賃　小児　スイカ利用合計金額】
    SHUKEI_NORI_ZAI_OTONA_SUICA_RIYO    '【乗換在来運賃　大人　スイカ利用合計金額】
    SHUKEI_NORI_ZAI_SHONI_SUICA_RIYO    '【乗換在来運賃　小児　スイカ利用合計金額】
End Enum

'GAIBU_OUTPUT.INIのキー番号
Private Enum mintGaibuOutputKey
    GAIBU_INI_TUKA = 0                  '通過データ
    GAIBU_INI_ICSF_KIKAN                'ICSF発行期間別利用金額データ
    GAIBU_INI_IC_CARD_SHIHARAI          'ICカード会社間精算データ（運賃支払額）
    GAIBU_INI_AUTO_CHARGE               'オートチャージデータ
    GAIBU_INI_IC_UNKOU_FUNOU            'ICカード運行不能処理データ
    GAIBU_INI_TUKA_KAN_ZAI              '幹-在乗換通過データ
    GAIBU_INI_TUKA_ZAI_KAN              '在-幹乗換通過データ
    GAIBU_INI_IC_KIKAN_KANSEN           '幹線運賃IC発行機関別利用金額データ
    GAIBU_INI_IC_KIKAN_ZAIRAI           '乗換在来運賃IC発行期間別利用金額データ
    GAIBU_INI_KYUSAI                    '自駅入場救済通過データ
    GAIBU_INI_KAISHU_CHUSI              '磁気券回収中止通過データ
End Enum

'Private ReadSetteiSubGate()             As JIKAIINFO_IMAGE_FILE     'EKI_DISP_SUBGATE.CSVの1コーナ分のデータ    'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 DEL
'EG20 V30.1.0.1 ADD END
Private ReadSetteiSubGate(0 To 191)           As JIKAIINFO_IMAGE_FILE     'EKI_DISP_SUBGATE.CSV 1～32号機 ①～⑥（固定）     'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD
Private Const SUBGATE_ITEM_NUM = 6      ' SUB_GATE_KAN.INIの自社分の項目数：6                                   'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD

Private Const MAX_JPR_KETA_MAX = 30 'JPR1行最大30バイト(半角30文字)

Private Const MAX_KADO_PG = 6       '改札機1台当たりに入るプログラム数（プログラム判定データ)

Private Const FOOTER_STRING = "*************END**************"


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : cmdPrint_Click
'//  機能名称  : 「印字」釦押下時処理
'//  機能概要  : 印刷処理を実行する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17  CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdPrint_Click()
    Dim i       As Integer
    Dim bRet    As Boolean
    Dim intCount    As Integer
    
    '「ジャーナル印字画面：印字開始」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JPR_PRINT_BUTTON, 0)
    
    'ボタン、チェックボックスを非アクティブにする
    Call JPRScreenEnable(False)
    
    ' 設定状態を取得する
    Call GetPrintSettings
    
    ' コーナチェック
    If udtJprPrintSetteingInfo.iCornerCount = 0 Then
        'コーナに何もチェックされていないので処理失敗
        LstStatus.AddItem "コーナがチェックされていません"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        Call JPRScreenEnable(True)
        Exit Sub
    End If
    
    '号機チェック
    If udtJprPrintSetteingInfo.iGoukiCount = 0 Then
        '号機に何もチェックされていないので処理失敗
        LstStatus.AddItem "号機がチェックされていません"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        Call JPRScreenEnable(True)
        Exit Sub
    End If
    
    '印字項目指定チェック
    If udtJprPrintSetteingInfo.iJprCount = 0 Then
        'コーナに何もチェックされていないので処理失敗
        LstStatus.AddItem "印字項目が指定されていません"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        Call JPRScreenEnable(True)
        Exit Sub
    End If
    
    'チェックされたコーナは設置されているかいないかの情報をセットしておく
    Erase glngTergetCorner
    For intCount = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        'そのコーナが設置されているか？
        If IsTaisyoCorner(udtJprPrintSetteingInfo.iCorner(intCount)) = True Then
            glngTergetCorner(udtJprPrintSetteingInfo.iCorner(intCount) - 1) = CMN_ONOFF.CMN_ON
        Else
            glngTergetCorner(udtJprPrintSetteingInfo.iCorner(intCount) - 1) = CMN_ONOFF.CMN_OFF
        End If
    Next intCount
    
    '印字項目チェックに沿って編集処理を呼び出す。
    iJprIdx = 0
    Call JprOutputProc
 
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : JPREdit_EkiInfo
'//  機能名称  : 駅都度データ確認(駅情報)イメージファイル作成
'//  機能概要  : 駅都度データ確認(駅情報)のジャーナルイメージファイルを作成する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-27   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG30 V32.1.0.1) 2016-06-15  REVISED BY [TCC] T.Nakajima
'//                 2016年度施策対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function JPREdit_EkiInfo() As Boolean

    Dim strFileName          As String          'ファイル名
    Dim iResponse            As Integer         'MsgBox戻り値
    Dim lRetVal              As Long            '戻り値
    Dim sCommand             As String          'コマンド文字列
'V1.12.0.1 ADD START
    Dim sWriteDir            As String              '書き込み先フォルダ名
    Dim intFileNumber        As Integer             'ファイルポインタ
    Dim strLineCount         As String              '行数カウンタ
    Dim i                    As Integer             'ループカウンタ１
    Dim j                    As Integer             'ループカウンタ２
    Dim k                    As Integer             'ループカウンタ３
    Dim l                    As Integer             'ループカウンタ４
    Dim ReadFileSettei()     As EKIINFO_IMAGE_FILE  'ファイル読込用構造体
    Dim fso         As New FileSystemObject         'ファイルシステムオブジェクト
    Dim FsoTS As TextStream

    Dim bRet                 As Boolean         '関数戻り値
    Dim lErrCode             As Long            'エラーコード
    
    Dim strNowType          As String           '処理中大分類
    Dim strNowShoNo         As String           '処理中項目番号
    Dim strNowTuban         As String           '処理中項目通番
    Dim strNowCorner        As String           '処理中コーナ
    Dim strNowKubn          As String           '処理中区分
    
    Dim intCount            As Integer          'コーナインデックス ０：コーナ1
    
    'EG30 V32.1.0.1 ADD START
    Dim strEkiSettiBefPath  As String           '現在駅設定データ（変更前保存）
    Dim strGetValue         As String * 64      'DLLによって設定されるため、64固定長にしている
    Dim strCompValue        As String           '設定値（変更前保存）
    Dim strChangeFlg        As String           '変更印
    Dim intValueLen         As Integer          '取得した設定値の長さ
    'EG30 V32.1.0.1 ADD END
    
    On Error GoTo Err_handler
    
    
    '画面で指定されたコーナは設置されているか？
    If pfSettingCheck(False) = False Then
        'すべて未設置のコーナなのでエラーとする
        GoTo Err_handler
    End If
    
    '////////////////////////////////////////////////
    '// コーナ名を一通り取得
    gsGetCornerName
   
    'テストでとりあえず、コーナ1
    intCount = 0
    
    '駅都度データ確認（駅情報）イメージファイル作成
    bRet = dllGetEkiIniData(0, EKI_TUDO_CHK_EKI_INFO_FILE, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        '駅都度データ確認（駅情報）イメージファイル削除
        Kill EKI_TUDO_CHK_EKI_INFO_FILE
        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)
        JPREdit_EkiInfo = False
        Exit Function
    End If
    
    'CSVファイルの件数取得
    intFileNumber = FreeFile
    Open EKI_TUDO_CHK_EKI_INFO_FILE For Input As #intFileNumber
    
    Do While Not EOF(intFileNumber)                     ' EG20 V3.3.0.1追加
        Line Input #intFileNumber, strLineCount
        j = j + 1
    Loop
    'CSVファイルクローズ
    Close #intFileNumber
    
    '上記件数分、メモリ上に保持
    '再設定
    ReDim ReadFileSettei(j) As EKIINFO_IMAGE_FILE   'ファイル読込用エリア
        
    'CSVファイルオープン
    Open EKI_TUDO_CHK_EKI_INFO_FILE For Input As #intFileNumber

    'リスト表示分読み込み（ファイル終端までループを繰り返す）
    For i = 0 To UBound(ReadFileSettei) - 1
        Input #intFileNumber, ReadFileSettei(i).sType, ReadFileSettei(i).sGoki, ReadFileSettei(i).sNo, _
        ReadFileSettei(i).sCorner, ReadFileSettei(i).sTuuban, ReadFileSettei(i).sKoumoku, ReadFileSettei(i).sKubun, _
        ReadFileSettei(i).sSettei, ReadFileSettei(i).sSyosai
    Next i

    'CSVファイルクローズ
    Close #intFileNumber
    
    'EG30 V32.1.0.1 ADD START
    'そのコーナの変更前データ保存されたデータをメモリ上に展開する
    strEkiSettiBefPath = Replace(EKI_SETTI_FILE_BEF, "#", "0")
    Call dllGetEkiIniDataBefore(strEkiSettiBefPath, lErrCode)
    'EG30 V32.1.0.1 ADD END
    
    '/////////////////////////////////////
    'ジャーナルイメージファイル作成
    '未使用のファイル番号取得
    intFileNumber = FreeFile
   
    'ジャーナル出力イメージファイルを作成
    Open EKI_JPR_EKIINFO_TXTFILE For Output As #intFileNumber
    
    'タイトル表示
    'PrintHeader intFileNumber, "駅都度データ確認（駅情報）"    'EG30 V32.1.0.1 DEL
    PrintHeader3 intFileNumber, "駅都度データ確認（駅情報）", pfGetSaveDate(0) 'EG30 V32.1.0.1 ADD
    Print #intFileNumber, "設置駅：" & Trim(pfGetEkiNameInfo(NotEkiVer))
    'チェックされたコーナ数分でループ
    For k = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        intCount = udtJprPrintSetteingInfo.iCorner(k) - 1  '画面で指定されたコーナ-1
        If pfCornerGokiCheck(udtJprPrintSetteingInfo.iCorner(k)) = True Then
            
            ' そのコーナは設置されているのでジャーナル出力へ
            '1コーナ目だけ設置駅と設置コーナの間は空行がない
            If k <> 0 Then
                Print #intFileNumber, ""
            End If
            Print #intFileNumber, "設置コーナ：" & gstrCornerName(intCount)

            '////////////////////////////////
            '// 各設定を出力
            '////////////////////////////////
            strNowType = ""
            strNowShoNo = ""
            strNowKubn = ""
            
            For i = 0 To UBound(ReadFileSettei) - 1
            
                If strNowType <> ReadFileSettei(i).sType Then
                    '新しい大分類区分になったのでタイトルを印字
                    Print #intFileNumber, ""
                    Select Case ReadFileSettei(i).sType
                        Case "1"
                            'Print #intFileNumber, "【駅情報】"     'EG30 V32.1.0.1 DEL
                            Print #intFileNumber, "　【駅情報】"    'EG30 V32.1.0.1 ADD
                        Case "2"
                            'Print #intFileNumber, "【監視】"   'EG30 V32.1.0.1 DEL
                            Print #intFileNumber, "　【監視】"  'EG30 V32.1.0.1 ADD
                        Case "3"
                            'Print #intFileNumber, "【ネットワーク】"   'EG30 V32.1.0.1 DEL
                            Print #intFileNumber, "　【ネットワーク】"  'EG30 V32.1.0.1 ADD
                        Case "7"
                            'Print #intFileNumber, "【画面】"   'EG30 V32.1.0.1 DEL
                            Print #intFileNumber, "　【画面】"  'EG30 V32.1.0.1 ADD
                    End Select
                    strNowType = ReadFileSettei(i).sType
                End If
                
                '項目番号が前回と同じ場合は出力しない
                'If strNowShoNo <> ReadFileSettei(i).sNo Then
                If strNowShoNo <> ReadFileSettei(i).sNo Or strNowKubn <> ReadFileSettei(i).sKubun Then
                    '項目名+区分+設定値を出す
                    If (CInt(ReadFileSettei(i).sCorner) = intCount + 1) Or (CInt(ReadFileSettei(i).sCorner) = 0) Then
                        
                        'EG30 V32.1.0.1 ADD START
                        '変更前データ保存された設定値と比較する
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
                        '//下記の項目は駅都度データとジャーナルの出力データが異なる形式になるので強制的に変換する
                        '/////////////////////////////////////////
                        
                        '大分類：１ 中分類：０ 小分類：１８「分類」の値は「9 9 9 9 9 9」形式 半角スペース2文字→1文字に変更
                        If (ReadFileSettei(i).sType = "1") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "18") Then
                           ReadFileSettei(i).sSettei = Replace(ReadFileSettei(i).sSettei, "  ", " ")
                        End If
                        '大分類：２ 中分類：０ 小分類：１「コーナ番号（対ＩＤサーバ)」の値は「ＩＤ」は全角
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "1") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "ＩＤ")
                        End If
                        '大分類：２ 中分類：０ 小分類：２「（対ＩＤサーバ)」の値は「ＩＤ」は全角
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "2") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "ＩＤ")
                        End If
                        '大分類：２ 中分類：０ 小分類：２「（対ＩＤサーバ)」「ＩＤ」は全角
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "2") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "ＩＤ")
                        End If
                        '大分類：２ 中分類：０ 小分類：３「（対ＩＤサーバ)」「ＩＤ」は全角
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "3") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "ＩＤ")
                        End If
                        '大分類：２ 中分類：０ 小分類：８「（対ＩＤサーバ)」「ＩＤ」は全角
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "8") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "ＩＤ")
                        End If

                        '大分類：２ 中分類：０ 小分類：９「（対ＩＤサーバ)」「ＩＤ」は全角
                        If (ReadFileSettei(i).sType = "2") And _
                           (ReadFileSettei(i).sGoki = "0") And _
                           (ReadFileSettei(i).sNo = "9") Then
                           ReadFileSettei(i).sKoumoku = Replace(ReadFileSettei(i).sKoumoku, "ID", "ＩＤ")
                        End If
                     
                        '大分類：７ 中分類：０ 小分類：２１「保守ユーザ設定メニュー画面 無人モード動作設定釦」
                        '項目名と区分の間にスペースをひとつ多く入れる
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
             '設置されていないコーナなので次のコーナへ
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
    '異常ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JPR_PRINT_OUTPUT_ERR, 0)
    
    '異常終了
    'iResponse = MsgBox("異常終了しました。", vbOKOnly + vbCritical, "駅設定テキスト出力結果")
    JPREdit_EkiInfo = False

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : cmdReturn_Click
'//  機能名称  : 「メニュー画面へ戻る」釦押下
'//  機能概要  : 自画面を消去する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-14   CODED   BY [TCC] N.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub cmdReturn_Click()
    '「ジャーナル印字画面：消去」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JPR_PRINT_GAMEN_END, 0)
    
    Unload Me
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Activate
'//  機能名称  : システム初期化(監視盤)画面(アクティブ時)
'//  機能概要  : 最前面表示を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    pfFormActive (hwnd)
    'メール受信タイマを起動する。
    tmrMail.Enabled = True
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : Form_Deactivate
'//  機能名称  : システム初期化(監視盤)画面(ディアクティブ時)
'//  機能概要  : メール受信用のタイマ停止
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.3.0.1) 2009-03-16   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Deactivate()
   On Error Resume Next
    'メール受信タイマを停止する。
    tmrMail.Enabled = False
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : Form_Load
'//  機能名称  : ジャーナル印字画面(ロード時)
'//  機能概要  : 初期処理を行う。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.0.1.1) 2014-03-25   CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dim i As Integer    'カウンター
   
    On Error Resume Next
    
    '「ジャーナル印字画面：表示」ログ出力
    Call sLogTraceReq(LTYP_NORMAL, L3AN_ETC, JPR_PRINT_GAMEN_START, 0)
    
    ' コーナチェックボックス
    For i = 0 To chkCorner.Count - 1
        chkCorner(i).Value = 1
    Next i

    ' 号機チェックボックス
    For i = 0 To chkGouki.Count - 1
        chkGouki(i).Value = 1
    Next i

    ' データ項目
    For i = 0 To chkJprKind.Count - 1
        chkJprKind(i).Value = 0
    Next i
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9000
    Me.Width = 12000
    
   'メール受信タイマのインターバルを'１秒にセット
    tmrMail.Interval = MN_MAIL_INTERVAL
    tmrMail.Enabled = False
   
   'INIファイルよりアプリ起動タイマ値を取得
   lngMAX_Time = GetPrivateProfileInt(APLCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      APL_INTERVAL, HOSHU_FILE)
   '取得値が0の場合、デフォルト値を設定
   If lngMAX_Time = 0 Then
      lngMAX_Time = APL_INTERVAL
   End If
   
   'INIファイルよりログ起動タイマ値を取得
   lngLogMAX_Time = GetPrivateProfileInt(LOGCHKTIMER_SEC, APLSTATIMER_KEY, _
                                      LOG_INTERVAL, HOSHU_FILE)
   '取得値が0の場合、デフォルト値を設定
   If lngLogMAX_Time = 0 Then
      lngLogMAX_Time = LOG_INTERVAL
   End If
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : tmrMail_Timer
'//  機能名称  : メール受信タイマ、タイムアップ処理
'//  機能概要  : メールを受信する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.1.0.1) 2014-03-25   CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub tmrMail_Timer()

    Dim udtReadMail As ML_KYOTU_INF           'メール受信エリア
    Dim lngLength As Long                    '受信メールバイトサイズ
    Dim lngMlSts  As Long                    '受信メールのステータス
    Dim bRet  As Boolean
    Dim lngDataKind As Long                 '画面出力要求RESのデータ種別
    
    On Error Resume Next

    'メールを受信する。
    lngLength = DssMailRead(plMSlot_MN, udtReadMail)
    If lngLength > 0 Then
   '受信メールがあれば、メールＩＤ毎の処理をする。
        Select Case udtReadMail.udtlHeader.dwId        'メールＩＤ
            Case ML_ID_JPR_PRINT_RES
                '「ジャーナル印刷要求RES受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, JPR_PRINT_RES_RECV, 0)
                lngMlSts = udtReadMail.lngData(0)
                If (udtJprPrintSetteingInfo.iJprKind(iJprIdx) = JPR_KIND.JPR_KIND_TUKA_DATA) Or _
                   (udtJprPrintSetteingInfo.iJprKind(iJprIdx) = JPR_KIND.JPR_KIND_RIYO_KINGAKU) Then
                    
                    '通過データまたは利用金額を出力しているときはジャーナル印字要求RESを受信したら、
                    '集計に画面出力完了通知を送信する。
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
                '「情報要求RES受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GETINFO_RES_RECV, 0)
                ' 処理結果を確認
                lngMlSts = udtReadMail.lngData(1)
                If lngMlSts = 0 Then
                    '編集処理を行う。
                    bRet = JprEdit_EkimuId()
                    If bRet = True Then
                       'ジャーナル印字要求CMDを送信
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
                '「画面出力要求RES受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, GETINFO_RES_RECV, 0)
                '処理結果確認
                lngMlSts = udtReadMail.lngData(1)
                'データ種別
                lngDataKind = udtReadMail.lngData(2)
                If lngMlSts = 0 Then
                    '編集処理を行う
                    bRet = JprEdit_TukaData(lngDataKind)
                    If bRet = True Then
                        'ジャーナル印字要求CMDを送信
                        If lngDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
                            bRet = SendMessageJprPrint(TUKA_TXTFILE, ML_CUT_ARI)
                        ElseIf lngDataKind = Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI Then
                            bRet = SendMessageJprPrint(ICRIYO_TXTFILE, ML_CUT_ARI)
                        Else
                            bRet = False
                        End If
                            
                        If bRet = False Then
                            '編集処理が失敗、画面出力完了通知を異常で送信
                            SendMessageGamenOutComplete (ML_GAMEN_OUT_STS.ML_STS_NG)
                            '異常なので、画面出力完了通知メッセージの送信に失敗しようが処理結果は異常
                            Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                            Exit Sub
                        End If
                    Else
                        '編集処理が失敗、画面出力完了通知を異常で送信
                        SendMessageGamenOutComplete (ML_GAMEN_OUT_STS.ML_STS_NG)
                        '異常なので、画面出力完了通知メッセージの送信に失敗しようが処理結果は異常
                        Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), bRet)
                        Exit Sub
                    End If
                Else
                    'RESが異常のため、終了（画面出力完了通知は送信しない)
                    Call ResultDisp(udtJprPrintSetteingInfo.iJprKind(iJprIdx), False)
                End If
                
            Case ML_ID_PROEND_ORD
                '「プロセス終了指示」を受信した場合、
                '「プロセス終了指示受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, PROCESS_END_ORD_RECV, 0)
                'プロセスの終了処理を行う
                pfAbortProc
            
            Case ML_ID_HOSHU_ACTIVE_REQ
                '「保守画面アクティブ表示要求受信正常」ログ出力
                Call sLogTraceReq(LTYP_NORMAL, L3AN_RECV, HOSHU_ACTIVE_REQ_RECV, 0)
                '保守画面アクティブ要求を受信したら、自画面を前面に表示させる。
                AppActivate frmJprPrint.Caption, False
                pfFormActive (frmJprPrint.hwnd)
                
            Case Else
                '「メールID不正」ログ出力
                Call sLogTraceReq(LTYP_ERROR, L3AN_RECV, MAIL_FUSEI_RECV, 0)
        End Select
    End If
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : SendMessageJprPrint
'//  機能名称  : ジャーナル印字要求メッセージを送信する
'//  機能概要  : 出力プロセスにジャーナル印字要求を送信する
'//
'//              型        名称      意味
'//  引数      : String    strFileName   出力ファイル名
'//              Byte      byCut         0:カットなし   1：カットあり
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function SendMessageJprPrint(strFileName As String, byCut As Byte) As Boolean

    Dim udtMail As MAIL_JPR_PRINT_CMD   'ジャーナル印刷要求メール送信エリア
    Dim lngRet As Long                  '関数戻り値
    Dim lngErrCode As Long              'エラーコード
    Dim bTmpArray() As Byte
    Dim i       As Integer
    On Error Resume Next
    
    
    'ジャーナル印字要求を出力プロセスに送信する。
    udtMail.mlHeader.dwId = ML_ID_JPR_PRINT_REQ
    udtMail.mlHeader.dwSize = MlSize.JPR_PRINT_REQ
    udtMail.mlHeader.dwProid = RHOSHU_ID
    udtMail.mlHeader.dwSubArea = 0
    bTmpArray = StrConv(strFileName, vbFromUnicode)
    For i = 0 To UBound(bTmpArray)
        'udtMail.byOutputFilePath(i) = Chr(bTmpArray(i))
        udtMail.byOutputFilePath(i) = bTmpArray(i)
    Next
    udtMail.dwCut = byCut                                   'カット有無
    udtMail.dwOutputDataPoint = 0                           '出力データポイント
    
    lngRet = DssSendMail(MAIL_SLOT_OUTPUT, Len(udtMail), udtMail.mlHeader)
    If lngRet = False Then
       '「ジャーナル印字画面：ジャーナル印刷要求送信異常」ログ出力
       lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
       Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, JPR_PRINT_REQ_SEND, lngErrCode)
       SendMessageJprPrint = False
       Exit Function
    Else
       '「ジャーナル印字画面：ジャーナル印刷要求送信正常」ログ出力
       Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, JPR_PRINT_REQ_SEND, 0)
       SendMessageJprPrint = True
    End If
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : SendMessageInfoReq
'//  機能名称  : 情報要求CMDメッセージを送信する
'//  機能概要  : ID制に情報要求要求CMDを送信する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function SendMessageInfoReq() As Boolean
    
    Dim bRet As Boolean                 '戻り値
    Dim lngErrCode As Long              'エラーコード
    Dim udtMail As MAIL_INFO_CMD        '画面表示要求
    Dim uMail As ML_KYOTU_INF           'メール
 
   'バッファフラッシュ要求をログプロセスに送信する
   '情報要求CMD(駅務機器ID=0)をID制御に送信する
   udtMail.mlHeader.dwId = ML_ID_INFO_CMD
   udtMail.mlHeader.dwSize = MlSize.INFO_CMD
   udtMail.mlHeader.dwProid = RHOSHU_ID
   udtMail.mlHeader.dwSubArea = 0
   udtMail.dwRequestType = MailCmdType.ML_DT_EKIMU_ID
   iSendType = MailCmdType.ML_DT_EKIMU_ID
   bRet = DssSendMail(MAIL_SLOT_IDSEI, Len(udtMail), udtMail.mlHeader)
   If bRet = False Then
      '「駅務機器ID確認：情報要求CMD送信異常」ログ出力
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
      SendMessageInfoReq = False
      Exit Function
   Else
      '「駅務機器ID確認：情報要求CMD送信正常」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
   End If
   
   SendMessageInfoReq = True

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : SendMessageGamenOutReq
'//  機能名称  : 画面出力要求CMD送信
'//  機能概要  : 集計に画面出力要求CMDを送信する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function SendMessageGamenOutReq(dwDataKind As Long) As Boolean
    
    Dim bRet As Boolean                     '戻り値
    Dim lngErrCode As Long                  'エラーコード
    Dim udtMail As MAIL_GAMEN_OUTPUT_CMD    '画面出力要求
    Dim uMail As ML_KYOTU_INF               'メール
 
   'バッファフラッシュ要求をログプロセスに送信する
   '画面出力要求CMDを集計に送信する
   udtMail.mlHeader.dwId = ML_ID_GAMEN_OUTPUT_REQ
   udtMail.mlHeader.dwSize = MlSize.GAMEN_OUT_REQ
   udtMail.mlHeader.dwProid = RHOSHU_ID
   udtMail.mlHeader.dwSubArea = 0
   udtMail.dwSeqence = 0                ' シーケンス番号0固定
   udtMail.dwDataKind = dwDataKind
   bRet = DssSendMail(MAIL_SLOT_SHUKEI, Len(udtMail), udtMail.mlHeader)
   If bRet = False Then
      '「画面出力要求CMD送信異常」ログ出力
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
      SendMessageGamenOutReq = False
      Exit Function
   Else
      '「画面出力要求CMD送信正常」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
   End If

   SendMessageGamenOutReq = True

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : SendMessageGamenOutComplete
'//  機能名称  : 画面出力要求完了通知送信
'//  機能概要  : 集計に画面出力完了通知を送信する
'//
'//              型        名称      意味
'//  引数      : Long     dwStatus   メッセージにセットするステータス
'//
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function SendMessageGamenOutComplete(dwStatus As Long) As Boolean
    
    Dim bRet As Boolean                     '戻り値
    Dim lngErrCode As Long                  'エラーコード
    Dim udtMail As MAIL_GAMEN_OUTPUT_COMP   '画面出力要求完了通知
    Dim uMail As ML_KYOTU_INF               'メール
 
   'バッファフラッシュ要求をログプロセスに送信する
   '画面出力要求CMDを集計に送信する
   udtMail.mlHeader.dwId = ML_ID_GAMEN_OUTPUT_COMP
   udtMail.mlHeader.dwSize = MlSize.GAMEN_OUT_COMP
   udtMail.mlHeader.dwProid = RHOSHU_ID
   udtMail.mlHeader.dwSubArea = 0
   udtMail.dwSequence = 0                ' シーケンス番号0固定
   udtMail.dwStatus = dwStatus
   bRet = DssSendMail(MAIL_SLOT_SHUKEI, Len(udtMail), udtMail.mlHeader)
   If bRet = False Then
      '「画面出力要求完了通知送信異常」ログ出力
      lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_MAIL_IO + ECOD_MSEND
      Call sLogTraceReq(LTYP_ERROR, L3AN_SEND, GETINFO_CMD_SEND, lngErrCode)
      SendMessageGamenOutComplete = False
      Exit Function
   Else
      '「画面出力要求完了通知送信正常」ログ出力
      Call sLogTraceReq(LTYP_NORMAL, L3AN_SEND, GETINFO_CMD_SEND, 0)
   End If
   
   SendMessageGamenOutComplete = True

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : ResultDisp
'//  機能名称  : ジャーナル印刷結果表示
'//  機能概要  : ジャーナルの印刷結果を表示する。
'//
'//              型        名称      意味
'//  引数      : Integer    iJprKind    ジャーナル種別
'//              Boolean    bResult     結果(true/false)
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(30.3.0.1) 2014-10-01  REVISED BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi07_003_01】
'//     REVISIONS :(32.1.0.1) 2016-06-10  REVISED BY [TCC] T.Nakajima
'//                 2016年度施策対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub ResultDisp(iJprKind As Integer, bResult As Boolean)
    Dim strStatus   As String
    Dim strJprName  As String

    '処理結果文言作成
    Select Case iJprKind
        Case JPR_KIND.JPR_KIND_EKI_INFO
            strJprName = "駅都度データ確認（駅情報）"
            
        Case JPR_KIND.JPR_KIND_JIKAI_INFO
            strJprName = "駅都度データ確認（自改）"
            
        Case JPR_KIND.JPR_KIND_SETTING_LST
            strJprName = "設定値一覧"
            
        Case JPR_KIND.JPR_KIND_TUKA_DATA
            strJprName = "通過データ"
            
        Case JPR_KIND.JPR_KIND_RIYO_KINGAKU
            strJprName = "利用金額データ"
            
        Case JPR_KIND.JPR_KIND_KADO_VER
            strJprName = "稼働バージョン一覧"
            
        Case JPR_KIND.JPR_KIND_SIMEKIRI
            strJprName = "締切オフライン出力"
            
        Case JPR_KIND.JPR_KIND_EKIMU_ID
            strJprName = "駅務機器ＩＤ"
        'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
        Case JPR_KIND.JPR_KIND_SUBGATE_INFO
            strJprName = "駅都度データ確認（ｴﾝｺｰﾄﾞｺｰﾅ号機情報定義）"
        'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END
        'EG30 V32.1.0.1 ADD START
        Case JPR_KIND.JPR_KIND_GATE_CFG
            strJprName = "改札機保守設定データ"
        'EG30 V32.1.0.1 ADD END
    End Select
    
    If bResult = True Then
        '正常
        LstStatus.AddItem strJprName & "    " & "正常終了しました"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        'Call JPRScreenEnable(True)
    Else
        '異常
        LstStatus.AddItem strJprName & "    " & "異常終了しました"
        LstStatus.Selected(LstStatus.ListCount - 1) = True
        'Call JPRScreenEnable(True)
    End If

    iJprIdx = iJprIdx + 1
    If iJprIdx < udtJprPrintSetteingInfo.iJprCount Then
        '2種類目以降のジャーナル出力
        JprOutputProc
    Else
        '全ジャーナル出力完了ならば、釦、チェックボックス操作可能にする。
        Call JPRScreenEnable(True)
        iJprIdx = 0
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : JPRScreenEnable
'//  機能名称  : ジャーナル印刷画面の設定変更可否制御
'//  機能概要  : ジャーナル印字画面の内容を変更の可否を制御する
'//
'//              型        名称      意味
'//  引数      : Boolean   bEnable    true:変更可能  false:変更不可
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG30 V32.2.0.1) 2016-07-20  CODED   BY [TCC] T.Nakajima
'//                 2016年度施策対応 プログレスバー非表示対応
'//                 ロール紙半分以上を使用するジャーナルがあるため、プログレスバー表示中に間に合わなくなるため
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub JPRScreenEnable(bEnable As Boolean)
    Dim i   As Integer
    
    ' コーナチェックボックス
    For i = 0 To chkCorner.Count - 1
        chkCorner(i).Enabled = bEnable
    Next i
    
    ' 号機チェックボックス
    For i = 0 To chkGouki.Count - 1
        chkGouki(i).Enabled = bEnable
    Next i
 
    ' データ項目
    For i = 0 To chkJprKind.Count - 1
        chkJprKind(i).Enabled = bEnable
    Next i
    
    '印字ボタン
    cmdPrint.Enabled = bEnable
    
    '戻るボタン
    cmdReturn.Enabled = bEnable
    
    If bEnable = False Then
        'ステータス表示部をクリアする（印字ボタン押下で前回の処理結果をクリア)
         LstStatus.Clear
        'プログレスバーを表示する
        ' EG30 V32.2.0.1 DEL START
        'Call SendMessageProgress(ML_ID_PRGBAR_SHOW_REQ, PRG_JPR_OUT)
        ' EG30 V32.2.0.1 DEL END
    Else
        'プログレスバーを消去する
        ' EG30 V32.2.0.1 DEL START
        'Call SendMessageProgress(ML_ID_PRGBAR_HIDE_REQ, PRG_JPR_OUT)
        ' EG30 V32.2.0.1 DEL END
    End If

End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : GetPrintSettings
'//  機能名称  : 画面のチェック状態を取得
'//  機能概要  : 指定されたコーナ数を取得する
'//
'//              型        名称      意味
'//  引数      : Boolean   bEnable    true:変更可能  false:変更不可
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub GetPrintSettings()
    Dim i               As Integer
    Dim k               As Integer
    Dim iCornerCount    As Integer
    Dim iGoukiCount     As Integer
    Dim iJprCount       As Integer
    
    ' ジャーナル設定情報をクリアする
    udtJprPrintSetteingInfo = udtInitJprSetting
    
    
    ' コーナのチェック状態
    k = 0
    For i = 0 To chkCorner.Count - 1
        If chkCorner(i).Value = 1 Then
            iCornerCount = iCornerCount + 1
            udtJprPrintSetteingInfo.iCorner(k) = i + 1
            k = k + 1
        End If
    Next i
    udtJprPrintSetteingInfo.iCornerCount = iCornerCount
    
    '号機のチェック状態
    k = 0
    For i = 0 To chkGouki.Count - 1
        If chkGouki(i).Value = 1 Then
            iGoukiCount = iGoukiCount + 1
            udtJprPrintSetteingInfo.iGouki(k) = i + 1
            k = k + 1
        End If
    Next i
    udtJprPrintSetteingInfo.iGoukiCount = iGoukiCount
    
    'ジャーナル種別のチェック状態
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
'//  関数名称  : JprOutputProc
'//  機能名称  : ジャーナル出力処理
'//  機能概要  : 出力ファイル作成と出力プロセスに要求を送信
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-17   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(30.3.0.1) 2014-10-01  REVISED BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi07_003_01】
'//     REVISIONS :(32.1.0.1) 2016-06-10  REVISED BY [TCC] T.Nakajima
'//                 2016年度施策対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub JprOutputProc()
    Dim bRet        As Boolean
    
    'EG30 V32.1.0.1 ADD START
    Dim i, j            As Integer  'コーナ、号機カウンタ
    Dim intComSts       As Integer  '通信状態
    Dim blnSkipFlg      As Boolean  '保守設定データなし
    Dim intGateNo       As Integer  '号機番号（1～32）
    'EG30 V32.1.0.1 ADD END
    Select Case udtJprPrintSetteingInfo.iJprKind(iJprIdx)
        Case JPR_KIND.JPR_KIND_EKI_INFO          ' 駅都度データ確認(駅情報)
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
            
        Case JPR_KIND.JPR_KIND_JIKAI_INFO            ' 駅都度データ確認(自改)
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

        Case JPR_KIND.JPR_KIND_SETTING_LST           ' 設定値一覧
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
        
        Case JPR_KIND.JPR_KIND_TUKA_DATA             ' 通過データ
            ' メッセージを送信して編集元ファイルの作成を依頼するので、ここでは編集処理は呼ばない。
            ' 編集処理を呼ぶのはRESメールを受信したとき。ジャーナル印字要求は編集処理が終わってから呼ぶ。
            ' 指定されたコーナが未設置ならば処理はしない
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
        
        Case JPR_KIND.JPR_KIND_RIYO_KINGAKU          ' 利用金額データ
            ' メッセージを送信して編集元ファイルの作成を依頼するので、ここでは編集処理は呼ばない。
            ' 編集処理を呼ぶのはRESメールを受信したとき。ジャーナル印字要求は編集処理が終わってから呼ぶ。
            ' 指定されたコーナが未設置ならば処理はしない
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
        
        Case JPR_KIND.JPR_KIND_KADO_VER              ' 稼動バージョン一覧
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
        
        Case JPR_KIND.JPR_KIND_SIMEKIRI              ' 締切オフライン出力
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
        
        Case JPR_KIND.JPR_KIND_EKIMU_ID              ' 駅務機器ID
            ' メッセージを送信して編集元ファイルの作成を依頼するので、ここでは編集処理は呼ばない。
            ' 編集処理を呼ぶのはRESメールを受信したとき。ジャーナル印字要求は編集処理が終わってから呼ぶ。
            bRet = SendMessageInfoReq
            If bRet = False Then
                Call ResultDisp(JPR_KIND.JPR_KIND_EKIMU_ID, bRet)
                Exit Sub
            End If
        'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
        Case JPR_KIND.JPR_KIND_SUBGATE_INFO         ' 駅都度データ確認(エンコードコーナ号機情報定義)
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
        'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END
        'EG30 V32.1.0.1 ADD START
        Case JPR_KIND.JPR_KIND_GATE_CFG             ' 改札機保守設定データ
            'チェックされている改札機の通信状態を取得する
            For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
                For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
                    'そのコーナ、号機は設置されているか？
                    If pfCornerGokiToGateNo(udtJprPrintSetteingInfo.iCorner(i), udtJprPrintSetteingInfo.iGouki(j), intGateNo) = True Then
                    
                        '監視盤起動有無チェック
                        If CheckAppStart(PROC_KANRI) <> 0 Then
                            gpfGetjikaiConectSts intComSts, intGateNo
                            If intComSts <> CONECTSTS_NORMAL Then
                                Exit For
                            End If
                        End If
                    End If
                Next j
                '1台でも通信異常の改札機があれば、警告を表示するので、コーナ単位のループを抜ける
                If intComSts <> CONECTSTS_NORMAL Then
                    Exit For
                End If
            Next i
            
            'ステータス表示部に通信異常改札機があることを表示する
            If intComSts <> CONECTSTS_NORMAL Then
                LstStatus.AddItem "選択したコーナに通信異常の改札機があります"
                LstStatus.AddItem "通信異常号機の改札機保守設定データは最新で無い可能性があります"
                LstStatus.Selected(LstStatus.ListCount - 1) = True
            End If
            
            bRet = JprEdit_GateCfg(blnSkipFlg)
            '改札機保守設定データ未受信の改札機があったため、ジャーナル印字できなかったことを表示する。
            If blnSkipFlg = True Then
                LstStatus.AddItem "改札機保守設定データを印字できなかった改札機があります"
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
'//  関数名称  : JPREdit_JikaiInfo
'//  機能名称  : 「印字」釦押下時処理
'//  機能概要  : 現在駅設定ファイル（自改)をテキスト表示する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-19   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(30.1.0.1) 2014-05-01  CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(30.3.0.1) 2014-10-01  CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi07_003_01】
'//                 ・改札機設置条件の印字は別ジャーナルへ独立させる
'//     REVISIONS :(32.1.0.1) 2016-06-16  CODED   BY [TCC] T.Nakajima
'//                 2016年度施策対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function JPREdit_JikaiInfo() As Boolean

    Dim strFileName             As String                   'ファイル名
    Dim bRet                    As Boolean                  '関数戻り値
    Dim lErrCode                As Long                     'エラーコード
    Dim strLineCount            As String                   '行数カウンタ
    
    Dim sWriteDir               As String                   '書き込み先フォルダ名
    Dim intFileNumber           As Integer                  'ファイルポインタ
    
    Dim ReadFileSettei()        As JIKAIINFO_IMAGE_FILE     '自改情報イメージファイル
    Dim i                       As Integer
    Dim j                       As Integer
    Dim iCornerIdx              As Integer                  'コーナインデックス(何番目のコーナ)
    
    Dim fso                     As New FileSystemObject     'ファイルシステムオブジェクト
    Dim FsoTS                   As TextStream

    Dim strNowShobunrui         As String                   '現在編集中の小分類コード
    Dim strNowKubun             As String                   '現在編集中の区分
    Dim strNowCorner            As String                   '現在編集中のコーナ
    
    'EG30 V32.1.0.1 ADD START
    Dim strEkiSettiBefPath      As String           '現在駅設定データ（変更前保存）
    Dim strGetValue             As String * 64      'DLLによって設定されるため、64固定長にしている
    Dim strCompValue            As String           '設定値（変更前保存）
    Dim strChangeFlg            As String           '変更印
    Dim intValueLen             As Integer          '取得した設定値の長さ
    Dim intGateNo               As Integer          '1～32号機
    'EG30 V32.1.0.1 ADD END
    
    'エラールーチンを宣言
    On Error GoTo OUTPUT_ERROR
    '画面で指定されたコーナは設置されているか？
    If pfSettingCheck(True) = False Then
        'すべて未設置のコーナ、号機なのでエラーとする
        GoTo OUTPUT_ERROR
    End If
    
    'イメージファイルの出力先
    sWriteDir = EKI_JPR_GATE_TXTFILE

    '駅都度データ確認（自改）イメージファイル作成
    bRet = dllGetEkiIniDataJpr(1, EKI_TUDO_CHK_GATE_FILE_JPR, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        '駅都度データ確認（自改）イメージファイル削除
        Kill EKI_TUDO_CHK_GATE_FILE_JPR
        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)
        JPREdit_JikaiInfo = False
        Exit Function
    End If
    
    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL START
'    'EG20 V30.1.0.1 ADD START
'    '自改補助CSVファイル作成
'    bRet = dllGetEkiIniData(2, EKI_TUDO_CHK_SUBGATE_FILE, EKI_SETTI_FILE, lErrCode)
'    If bRet = False Then
'        '自改補助CSVファイル削除
'        Kill EKI_TUDO_CHK_SUBGATE_FILE
'        '異常ログ出力
'        Call pfOutPutErrLog(lErrCode)
'        JPREdit_JikaiInfo = False
'        Exit Function
'    End If
'    'EG20 V30.1.0.1 ADD END
    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
    
    
    ' コーナ名称設定処理
    Call gsGetCornerName

    '初期値設定
    strFileName = ""

    '----------------------------------------------------
    '現在駅設定ファイル検索
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    'ファイルが存在しない場合
    If strFileName = "" Then
    
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        '異常終了
        JPREdit_JikaiInfo = False
        Exit Function
        
    End If

    '駅都度データ(自改)イメージファイルの件数を取得
    'ファイル番号取得
    intFileNumber = FreeFile
    Open EKI_TUDO_CHK_GATE_FILE_JPR For Input As #intFileNumber
    
    'CSVファイル行数カウント（ファイル終端までループを繰り返す）
        Do While Not EOF(intFileNumber)                     ' EG20 V3.3.0.1追加
            Line Input #intFileNumber, strLineCount
            j = j + 1
        Loop
    
    'CSVファイルクローズ
    Close #intFileNumber

    'ファイル番号取得
    intFileNumber = FreeFile

    '再設定
    ReDim ReadFileSettei(j) As JIKAIINFO_IMAGE_FILE        'ファイル読込用エリア
        
    'CSVファイルオープン
    Open EKI_TUDO_CHK_GATE_FILE_JPR For Input As #intFileNumber

    'リスト表示分読み込み（ファイル終端までループを繰り返す）
        For i = 0 To j - 1
            Input #intFileNumber, ReadFileSettei(i).strBunrui_Dai, ReadFileSettei(i).strBunrui_Tyu, _
            ReadFileSettei(i).srtBunrui_Sho, ReadFileSettei(i).strCorner, ReadFileSettei(i).strKomoku, _
            ReadFileSettei(i).strKubun, ReadFileSettei(i).strData, ReadFileSettei(i).strSetShosai
        Next

    'CSVファイルクローズ
    Close #intFileNumber
    
    'EG30 V32.1.0.1 ADD START
    'そのコーナの変更前データ保存されたデータをメモリ上に展開する(コーナ0）
    strEkiSettiBefPath = Replace(EKI_SETTI_FILE_BEF, "#", "0")
    Call dllGetEkiIniDataBefore(strEkiSettiBefPath, lErrCode)
    'EG30 V32.1.0.1 ADD END
    
    '///////////////////////////////////////
    '// ジャーナル出力イメージファイルを作成
    '///////////////////////////////////////
    '未使用のファイル番号取得
    intFileNumber = FreeFile
    
    'ジャーナル出力イメージファイルを作成
    Open sWriteDir For Output As #intFileNumber
    
    'タイトル表示
    'PrintHeader intFileNumber, "駅都度データ確認"  'EG30 V32.1.0.1 DEL
    PrintHeader3 intFileNumber, "駅都度データ確認", pfGetSaveDate(0)
    Print #intFileNumber, "設置駅：" & Trim(pfGetEkiNameInfo(NotEkiVer))
    
    strNowShobunrui = ""
    strNowKubun = ""
    strNowCorner = ""
    
    For i = 0 To UBound(ReadFileSettei) - 1
        'コーナが切り替わったか？
        If (ReadFileSettei(i).strCorner <> strNowCorner) Then
            'iCornerIdx = iCornerIdx + 1
            'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL START
'            'EG20 V30.1.0.1 ADD START
'            '複数コーナ出力で2コーナ目以降がある場合、2コーナ目に入る前に自改補助を出力
'            If strNowCorner <> "" Then
'                pfOutPutSubGate CInt(strNowCorner), intFileNumber
'            End If
'            'EG20 V30.1.0.1 ADD END
            'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
            '設置コーナを出力
            If IsTaisyoCorner(CInt(ReadFileSettei(i).strCorner)) = True Then
                
                '対象コーナであっても対象号機がないかもしれない
                For j = 0 To 15
                    If IsTaisyoGoki(CInt(ReadFileSettei(i).strCorner), j + 1) = True Then
                        If i <> 0 Then
                            Print #intFileNumber, ""
                        End If
                        Print #intFileNumber, "設置コーナ：" & gstrCornerName(CInt(ReadFileSettei(i).strCorner) - 1)
                        Exit For
                    End If
                Next j
            End If
            strNowCorner = ReadFileSettei(i).strCorner
        End If
    
        'その号機は出力対象か？
        If IsTaisyoGoki(CInt(ReadFileSettei(i).strCorner), CInt(ReadFileSettei(i).strBunrui_Tyu)) = True Then
            '小分類と区分が一致しなければタイトルを出力する
            If (ReadFileSettei(i).srtBunrui_Sho <> strNowShobunrui) Or (ReadFileSettei(i).strKubun <> strNowKubun) Then
                'タイトルを出力
                Print #intFileNumber, ""
                'Print #intFileNumber, "【" & ReadFileSettei(i).strKomoku & "】" & ReadFileSettei(i).strKubun   'EG30 V32.1.0.1 DEL
                Print #intFileNumber, "　【" & ReadFileSettei(i).strKomoku & "】" & ReadFileSettei(i).strKubun  'EG30 V32.1.0.1 ADD
                strNowShobunrui = ReadFileSettei(i).srtBunrui_Sho
                strNowKubun = ReadFileSettei(i).strKubun
            End If
            '各号機の設定を出力
            
            'EG30 V32.1.0.1 ADD START
            'ジャーナル編集データファイルの中分類はコーナ号機番号がセットされ、さらにコーナ番号もセットされているが、
            '比較相手となるEKI_SETTI.CSVは中分類は１～３２でコーナ番号は０となっているため、コーナ号機番号を１～３２に変換する
            If pfCornerGokiToGateNo(CInt(ReadFileSettei(i).strCorner), CInt(ReadFileSettei(i).strBunrui_Tyu), intGateNo) = True Then
            
                '変更前データ保存された設定値と比較する(大分類が改札機の場合は、コーナは0固定で検索）
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
            '比較相手の号機がいなかったら「★」
            Else
                strChangeFlg = DIFF_MARK_STRING_ON
            End If
            'EG30 V32.1.0.1 ADD END
            
            '号機番号を表示する項目は号機番号を99形式に変換する
            If (ReadFileSettei(i).strBunrui_Dai = 4 And ReadFileSettei(i).srtBunrui_Sho = 5) Or _
               (ReadFileSettei(i).strBunrui_Dai = 4 And ReadFileSettei(i).srtBunrui_Sho = 7) Then
                ReadFileSettei(i).strData = Format(ReadFileSettei(i).strData, "0#")
            End If
                
            'Print #intFileNumber, ReadFileSettei(i).strBunrui_Tyu & "号機 " & ReadFileSettei(i).strData    'EG30 V32.1.0.1 DEL
            Print #intFileNumber, strChangeFlg & ReadFileSettei(i).strBunrui_Tyu & "号機 " & ReadFileSettei(i).strData    'EG30 V32.1.0.1 ADD
            'bJprFlg = True
        End If
    Next i
    Print #intFileNumber, ""
    
    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL START
'    'EG20 V30.1.0.1 ADD START
'    '自改補助を出力
'    pfOutPutSubGate CInt(strNowCorner), intFileNumber
'    Print #intFileNumber, ""
'    'EG20 V30.1.0.1 ADD END
    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
    
    Print #intFileNumber, FOOTER_STRING
    'ファイルをクローズする。
    Close #intFileNumber
    Set fso = Nothing
    JPREdit_JikaiInfo = True
    Exit Function

OUTPUT_ERROR:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    Set fso = Nothing
    '異常ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JPR_PRINT_OUTPUT_ERR, 0)
    
    JPREdit_JikaiInfo = False
End Function
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : JPREdit_SubGateInfo
'//  機能名称  : 「印字」釦押下時処理
'//  機能概要  : 現在駅設定ファイル（エンコードコーナ号機情報定義)をテキスト表示する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(30.3.0.1) 2014-10-01  CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi07_003_01】
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function JPREdit_SubGateInfo() As Boolean

    Dim strFileName             As String                   'ファイル名
    Dim bRet                    As Boolean                  '関数戻り値
    Dim lErrCode                As Long                     'エラーコード
    Dim strLineCount            As String                   '行数カウンタ
    
    Dim sWriteDir               As String                   '書き込み先フォルダ名
    Dim intFileNumber           As Integer                  'ファイルポインタ
    
    Dim ReadFileSettei()        As JIKAIINFO_IMAGE_FILE     '自改情報イメージファイル
    Dim i                       As Integer
    Dim j                       As Integer
    Dim iCornerIdx              As Integer                  'コーナインデックス(何番目のコーナ)
    
    Dim fso                     As New FileSystemObject     'ファイルシステムオブジェクト
    Dim FsoTS                   As TextStream

    Dim strNowShobunrui         As String                   '現在編集中の小分類コード
    Dim strNowKubun             As String                   '現在編集中の区分
    Dim strNowCorner            As String                   '現在編集中のコーナ
    
    'エラールーチンを宣言
    On Error GoTo OUTPUT_ERROR
    
    'イメージファイルの出力先
    sWriteDir = EKI_JPR_SUBGATE_TXTFILE

    '自改補助CSVファイル作成
    bRet = dllGetEkiIniData(2, EKI_TUDO_CHK_SUBGATE_FILE, EKI_SETTI_FILE, lErrCode)
    If bRet = False Then
        '自改補助CSVファイル削除
        Kill EKI_TUDO_CHK_SUBGATE_FILE
        '異常ログ出力
        Call pfOutPutErrLog(lErrCode)
        JPREdit_SubGateInfo = False
        Exit Function
    End If
    
    '初期値設定
    strFileName = ""

    '----------------------------------------------------
    '現在駅設定ファイル検索
    '----------------------------------------------------
    strFileName = Dir(EKI_SETTI_FILE)

    'ファイルが存在しない場合
    If strFileName = "" Then
    
        '異常ログ出力
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, NOT_FILE_EKI_SETTI, 0)
        
        '異常終了
        JPREdit_SubGateInfo = False
        Exit Function
        
    End If

    '///////////////////////////////////////
    '// ジャーナル出力イメージファイルを作成
    '///////////////////////////////////////
    '未使用のファイル番号取得
    intFileNumber = FreeFile
    
    'ジャーナル出力イメージファイルを作成
    Open sWriteDir For Output As #intFileNumber
    
    'タイトル表示
    PrintHeader2 intFileNumber, "駅都度データ確認", "(エンコードコーナ号機情報定義)"
    Print #intFileNumber, "設置駅：" & Trim(pfGetEkiNameInfo(NotEkiVer))
    
    strNowShobunrui = ""
    strNowKubun = ""
    strNowCorner = ""
    
    If pfOutPutSubGate(0, intFileNumber) = False Then
        GoTo OUTPUT_ERROR
    End If
    Print #intFileNumber, ""
    
    Print #intFileNumber, FOOTER_STRING
    'ファイルをクローズする。
    Close #intFileNumber
    Set fso = Nothing
    JPREdit_SubGateInfo = True
    Exit Function

OUTPUT_ERROR:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    Set fso = Nothing
    '異常ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JPR_PRINT_OUTPUT_ERR, 0)
    
    JPREdit_SubGateInfo = False
End Function
'EG30 V32.1.0.1 ADD START
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  関数名称  : JprEdit_GateCfg
'//  機能名称  : 改札機保守設定データジャーナルイメージファイル作成
'//  機能概要  : 改札機保守設定データジャーナルイメージファイルを作成する
'//
'//              型        名称      意味
'//  引数      : Boolean   bSkipFlg  改札機保守データが無いため、ジャーナル編集をスキップした号機がある。
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-10   CODED   BY [TCC] T.Nakajima
'//             2016年度施策対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_GateCfg(ByRef bSkipFlg As Boolean) As Boolean

    Dim strOutputFile As String         '出力ファイル
    Dim lngRet As Long                  '関数返り値
    Dim lngErrCode As Long              'エラーコード
    Dim iOutFile    As Integer          'ファイル番号
    Dim ReadFileGateCfg()    As GATE_CFG_DATA_FILE  '改札機保守設定データ
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim l           As Integer
    Dim strJpCfgPath            As String  '号機別設定コンフィグファイル(JP用)
    Dim strSetteiBefFolder      As String  '統合監視盤変更前保存領域
    Dim strJpCfgPathBef         As String  '変更前保存した号機別設定コンフィグファイル名
    Dim strDispImageFileName    As String  '編集データファイル名
    Dim objFs                   As New FileSystemObject
    Dim intFileNo               As Integer
    
    Dim blnRet                  As Boolean  '編集データ作成関数戻り値
    
    Dim strMutexName    As String       'ミューテックス名
    
    Dim strNowInfoName  As String       '現在出力中の情報部名
    Dim strNowDai       As String       '現在出力中の大項目
    Dim strNowChu       As String       '現在出力中の中項目
    
    Dim iKoumokuByte          As Integer '項目名のバイト数
    Dim iValueByte          As Integer '設定値のバイト数
    Dim iSpaceByte          As Integer '中間に挿入するスペースのバイト数
    Dim strChangeFlg        As String  '変更フラグ
    Dim strSyoName          As String  '小項目名
    Dim strValue            As String  '設定値
    Dim blnInfoNameFlg      As Boolean '改行フラグ（情報部名直後の大項目－中項目名の直前は改行なし)
    Dim intOutCount         As Integer  '出力可能号機数
    Dim intOutCountbyCorner(0 To 5) As Integer '出力可能号機数（コーナ毎）
    Dim intGateNo           As Integer  '1～32号機
    Dim bResult(0 To 5, 0 To 15) As Boolean  'ジャーナル編集データファイル出力結果
    Dim strExistsCheckFileName As String     '0101.CSV～0616.CSVまでのファイルの存在をチェック
    Dim bDelFlg                 As Boolean  '削除フラグ
    Const COLON_LEN = 2                 '「：」のバイト数
    
    On Error GoTo Err_handler
    bDelFlg = False
    intOutCount = 0
    
    '設置駅
    gsGetStationName
    '自改情報
    gsGetGateInfo
    'コーナ名
    gsGetCornerName
    'コーナタイプ
    gsGetCornerType
    
    '画面で指定されたコーナは設置されているか？
    If pfSettingCheck(True) = False Then
        'すべて未設置のコーナ、号機なのでエラーとする
        GoTo Err_handler
    End If
    
    '出力ファイル名編集
    strOutputFile = GATE_CFG_TXTFILE
    
    'ジャーナル編集データファイルをすべて削除する
    '削除ファイルが存在しない場合はErr_Handlerにいってしまうため、存在チェックを行う。
    'ファイルが一つでも見つかれば、ワイルドカードによるファイル削除ができるので、ループを抜ける
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
        'JP_CFG*.CSVで削除済み
        If bDelFlg = True Then
            Exit For
        End If
    Next i
    
    bSkipFlg = False
    
    'ファイル出力関数をCall
    'チェックされているコーナ、号機分について処理
    For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        intOutCountbyCorner(i) = 0
        For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
            'そのコーナ、号機は設置されているか？
            If pfCornerGokiToGateNo(udtJprPrintSetteingInfo.iCorner(i), udtJprPrintSetteingInfo.iGouki(j), intGateNo) = True Then
                
                strDispImageFileName = Replace(EDIT_DATA_GATECFG, "####", _
                    Format(udtJprPrintSetteingInfo.iCorner(i), "0#") & Format(udtJprPrintSetteingInfo.iGouki(j), "0#")) & ".csv"
                
                'ミューテックス名を作成
                strMutexName = Replace(MU_N_CFG, "##", Format(intGateNo, "0#"))

                strJpCfgPath = PATH_DATA & Replace(JP_CFG, "##", Format(intGateNo, "0#"))
                strSetteiBefFolder = PATH_OPERATE & "CORNER" & udtJprPrintSetteingInfo.iCorner(i) & "\\SETTEI_BEF\\"
                strJpCfgPathBef = strSetteiBefFolder & Replace(JP_CFG, "##", Format(intGateNo, "0#"))
                
                '元データ(JP_CFGnn.GAT)が存在した場合は、ジャーナルデータファイルを作成
                If objFs.FileExists(strJpCfgPath) = True Then
                    bResult(i, j) = dllCreateGateCfgData(gintCornerType(udtJprPrintSetteingInfo.iCorner(i) - 1), _
                                                strDispImageFileName, strJpCfgPath, strJpCfgPathBef, strMutexName, lngErrCode)
                    If bResult(i, j) <> False Then
                        'ジャーナル編集データファイルが一つ以上作れていれば、他の号機で失敗しても印刷可能のため。
                        intOutCount = intOutCount + 1
                        intOutCountbyCorner(i) = intOutCountbyCorner(i) + 1
                    Else
                        'テキスト作成処理失敗によりスキップ
                        bSkipFlg = True
                    End If
                Else
                    ' 「改札機保守設定データを印字できなかった改札機があります」を表示するためにON
                    bSkipFlg = True
                End If
                
            End If
        Next j
    Next i
    
    'ジャーナル編集データファイルが作れていれば、ジャーナル出力可能
    If intOutCount > 0 Then
        '改札機保守設定データ ジャーナルイメージファイルを作成
        iOutFile = FreeFile
        Open strOutputFile For Output As #iOutFile
        
        'ヘッダー部
        PrintHeader iOutFile, "改札機保守設定データ確認"
        
        '設置駅
        Print #iOutFile, "設置駅：" & gstrStationName(0)
        
        'チェックされたコーナ数分ループ
        For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
            Erase ReadFileGateCfg
            
            'そのコーナの改札機がすべて改札機保守設定データを持っていない場合は印字しない
            If intOutCountbyCorner(i) > 0 Then
                'コーナ名
                Print #iOutFile, "設置コーナ：" & gstrCornerName(udtJprPrintSetteingInfo.iCorner(i) - 1)
                '保存日時
                Print #iOutFile, "保存日時：" & pfGetSaveDate(udtJprPrintSetteingInfo.iCorner(i))
    
                For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
                    ' ジャーナル編集データファイル作成処理が正常の場合は
                    If bResult(i, j) <> False Then
                        'その号機が設置されているか？
                        intFileNo = FreeFile
                        strDispImageFileName = Replace(EDIT_DATA_GATECFG, "####", _
                            Format(udtJprPrintSetteingInfo.iCorner(i), "0#") & Format(udtJprPrintSetteingInfo.iGouki(j), "0#")) & ".csv"
                    
                        'ジャーナル編集データファイルをオープン（このファイルが存在しない場合はここには来ない）
                        Open strDispImageFileName For Input As #intFileNo
                
                        '画面表示用データ(csv)をエリアに読み込む
                        k = 0
                        Do While Not EOF(intFileNo)
                            ReDim Preserve ReadFileGateCfg(k)
                            Input #intFileNo, _
                                    ReadFileGateCfg(k).strInfoName, _
                                    ReadFileGateCfg(k).strBunrui_Dai, ReadFileGateCfg(k).strBunrui_Chu, _
                                    ReadFileGateCfg(k).strBunrui_Syo, ReadFileGateCfg(k).strValue, ReadFileGateCfg(k).strChangeFlg
                            k = k + 1
                        Loop
                        'ファイルクローズ
                        Close #intFileNo
                        
                        '号機番号
                        Print #iOutFile, "号機番号：" & Format(udtJprPrintSetteingInfo.iGouki(j), "00") & "号機"
                        
                        'ここから1台分の改札機保守設定データの内容(本文)を印字するループ
                        strNowInfoName = ""
                        strNowDai = ""
                        strNowChu = ""
                        blnInfoNameFlg = False
                        For l = 0 To UBound(ReadFileGateCfg)
                            '情報部名が異なれば、区切りタイトルを出力する ≪情報部名≫
                            If strNowInfoName <> ReadFileGateCfg(l).strInfoName Then
                                Print #iOutFile, ""
                                Print #iOutFile, "　" & ReadFileGateCfg(l).strInfoName
                                strNowInfoName = ReadFileGateCfg(l).strInfoName
                                blnInfoNameFlg = True
                            End If
                            '大分類、中分類が異なる場合は、区切りタイトルを出力する 【大分類-中分類】
                            If strNowDai <> ReadFileGateCfg(l).strBunrui_Dai Or strNowChu <> ReadFileGateCfg(l).strBunrui_Chu Then
                                '区切りタイトル(大分類－中分類)の直前は1行改行させるが、情報部名の直後は改行なし
                                If blnInfoNameFlg = False Then
                                    Print #iOutFile, ""
                                Else
                                    blnInfoNameFlg = False
                                End If
                                Print #iOutFile, "　【" & ReadFileGateCfg(l).strBunrui_Dai & "－" & ReadFileGateCfg(l).strBunrui_Chu & "】"
                                strNowDai = ReadFileGateCfg(l).strBunrui_Dai
                                strNowChu = ReadFileGateCfg(l).strBunrui_Chu
                            End If
                            
                            '変更フラグ ＋ 項目名 ＋ ":" ＋ 設定値を出力
                            '何文字スペースを中間に入れるか？
                            strSyoName = RTrim(ReadFileGateCfg(l).strBunrui_Syo)
                            strValue = RTrim(ReadFileGateCfg(l).strValue)
                            iKoumokuByte = LenB(StrConv(strSyoName, vbFromUnicode))
                            iValueByte = LenB(StrConv(strValue, vbFromUnicode))
                            'ジャーナル1行が最大30バイト
                            iSpaceByte = MAX_JPR_KETA_MAX - DIFF_MARK_LEN - iKoumokuByte - COLON_LEN - iValueByte
                            If iSpaceByte <= 0 Then
                                iSpaceByte = 0
                            End If
                            If ReadFileGateCfg(l).strChangeFlg = "" Then
                                strChangeFlg = DIFF_MARK_STRING_OFF
                            Else
                                strChangeFlg = DIFF_MARK_STRING_ON
                            End If
                            
                            Print #iOutFile, strChangeFlg & strSyoName & Space(iSpaceByte) & "：" & strValue
                        
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
        '出力対象となる改札機が1台も存在しないので、スキップは無しとする。
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

    'MsgBox "異常終了しました。", vbCritical, "出力結果"
    '「ジャーナル印字画面（改札機保守設定データ）：ジャーナルイメージファイル作成異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, lngErrCode)
    JprEdit_GateCfg = False

End Function
'EG30 V32.1.0.1 ADD END

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : IsTaisyoGoki
'//  機能名称  : 指定号機確認処理
'//  機能概要  : イメージファイルに出力する項目は出力対象か確認する
'//
'//              型        名称      意味
'//  引数      : Integer   iCorner   コーナ番号
'//              Integer   iGouki    号機番号
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-19   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function IsTaisyoGoki(iCorner As Integer, iGouki As Integer) As Boolean
    Dim bRet        As Boolean
    Dim i           As Integer
    Dim j           As Integer
    
    bRet = False
    
    If pfCornerGokiCheck(iCorner, iGouki) = False Then
        '未設置の号機なので、画面上チェックされていてもfalse
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
'//  関数名称  : IsTaisyoCorner
'//  機能名称  : 指定コーナ確認処理
'//  機能概要  : イメージファイルに出力する項目は出力対象か確認する
'//
'//              型        名称      意味
'//  引数      : Integer   iCorner   コーナ番号
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-19   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function IsTaisyoCorner(iCorner As Integer) As Boolean
    Dim bRet        As Boolean
    Dim i           As Integer
    
    bRet = False
    'そのコーナは設置されているか？
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
'//  関数名称  : JprEdit_SetteiList
'//  機能名称  : 設定値一覧出力
'//  機能概要  : 設定値一覧ジャーナルのイメージファイルを編集する
'//
'//              型        名称      意味
'//  引数      : Integer   iCorner   コーナ番号
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(7.2.0.1) 2013-06-19   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(7.4.0.1) 2013-07-22   REVISED BY [TCC] T.Nakajima
'//                日またがり出場フリー設定画面対応
'//     REVISIONS :(EG30 V32.1.0.1) 2016-06-17   REVISED BY [TCC] T.Nakajima
'//                2016年度施策対応
'//     REVISIONS :(EG30 V35.3.0.1) 2019-07-03   REVISED BY [TCC] H.Kondoh
'//                2019年度施策対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_SetteiList() As Boolean
    Dim strFilePath As String           '出力ファイルパス
    Dim intCount As Integer             'カウンタ
    Dim intOutFile As Integer           '出力ファイル番号
    Dim intTgtFileNo As Integer         '出力対象設定ファイル番号
    Dim strTgtFileName As String        '出力対象設定ファイル
    Dim strTargetFile() As String       '出力対象ファイル
    Dim intFileNum As Integer
    Dim objFileObj As FileSystemObject  'ファイルシステムオブジェクト
    Dim ReadFileSettei()    As SETTEI_OUTPUT_IMAGE_FILE   'ファイル読込用構造体
    Dim strCsvBuffer        As String
    Dim strCammaArray()     As String
    Dim i As Integer
    Dim strNowDaikomoku     As String
    Dim strNowKomoku        As String
    Dim FsoTS   As TextStream
    Dim iKomkuByte          As Integer '項目名のバイト数
    Dim iValueByte          As Integer '設定値のバイト数
    Dim iSpaceByte          As Integer '中間に挿入するスペースのバイト数
    Dim intJprFile            As Integer
    Dim strNyujoFree(3)       As String
    Dim iSeparatePos          As Integer    '画面名称が「:」で区切られていた場合の区切り位置
    'EG30 V32.1.0.1 ADD START
    Dim strChangeFlg        As String  '変更フラグ
    'EG30 V32.1.0.1 ADD END

    Set objFileObj = New FileSystemObject
    
    On Error GoTo Err_handler
    
    'EG20 V30.1.0.1 ADD START
    '設置駅
    gsGetStationName
    '自改情報
    gsGetGateInfo
    'コーナ名
    gsGetCornerName
    'コーナタイプ
    gsGetCornerType
    'EG20 V30.1.0.1 ADD END
    '画面で指定されたコーナは設置されているか？
    If pfSettingCheck(False) = False Then
        'すべて未設置なのでエラー
        GoTo Err_handler
    End If
    
    '出力対象設定ファイルをオープンする。
    intTgtFileNo = FreeFile
    strTgtFileName = OUTPUT_TARGET_FILE
    
    '出力対象設定ファイルが存在しない場合は異常終了
    If objFileObj.FileExists(strTgtFileName) = False Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_SEARCH_ERROR & ":" & strTgtFileName, 0)
        GoTo Err_handler
    End If
    
    Open strTgtFileName For Input As #intTgtFileNo
    
    '出力対象ファイル数を取得
    Input #intTgtFileNo, intFileNum
    
    '出力対象ファイルを取得
    ReDim strTargetFile(intFileNum - 1)
    For intCount = 0 To UBound(strTargetFile)
        Input #intTgtFileNo, strTargetFile(intCount)
    Next
    
    Close #intTgtFileNo
    
    'EG20 V30.1.0.1 ADD START
    '幹線コーナーに対する出力対象ファイルの内容を確保する
    intTgtFileNo = FreeFile
    strTgtFileName = OUTPUT_TARGET_FILE_KAN
    
    '出力対象設定ファイルが存在しない場合は異常終了
    If objFileObj.FileExists(strTgtFileName) = False Then
        Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, FILE_SEARCH_ERROR & ":" & strTgtFileName, 0)
        GoTo Err_handler
    End If
    
    Open strTgtFileName For Input As #intTgtFileNo
    
    '出力対象ファイル数を取得
    Input #intTgtFileNo, intFileNum
    
    '出力対象ファイルを取得
    ReDim strTargetFileKan(intFileNum - 1)
    For intCount = 0 To UBound(strTargetFileKan)
        Input #intTgtFileNo, strTargetFileKan(intCount)
    Next
    
    Close #intTgtFileNo
    'EG20 V30.1.0.1 ADD END
    
    '////////////////////////////////
    'ジャーナルイメージファイル作成
    '前回の出力済みのジャーナルイメージファイルは消しておく(コーナ単位で追記していくため)
    If Dir(SETTI_TXTFLE) <> "" Then
        Kill SETTI_TXTFLE
    End If
    
    'コーナ単位で設定値一覧のCSVファイルを作成する
    'ヘッダ部作成
    intJprFile = FreeFile
    Open SETTI_TXTFLE For Output As #intJprFile
    PrintHeader intJprFile, "設定値一覧"
    
    '設置駅
    'gsGetStationName   'EG20 V30.1.0.1 DEL
    Print #intJprFile, "設置駅：" & gstrStationName(0)
    'コーナ名
    'gsGetCornerName    'EG20 V30.1.0.1 DEL

    For intCount = 0 To UBound(glngTergetCorner)
        
        If glngTergetCorner(intCount) = CMN_ONOFF.CMN_ON Then
            'コーナ単位で設定ファイル一覧(編集用CSV)作成 OPERATE_SETTI99.csv
            strFilePath = Replace(EDIT_DATA_SETTEI, "##", Format(intCount + 1, "0#"))
            
            '---- 設定一覧テキスト作成 開始
            'ファイル作成
            If objFileObj.FileExists(strFilePath) = True Then
                objFileObj.DeleteFile (strFilePath)
            End If
            Call objFileObj.CreateTextFile(strFilePath)
            
            '出力ファイルをオープンする。
            intOutFile = FreeFile
            Open strFilePath For Output As #intOutFile
    
            'ID設定値を出力
            'If gsubOutput_Id(intCount + 1, intOutFile, True) = False Then      'EG30 V32.1.0.1 DEL
            If gsubOutput_Id_JPR(intCount + 1, intOutFile, True) = False Then   'EG30 V32.1.0.1 ADD
                GoTo Err_handler
            End If
            
            'EG20 V30.1.0.1 DEL START
            '入出場フリーファイルを出力
'            If gsubOutput_Free_InOut(intCount + 1, intOutFile) = False Then
'                GoTo Err_Handler
'            End If
'
'            '祝祭日ファイルを出力
'            If gsubOutput_Shukusai(intCount + 1, intOutFile) = False Then
'                GoTo Err_Handler
'            End If
            'EG20 V30.1.0.1 DEL END
            
            'EG20 V30.1.0.1 ADD START
            If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                '幹線コーナーの場合
                '新幹線不正パラメータを出力
                'If gsubOutput_ParaKan(FILE_PR_FSK, OUTPUT_PRFSK_FILE, intCount + 1, intOutFile) = False Then   'EG30 V32.1.0.1 DEL
                If gsubOutput_ParaKan_JPR(FILE_PR_FSK, OUTPUT_PRFSK_FILE, intCount + 1, intOutFile) = False Then    'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
                '在来線IC判定パラメータを出力
                'If gsubOutput_ParaKan(FILE_PR_ICZ, OUTPUT_PRICZ_FILE, intCount + 1, intOutFile) = False Then   'EG30 V32.1.0.1 DEL
                If gsubOutput_ParaKan_JPR(FILE_PR_ICZ, OUTPUT_PRICZ_FILE, intCount + 1, intOutFile) = False Then    'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
                
                '在来線IC通過処理パラメータを出力
                'If gsubOutput_ParaKan(FILE_PR_ITZ, OUTPUT_PRITZ_FILE, intCount + 1, intOutFile) = False Then   'EG30 V32.1.0.1 DEL
                If gsubOutput_ParaKan_JPR(FILE_PR_ITZ, OUTPUT_PRITZ_FILE, intCount + 1, intOutFile) = False Then    'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
            Else
                '在来コーナーの場合
                '入出場フリーファイルを出力
                'If gsubOutput_Free_InOut(intCount + 1, intOutFile) = False Then    'EG30 V32.1.0.1 DEL
                If gsubOutput_Free_InOut_JPR(intCount + 1, intOutFile) = False Then     'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
                
                '祝祭日ファイルを出力
                'If gsubOutput_Shukusai(intCount + 1, intOutFile) = False Then  'EG30 V32.1.0.1 V32.1.0.1 DEL
                If gsubOutput_Shukusai_JPR(intCount + 1, intOutFile) = False Then   'EG30 V32.1.0.1 ADD
                    GoTo Err_handler
                End If
            End If
            'EG20 V30.1.0.1 ADD END

            Close #intOutFile
            '---- 設定一覧テキスト作成 終了
            
            '出力した編集元データをエリアにセットする
            Set FsoTS = objFileObj.OpenTextFile(strFilePath, ForReading)
            i = 0
            Do Until FsoTS.AtEndOfStream = True
                ReDim Preserve ReadFileSettei(i)
                strCsvBuffer = FsoTS.ReadLine
                'カンマをキーワードに各項目を切り出す。
                strCammaArray = Split(strCsvBuffer, ",")
                ReadFileSettei(i).strDaiKomoku = strCammaArray(0)   '大項目
                ReadFileSettei(i).strKomoku = strCammaArray(1)      '項目名
                ReadFileSettei(i).strValue = strCammaArray(2)       '設定値
                ReadFileSettei(i).strChangeFlg = strCammaArray(3)   '変更フラグ
                
                i = i + 1
            Loop
            FsoTS.Close
            
            '読み込んだエリアからジャーナルイメージファイルを作成する
            
            'コーナ名
            Print #intJprFile, "設置コーナ：" & gstrCornerName(intCount)
            '保存日時
            Print #intJprFile, "保存日時：" & pfGetSaveDate(intCount + 1)

            strNowDaikomoku = ""
            strNowKomoku = ""
            
            For i = 0 To UBound(ReadFileSettei)
                '大項目を出力するか？
                If strNowDaikomoku <> ReadFileSettei(i).strDaiKomoku Then
                    'ただし、NULLの場合は初号機以降なので継続
                    If ReadFileSettei(i).strDaiKomoku <> "" Then
                        'Print #intJprFile, ""  'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
                        If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                            '中項目レベルの切替は改行しない
                            iSeparatePos = InStr(ReadFileSettei(i).strDaiKomoku, ":")
                            If iSeparatePos > 0 Then
                                If Left(strNowDaikomoku, iSeparatePos - 1) <> Left(ReadFileSettei(i).strDaiKomoku, iSeparatePos - 1) Then
                                    '大分類レベルで異なっているので改行
                                    Print #intJprFile, ""
                                Else
                                End If
                            Else
                                Print #intJprFile, ""
                            End If
                        Else
                            Print #intJprFile, ""
                        End If
                        
                        'Print #intJprFile, "【" & ReadFileSettei(i).strDaiKomoku & "】"    'EG20 V30.1.0.1 DEL
                        'EG20 V30.1.0.1 ADD START
                        If gintCornerType(intCount) = CORNER_TYPE_KANSEN Then
                            '幹線コーナの場合
                            '大項目(画面名称が":"で区切られていたらそこで分ける)
                            iSeparatePos = InStr(ReadFileSettei(i).strDaiKomoku, ":")
                            If iSeparatePos > 0 Then
                                '大分類が同じだったら出力しない
                                If Left(strNowDaikomoku, iSeparatePos - 1) <> Left(ReadFileSettei(i).strDaiKomoku, iSeparatePos - 1) Then
                                    'Print #intJprFile, "【" & Left(ReadFileSettei(i).strDaiKomoku, iSeparatePos - 1) & "】"    'EG30 V32.1.0.1 DEL
                                    Print #intJprFile, "　【" & Left(ReadFileSettei(i).strDaiKomoku, iSeparatePos - 1) & "】"   'EG30 V32.1.0.1 ADD
                                    '中項目を出力する
                                    'Print #intJprFile, Mid(ReadFileSettei(i).strDaiKomoku, iSeparatePos + 1)   'EG30 V32.1.0.1 DEL
                                    Print #intJprFile, "　" & Mid(ReadFileSettei(i).strDaiKomoku, iSeparatePos + 1) 'EG30 V32.1.0.1 ADD
                                Else
                                    '大分類までは同じなので、中項目だけを出力する。
                                    'Print #intJprFile, Mid(ReadFileSettei(i).strDaiKomoku, iSeparatePos + 1)   'EG30 V32.1.0.1 DEL
                                    Print #intJprFile, "　" & Mid(ReadFileSettei(i).strDaiKomoku, iSeparatePos + 1) 'EG30 V32.1.0.1 ADD
                                End If
                            Else
                                'Print #intJprFile, "【" & ReadFileSettei(i).strDaiKomoku & "】"    'EG30 V32.1.0.1 DEL
                                Print #intJprFile, "　【" & ReadFileSettei(i).strDaiKomoku & "】"   'EG30 V32.1.0.1 ADD
                            End If
                        Else
                            'Print #intJprFile, "【" & ReadFileSettei(i).strDaiKomoku & "】"    'EG30 V32.1.0.1 DEL
                            Print #intJprFile, "　【" & ReadFileSettei(i).strDaiKomoku & "】"   'EG30 V32.1.0.1 ADD
                        End If
                        'EG20 V30.1.0.1 ADD END
                        strNowDaikomoku = ReadFileSettei(i).strDaiKomoku
                    End If
                End If
                
                '入場フリー設定画面は設定値を改行させる必要がある。
                If ReadFileSettei(i).strDaiKomoku = "入場フリー設定画面" Then
                    '入場フリー1～6の数字は全角に変更する。(仕様にあわせるため)
                    Select Case ReadFileSettei(i).strKomoku
                        Case "入場フリー1"
                            'strNyujoFree(0) = "入場フリー１"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　入場フリー１"    'EG30 V32.1.0.1 ADD
                        Case "入場フリー2"
                            'strNyujoFree(0) = "入場フリー２"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　入場フリー２"    'EG30 V32.1.0.1 ADD
                        Case "入場フリー3"
                            'strNyujoFree(0) = "入場フリー３"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　入場フリー３"    'EG30 V32.1.0.1 ADD
                        Case "入場フリー4"
                            'strNyujoFree(0) = "入場フリー４"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　入場フリー４"    'EG30 V32.1.0.1 ADD
                        Case "入場フリー5"
                            'strNyujoFree(0) = "入場フリー５"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　入場フリー５"    'EG30 V32.1.0.1 ADD
                        Case "入場フリー6"
                            'strNyujoFree(0) = "入場フリー６"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　入場フリー６"    'EG30 V32.1.0.1 ADD
                        Case "入場フリー7"
                            'strNyujoFree(0) = "入場フリー７"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　入場フリー７"    'EG30 V32.1.0.1 ADD
                        Case "入場フリー8"
                            'strNyujoFree(0) = "入場フリー８"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　入場フリー８"    'EG30 V32.1.0.1 ADD
                        Case "入場フリー9"
                            'strNyujoFree(0) = "入場フリー９"   'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　入場フリー９"    'EG30 V32.1.0.1 ADD
                    End Select
                    'EG30 V32.1.0.1 DEL START
'                    strNyujoFree(1) = MidByte(ReadFileSettei(i).strValue, 1, 8) & Space(6) & MidByte(ReadFileSettei(i).strValue, 11, 16) '開始日時
'                    strNyujoFree(2) = MidByte(ReadFileSettei(i).strValue, 28, 8) & Space(6) & MidByte(ReadFileSettei(i).strValue, 38, 16) '終了日時
'                    strNyujoFree(3) = MidByte(ReadFileSettei(i).strValue, 55, 4) & Space(22) & MidByte(ReadFileSettei(i).strValue, 61, 4) '券種
                    'EG30 V32.1.0.1 DEL END
                    'EG30 V32.1.0.1 ADD START
                    If ReadFileSettei(i).strChangeFlg = "　" Then
                        strChangeFlg = DIFF_MARK_STRING_OFF
                    Else
                        strChangeFlg = DIFF_MARK_STRING_ON
                    End If
                    strNyujoFree(1) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 1, 8) & Space(4) & MidByte(ReadFileSettei(i).strValue, 11, 16)  '開始日時
                    strNyujoFree(2) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 28, 8) & Space(4) & MidByte(ReadFileSettei(i).strValue, 38, 16) '終了日時
                    strNyujoFree(3) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 55, 4) & Space(20) & MidByte(ReadFileSettei(i).strValue, 61, 4) '券種
                    'EG30 V32.1.0.1 ADD END
                    '1行出力
                    Print #intJprFile, strNyujoFree(0)
                    Print #intJprFile, strNyujoFree(1)
                    Print #intJprFile, strNyujoFree(2)
                    Print #intJprFile, strNyujoFree(3)
                'EG20 V7.4.0.1 ADD START
                '日またがり出場フリー設定画面は設定を改行させる必要がある
                'ElseIf ReadFileSettei(i).strDaiKomoku = "日またがり出場フリー設定画面" Then    'EG30 V32.1.0.1 DEL
                ElseIf ReadFileSettei(i).strDaiKomoku = "日跨り出場フリー設定画面" Then         'EG30 V32.1.0.1 ADD
                    '入場フリー1～6の数字は全角に変更する。(仕様にあわせるため)
                    Select Case ReadFileSettei(i).strKomoku
                        'Case "日またがり出場フリー1"   'EG30 V32.1.0.1 DEL
                        Case "日跨り出場フリー1"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "日またがり出場フリー１" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　日跨り出場フリー１"  'EG30 V32.1.0.1 ADD
                        'Case "日またがり出場フリー2"   'EG30 V32.1.0.1 DEL
                        Case "日跨り出場フリー2"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "日またがり出場フリー２" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　日跨り出場フリー２"  'EG30 V32.1.0.1 ADD
                        'Case "日またがり出場フリー3"   'EG30 V32.1.0.1 DEL
                        Case "日跨り出場フリー3"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "日またがり出場フリー３" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　日跨り出場フリー３"  'EG30 V32.1.0.1 ADD
                        'Case "日またがり出場フリー4"   'EG30 V32.1.0.1 DEL
                        Case "日跨り出場フリー4"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "日またがり出場フリー４" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　日跨り出場フリー４"  'EG30 V32.1.0.1 ADD
                        'Case "日またがり出場フリー5"   'EG30 V32.1.0.1 DEL
                        Case "日跨り出場フリー5"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "日またがり出場フリー５" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　日跨り出場フリー５"  'EG30 V32.1.0.1 ADD
                        'Case "日またがり出場フリー6"   'EG30 V32.1.0.1 DEL
                        Case "日跨り出場フリー6"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "日またがり出場フリー６" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　日跨り出場フリー６"  'EG30 V32.1.0.1 ADD
                        'Case "日またがり出場フリー7"   'EG30 V32.1.0.1 DEL
                        Case "日跨り出場フリー7"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "日またがり出場フリー７" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　日跨り出場フリー７"  'EG30 V32.1.0.1 ADD
                        'Case "日またがり出場フリー8"   'EG30 V32.1.0.1 DEL
                        Case "日跨り出場フリー8"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "日またがり出場フリー８" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　日跨り出場フリー８"  'EG30 V32.1.0.1 ADD
                        'Case "日またがり出場フリー9"   'EG30 V32.1.0.1 DEL
                        Case "日跨り出場フリー9"        'EG30 V32.1.0.1 ADD
                            'strNyujoFree(0) = "日またがり出場フリー９" 'EG30 V32.1.0.1 DEL
                            strNyujoFree(0) = "　日跨り出場フリー９"  'EG30 V32.1.0.1 ADD
                    End Select
                    'EG30 V32.1.0.1 DEL START
                    'strNyujoFree(1) = MidByte(ReadFileSettei(i).strValue, 1, 8) & Space(6) & MidByte(ReadFileSettei(i).strValue, 11, 16) '開始日時
                    'strNyujoFree(2) = MidByte(ReadFileSettei(i).strValue, 28, 8) & Space(6) & MidByte(ReadFileSettei(i).strValue, 38, 16) '終了日時
                    'strNyujoFree(3) = MidByte(ReadFileSettei(i).strValue, 55, 4) & Space(22) & MidByte(ReadFileSettei(i).strValue, 61, 4) '券種
                    'EG30 V32.1.0.1 DEL END
                    'EG30 V32.1.0.1 ADD START
                    If ReadFileSettei(i).strChangeFlg = "　" Then
                        strChangeFlg = DIFF_MARK_STRING_OFF
                    Else
                        strChangeFlg = DIFF_MARK_STRING_ON
                    End If
                    strNyujoFree(1) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 1, 8) & Space(4) & MidByte(ReadFileSettei(i).strValue, 11, 16)  '開始日時
                    strNyujoFree(2) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 28, 8) & Space(4) & MidByte(ReadFileSettei(i).strValue, 38, 16)  '終了日時
                    strNyujoFree(3) = strChangeFlg & MidByte(ReadFileSettei(i).strValue, 55, 4) & Space(20) & MidByte(ReadFileSettei(i).strValue, 61, 4) '券種
                    'EG30 V32.1.0.1 ADD END
                    '1行出力
                    Print #intJprFile, strNyujoFree(0)
                    Print #intJprFile, strNyujoFree(1)
                    Print #intJprFile, strNyujoFree(2)
                    Print #intJprFile, strNyujoFree(3)
                'EG20 V7.4.0.1 ADD END
                Else
                
                    If ReadFileSettei(i).strKomoku = "" Then
                        '前回の項目名を使う
                        ReadFileSettei(i).strKomoku = strNowKomoku
                    End If
                    strNowKomoku = ReadFileSettei(i).strKomoku
                    '例外項目
                    '基本はテキスト出力のCSVの文言を使うが、下記の項目だけはジャーナル仕様にあわせる必要がある。
                    Select Case ReadFileSettei(i).strKomoku
                        Case "時間帯開始時刻", _
                             "時間帯終了時刻", _
                             "有効時間帯開始時刻", _
                             "有効時間帯終了時刻"
                            '「全号機：99時99分」→ 「 全号機 99時99分」に変換する
                            ReadFileSettei(i).strValue = Space(1) & Replace(ReadFileSettei(i).strValue, "：", " ")
                        
                        'Case "通過サービス不正保留"    'EG30 V32.1.0.1 DEL
                        'EG30 V32.1.0.1 ADD START
'EG30 V35.3.0.1 DEL Start
'                        Case "通過サービス不正保留", _
'                             "IC会社間経路連続性", _
'                             "オートチャージ機能", _
'                             "定期券フェールセーフ", _
'                             "普通券フェールセーフ", _
'                             "ICカード期限前予告", _
'                             "ICカード期限後案内", _
'                             "無人モード音声案内", _
'                             "IC案内表示画面"
'                        'EG30 V32.1.0.1 ADD END
'EG30 V35.3.0.1 ADD End
'EG30 V35.3.0.1 ADD Start
                        Case "通過サービス不正保留", _
                             "IC会社間経路連続性", _
                             "オートチャージ機能", _
                             "定期券フェールセーフ", _
                             "普通券フェールセーフ", _
                             "ICカード期限前予告", _
                             "ICカード期限後案内", _
                             "無人モード音声案内", _
                             "IC案内表示画面", _
                             "開始年月日", _
                             "終了年月日"
'EG30 V35.3.0.1 ADD End
                            '「通過サービス不正保留全号機：xx」→ 「通過サービス不正保留 全号機：xx」に変換する
                            ReadFileSettei(i).strValue = Space(1) & ReadFileSettei(i).strValue
                        
                        Case "無人モード動作設定"
                            '設定値を左詰にする
                            ReadFileSettei(i).strValue = ReadFileSettei(i).strValue & Space(8 - LenB(ReadFileSettei(i).strValue))
                    End Select
                
                    '項目名出力
                    '何文字スペースを中間に入れるか？
                    iKomkuByte = LenB(StrConv(ReadFileSettei(i).strKomoku, vbFromUnicode))
                    iValueByte = LenB(StrConv(ReadFileSettei(i).strValue, vbFromUnicode))
                    'ジャーナル1行が最大30バイト
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
                    '1行出力
                    'Print #intJprFile, ReadFileSettei(i).strKomoku & Space(iSpaceByte) & ReadFileSettei(i).strValue    'EG30 V32.1.0.1 DEL
                    'EG30 V32.1.0.1 ADD START
                    If ReadFileSettei(i).strChangeFlg = "　" Then
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
    
'エラー処理
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
    'エラーログの出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
    JprEdit_SetteiList = False
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : JprEdit_SimekiriOffline
'//  機能名称  : 締切オフライン出力ジャーナル編集処理
'//  機能概要  : 締切オフライン出力のイメージファイルを編集する
'//
'//              型        名称      意味
'//  引数      :
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.1.0.1) 2014-03-25   CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_SimekiriOffline() As Boolean
    
    Dim objFso As New FileSystemObject                  ' ファイルシステムオブジェクト
    Dim objTs   As TextStream
    Dim bProceed As Boolean                             ' 締切処理開始フラグ
    Dim nListCnt As Integer                             ' ファイル格納数
    Dim szSaveFolder As String                          ' 保存先フォルダ
    Dim szFileName As String                            ' ファイル名
    Dim iResponse As Integer
    Dim Index       As Integer                          'インデックス
    Dim iOutFile    As Integer
    
    On Error GoTo ErrorHandler                          ' エラーハンドルの登録
    
    'EG20 V30.1.0.1 ADD START
    ' コーナ名取得
    gsGetCornerName
    ' コーナタイプ取得
    gsGetCornerType
    
    ' 駅名取得
    gsGetStationName
    ' EG20 V30.1.0.1 ADD END
    
    'チェックされたコーナは設置されているか？（どれかひとつでもあればOK)
    If pfSettingCheck(False) = False Then
        'すべて未設置のコーナなのでエラー
        GoTo ErrorHandler
    End If
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // 初期化
    Index = 0
    Erase gOfflineFileList

    ReDim Preserve gOfflineFileList(0)
    bProceed = False
    nListCnt = 0
    
    ' /////////////////////////////////////////////////////////////////////////
    ' // ジャーナルイメージファイル作成
    'EG20 V30.1.0.1 DEL START
    ' コーナ名取得
    'gsGetCornerName
    
    ' 駅名取得
    'gsGetStationName
    'EG20 V30.1.0.1 DELEND
    
    'ジャーナルイメージファイルをオープン
    iOutFile = FreeFile
    Open SIMEKIRI_TXTFILE For Output As #iOutFile
    
    'ヘッダ部を出力
    PrintHeader iOutFile, "締切オフライン出力"
    
    '設置駅/設置コーナ
    Print #iOutFile, "設置駅：" & gstrStationName(0)
    
    For Index = 0 To UBound(glngTergetCorner)
    
        If glngTergetCorner(Index) = CMN_ONOFF.CMN_ON Then
            
            ' /////////////////////////////////////////////////////////////////////////
            ' // 締切出力データは存在するか？（D:\KANSI\SHUKEI\SEND_DATA\SIMEKIRI##.DAT）
            szFileName = Replace(FILENAME_SIMEKIRIDAT, "##", Format(Index + 1, "0#"))
            If objFso.FileExists(szFileName) = True Then              ' ファイル名の取得チェック
                nListCnt = nListCnt + 1                             ' ファイル数のカウンタをアップする
                ReDim Preserve gOfflineFileList(nListCnt)           ' ファイル名格納エリアを拡張する
                gOfflineFileList(nListCnt - 1) = szFileName         ' ファイルパスを格納
                bProceed = True
            End If
            
                
            If bProceed = False Then
                GoTo ErrorHandler
            End If
            
            ' /////////////////////////////////////////////////////////////////////////
            ' // 編集データファイルを作成
            ' // コーナごとの締切テキストファイルを作成
            bProceed = sOutPutOfflineData(Index)
            If bProceed = False Then
                GoTo ErrorHandler
            End If
            
            Print #iOutFile, "設置コーナ：" & gstrCornerName(Index)
            Print #iOutFile, ""
            
            '1コーナ分の締切データを読み込む
            szFileName = Replace(EDIT_DATA_SIMEKIRI, "##", Format(Index + 1, "0#"))
            Set objTs = objFso.OpenTextFile(szFileName, ForReading)
            Print #iOutFile, objTs.ReadAll
            objTs.Close
            Set objFso = Nothing
        End If
    Next Index
    
    'フッタ部出力
    Print #iOutFile, FOOTER_STRING
    
    Close #iOutFile
    Set objFso = Nothing

    JprEdit_SimekiriOffline = True
    Exit Function

' /////////////////////////////////////////////////////////
' // エラー処理
ErrorHandler:
    'Call MsgBox("異常終了しました。", vbOKOnly, "オフライン出力結果")
    If iOutFile > 0 Then
        Close #iOutFile
    End If

    Set objFso = Nothing

    'エラーログの出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)

End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 ALL Rights Reserved
'//
'//  関数名称  : sOutPutOfflineData
'//  機能名称  : オフラインデータ媒体出力処理
'//  機能概要  : コーナごとに締切ファイル(テキストファイル)を作成する。
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.1.0.1) 2014-03-25  CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function sOutPutOfflineData(dwCornerIdx As Integer) As Boolean
            
    Dim szFileName As String                            ' ファイル名
    Dim lResult As Long                                 ' 処理結果
    Dim dwSequense As Long                              ' シーケンス番号

    ' //////////////////////////////////////////////////////////////
    ' // ファイル作成処理
    ' // 参照元ファイルSIMEKIRI##.DATのファイル名を作成
    szFileName = Replace(FILENAME_SIMEKIRIDAT, "##", Format(dwCornerIdx + 1, "0#"))
    
    ' //////////////////////////////////////////////////////////////
    ' // コーナごとの締切データ(テキスト)を作成
    dwSequense = 0                              ' シーケンス番号:0固定
    'EG20 V30.1.0.1 DEL START
'    lResult = dllCreateShimekiriFileJpr(dwCornerIdx + 1, dwSequense, _
'                                        PATH_WORK, _
'                                        szFileName)
    'EG20 V30.1.0.1 ADD START
    If gintCornerType(dwCornerIdx) = CORNER_TYPE_KANSEN Then
        '幹線コーナならば幹線コーナ用の関数を呼び出す
        lResult = dllCreateShimekiriFileJprKan(dwCornerIdx + 1, dwSequense, _
                                                PATH_WORK, _
                                                szFileName)
    Else
        '在来コーナならば在来コーナ用の関数を呼び出す
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
'//  関数名称  : JprEdit_KadoVersion
'//  機能名称  : 稼動バージョンジャーナルイメージファイル作成
'//  機能概要  : 稼動バージョンジャーナルイメージファイルを作成する
'//
'//              型        名称      意味
'//  引数      : なし
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     ORIGINAL  :(EG20 V30.1.0.1) 2014-05-07   CODED   BY [TCC] T.Nakajima
'//             北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_KadoVersion() As Boolean

    Dim strOutputFile As String         '出力ファイル
    Dim lngRet As Long                  '関数返り値
    Dim lngErrCode As Long              'エラーコード
    Dim iOutFile    As Integer          'ファイル番号
    Dim ReadFileKado()    As KADO_VER_DISP_IMAGE_FILE '稼働バージョン一覧元データ
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim l           As Integer
    Dim strDispImageFileName As String
    Dim objFs       As New FileSystemObject
    Dim intFileNo   As Integer
    Dim iHeadFlg    As Integer
    
    
    On Error GoTo Err_handler
    
    '画面で指定されたコーナは設置されているか？
    If pfSettingCheck(True) = False Then
        'すべて未設置のコーナ、号機なのでエラーとする
        GoTo Err_handler
    End If
    
    '出力ファイル名編集
    strOutputFile = KADOVER_TXTFILE
    
    '// コーナ名を一通り取得 取得結果はgstrCornerName(0 to 5)に入っている
    gsGetCornerName
    'EG20 V30.0.1.1 ADD START
    ' コーナタイプ取得
    gsGetCornerType
    'EG20 V30.0.1.1 ADD END

    
    '駅名を取得   取得結果はgstrStationName(0 to 5)に入っている
    gsGetStationName
    
    iHeadFlg = 0
    
    'ファイル出力関数をCall
    'チェックされているコーナ、号機分のバージョンファイルをいったん出力
    For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
            'そのコーナ、号機は設置されているか？
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
                
                '異常終了時はエラー処理へ
                If lngRet = 0 Then
                    GoTo Err_handler
                    Exit Function
                End If
                
                'ファイルが存在しない場合はエラー処理へ
                If objFs.FileExists(strDispImageFileName) = False Then
                    GoTo Err_handler
                    Exit Function
                End If
            End If
        Next j
    Next i
    
    '稼働バージョン一覧 ジャーナルイメージファイルを作成
    iOutFile = FreeFile
    Open strOutputFile For Output As #iOutFile
    
    'ヘッダー部
    PrintHeader iOutFile, "稼働バージョン一覧"
    
    '設置駅
    Print #iOutFile, "設置駅：" & gstrStationName(0)
    Print #iOutFile, ""
    
    '画面表示用ファイルをオープン
    For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        Erase ReadFileKado
        If i > 0 Then
            '1コーナ目は全体バージョンを表示してからコーナ名を出力
            'そのコーナは設置されているか？
            'If pfCornerGokiCheck(udtJprPrintSetteingInfo.iCorner(i)) = True Then
            If IsTaisyoCorner(udtJprPrintSetteingInfo.iCorner(i)) = True Then
                '対象コーナであっても対象号機がないかもしれない
                For j = 0 To 15
                    If IsTaisyoGoki(udtJprPrintSetteingInfo.iCorner(i), j + 1) = True Then
                        Print #iOutFile, "コーナ名：" & gstrCornerName(udtJprPrintSetteingInfo.iCorner(i) - 1)
                        Exit For
                    End If
                Next j
                        
            End If
        End If
    
        For j = 0 To udtJprPrintSetteingInfo.iGoukiCount - 1
            'その号機が設置されているか？
            If pfCornerGokiCheck(udtJprPrintSetteingInfo.iCorner(i), udtJprPrintSetteingInfo.iGouki(j)) = True Then
    
                intFileNo = FreeFile
                strDispImageFileName = Replace(EDIT_DATA_KADOVERSION, "####", _
                    Format(udtJprPrintSetteingInfo.iCorner(i), "0#") & Format(udtJprPrintSetteingInfo.iGouki(j), "0#")) & ".csv"
                
                Open strDispImageFileName For Input As #intFileNo
        
                '画面表示用データ(csv)をエリアに読み込む
                k = 0
                Do While Not EOF(intFileNo)
                    ReDim Preserve ReadFileKado(k)
                    'intKishu, intCorner, intGokiDiv, strName, strMaker, strVer, strDate
                    Input #intFileNo, _
                            ReadFileKado(k).strKishu, ReadFileKado(k).strCorner, ReadFileKado(k).strGokiDiv, _
                            ReadFileKado(k).strName, ReadFileKado(k).strMaker, ReadFileKado(k).strVer, ReadFileKado(k).strDate
                    k = k + 1
                Loop
                'ファイルクローズ
                Close #intFileNo
                
                '最初のループだけ全体情報を表示
                'If i = 0 And j = 0 Then
                If iHeadFlg = 0 Then
                    
                    '統合監視盤全体バージョン
                    Print #iOutFile, "統合監視盤全体バージョン"
                    Print #iOutFile, ReadFileKado(0).strVer
                    
                    '統合監視盤
                    Print #iOutFile, "統合監視盤"
                    Print #iOutFile, ReadFileKado(1).strVer
                    
                    'ＩＤＵバージョン
                    Print #iOutFile, "ＩＤＵ"
                    Print #iOutFile, ReadFileKado(2).strVer
                    
                    'ＬＤＵバージョン
                    Print #iOutFile, "ＬＤＵ"
                    Print #iOutFile, ReadFileKado(3).strVer
                    Print #iOutFile, ""
                    
                    '操作卓
                    Print #iOutFile, "操作卓"
                    Print #iOutFile, ReadFileKado(4).strVer
                    Print #iOutFile, ""
                    
                    'コーナ名
                    Print #iOutFile, "コーナ名：" & gstrCornerName(udtJprPrintSetteingInfo.iCorner(i) - 1)
                    
                    iHeadFlg = 1
                End If
    
                '号機番号
                Print #iOutFile, "号機番号：" & Format(udtJprPrintSetteingInfo.iGouki(j), "00") & "号機"
                '各プログラムバージョン(6行目から各プログラムバージョン)
                For l = 0 To k - 1
                    If ReadFileKado(l).strKishu = "06" Then
                        '予備の場合はバージョンを出さない
                        If ReadFileKado(l).strName = "予備１" Or ReadFileKado(l).strName = "予備２" Then
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

    'MsgBox "異常終了しました。", vbCritical, "出力結果"
    '「稼働バージョン管理画面：稼働バージョン情報媒体出力処理異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, lngErrCode)
    JprEdit_KadoVersion = False

End Function


'///////////////////////////////////////////////////////////////////
'//    (C) Copyright TOSHIBA Corporation 2013 All Right Reserved
'//
'//  関数名称 : JprEdit_EkimuId
'//  概要     : 駅務機器IDジャーナルイメージファイル作成処理
'//  説明     : 駅務機器IDジャーナルイメージファイルを作成する
'//  ﾊﾟﾗﾒｰﾀ   :
'//           :
'//
'//  ORIGINAL  ：(EG20 V7.2.0.1) 2013-06-26  CODED BY  [TCC] T.Nakajima
'//  REVISIONS ：(x.x.x.x) xxxx-xx-xx  CODED BY  [XXX]
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_EkimuId() As Boolean
    
    Dim sEkimuIDFile    As String   '駅務機器IDファイルパス
    Dim iRet            As Integer  'INI取得戻り値
    Dim sFolder         As String * MAX_PATH_SIZE  'フォルダ名
    Dim sFile           As String   'ファイル名
    Dim MyName          As String   'ファイル検索結果
    Dim bRet            As Boolean  '戻り値
    Dim lngErrCode      As Long     'エラーコード
    Dim intFileNo       As Integer  'ファイル番号
    Dim strWork         As String   '作業エリア
    Dim dwErrsts        As Long
    Dim sFolderName     As String
    Dim objFso          As New FileSystemObject
    Dim objTs           As TextStream
    
        
    On Error GoTo Err_handler
    sFolder = ""
    
    '処理結果：正常時は画面表示処理
    iRet = GetPrivateProfileString(IDU_SECTION_NAME, _
                                   IDU_EKIMUID_KEY, _
                                   EKIMU_DEFU, sFolder, Len(sFolder), _
                                   PATH_IDU_INI_FILE)
    If iRet = 0 Then
      sFolder = EKIMU_DEFU
    End If
    sEkimuIDFile = ""
    '要求種別値よりファイル名作成
    sFile = Replace(EKIMU_ID_FILE, "##", Format(iSendType, "0#"))
    If iRet = 0 Then
       sFolderName = RTrim(sFolder)
    Else
       sFolderName = Mid(sFolder, 1, iRet)
    End If
    'パス変換処理
    sFolderName = pfChangeFolderName(sFolderName)
    '駅務機器IDファイルパス作成
    sEkimuIDFile = sFolderName & "\" & sFile
    'ファイル有無チェック
    If Dir(sEkimuIDFile, vbNormal) = "" Then
       Exit Function
    End If
    
    '/////////////////////////////////////////////////////////////////////
    '//保守専用関数：駅務機器ID画面表示用ファイル作成処理
    '////////////////////////////////////////////////////////////////////
    bRet = dllEKIMUKIKI(sEkimuIDFile, dwErrsts, MN_VERSI_FILE, PATH_IDU_APP, 1) 'V1.8.0.1 ADD
    
    If bRet = False Then
        GoTo Err_handler
        Exit Function
    End If
    
    
    '/////////////////////////////////////////////////////////////////////
    '//ジャーナルイメージファイルを作成
    '////////////////////////////////////////////////////////////////////
    intFileNo = FreeFile
    Open EKIMUKIKI_ID_TXTFILE For Output As #intFileNo
    
    'ヘッダ部出力
    PrintHeader intFileNo, "駅務機器ＩＤ出力"
    
    '設置駅名
    gsGetStationName
    Print #intFileNo, "設置駅：" & gstrStationName(0)
    Print #intFileNo, ""
    
    'データ部をつなげる
    Set objTs = objFso.OpenTextFile(MN_VERSI_FILE, ForReading)
    Print #intFileNo, objTs.ReadAll
    objTs.Close
    Set objFso = Nothing
    
    'フッタ部作成
    Print #intFileNo, FOOTER_STRING
    
    Close #intFileNo
    
    JprEdit_EkimuId = True
    
    Exit Function

Err_handler:

    If intFileNo > 0 Then
        Close #intFileNo
    End If
    
    
    Set objFso = Nothing

    'MsgBox "異常終了しました。", vbCritical, "出力結果"
    '「稼働バージョン管理画面：稼働バージョン情報媒体出力処理異常」ログ出力
    lngErrCode = EDAI_KANSHI + ECHU_HOSHU + ESHO_FILE_IO + ECOD_AFILE
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, lngErrCode)
    JprEdit_EkimuId = False
    
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2008 All Rights Reserved
'//
'//  関数名称  : pfChangeFolderName
'//  機能名称  : フォルダパス変換処理
'//  機能概要  : INIファイルより取得したフォルダ定義の変換を行う。
'//
'//              型        名称         意味
'//  引数      : String sFolderName    [IN]INI定義
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(1.4.0.1) 2009-03-23   CODED   BY [TCC] S.Terao
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考 ：
'///////////////////////////////////////////////////////////////////
Private Function pfChangeFolderName(sFolderName As String) As String
   Dim iPath As Integer
   Dim sRootPath As String
   Dim sFolder As String
      
   '「￥」位置を取得
   iPath = InStr(sFolderName, "\")
   If iPath = 0 Then
     sRootPath = Mid(sFolderName, 1)
   Else
     '「￥」前文字列を取得
     sRootPath = Mid(sFolderName, 1, iPath - 1)
     '「￥」後文字列を取得
     sFolder = Mid(sFolderName, iPath + 1)
   End If
   Select Case sRootPath
      Case APL
        'アプリルート
        sRootPath = PATH_IDU_APP
      Case LOG
        'ログルート
        sRootPath = PATH_IDU_LOG
      Case Data
        'DBルート
        sRootPath = PATH_IDU_DB
      Case BACKUP
        'バックアップルート
        sRootPath = PATH_BUC
   End Select
    'パス連結
    pfChangeFolderName = sRootPath + "\" + sFolder
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : JprEdit_TukaData
'//  機能名称  : 通過データ/利用金額ジャーナルイメージファイル作成
'//  機能概要  : 通過データ/利用金額ジャーナルイメージファイルを作成する
'//
'//              型        名称      意味
'//  引数      : long      dwDataKind データ種別    通過媒体：306010
'//                                                 利用媒体：306020
'//
'//              型        値        意味
'//  戻り値    : Boolean　　　　　　[OUT]戻り値
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG20 V30.1.0.1) 2014-04-01   CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線開業対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function JprEdit_TukaData(dwDataKind As Long) As Boolean
    
    Dim strFilePath As String           '出力ファイルパス
    Dim intCount As Integer             'カウンタ
    'EG20 V30.1.0.1 DEL START
'    Dim intOutFile As Integer           '出力ファイル番号
'    Dim strBaitaiFileName As String     ' 媒体出力ファイル TUKAコーナ名YYYYMMDDhhmmss.csv ICRIYOコーナ名YYYYMMDDhhmmss.csv
'    Dim ReadFileBaitai()  As BAITAI_OUTPUT_IMAGE_FILE '媒体出力ファイル
'    Dim strLineCount()  As String
'    Dim i As Integer
'    Dim j As Integer
'    Dim k As Integer
'    Dim l As Integer
'    Dim strCammaArray() As String   'カンマ区切りで1項目ずつ取り出したデータ

'    Dim fso As New FileSystemObject
'    Dim FsoTS As TextStream
    
'    Dim iKomokuMaxCnt       As Integer      ' 集計データ項目の最大数
'    Dim iStartLineKaisatu   As Integer      ' 改札側データの開始行（ＣＳＶファイルの）
'    Dim iStartLineShusatu   As Integer      ' 集札側データの開始行（ＣＳＶファイルの）
    'EG20 V30.1.0.1 DEL END
    
    'Dim intJprFile        As Integer
    
    On Error GoTo Err_handler
    '画面で指定されたコーナは設置されているか？
    If pfSettingCheck(False) = False Then
        'すべて未設置なのでエラー
        GoTo Err_handler
    End If
  
    '////////////////////////////////////////////////
    '// 設置駅・コーナ名を一通り取得
    gsGetStationName
    gsGetCornerName
    gsGetCornerType
    gsGetShukeiKoumoku     '集計項目の出力有無を取得    EG20 V30.1.0.1 ADD

   
    'コーナ単位で処理
    
    '/////////////////////////////////////////////
    '// ジャーナルイメージファイル作成
    
    '出力ファイルをオープンする。
    intJprFile = FreeFile
    If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
        Open TUKA_TXTFILE For Output As #intJprFile
    ElseIf dwDataKind = Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI Then
        Open ICRIYO_TXTFILE For Output As #intJprFile
    Else
        JprEdit_TukaData = False
        Exit Function
    End If

   'タイトル表示
   If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
        'EG20 V30.1.0.1 DEL START （在来と幹線によって設定値が異なるのでイメージファイル処理へ移動）
'        iKomokuMaxCnt = MAX_KOMOKU_NUM_TUKA
'        iStartLineKaisatu = 6   '改札側の明細は元ファイル(CSV)配列の(6)から
'        iStartLineShusatu = 60  '集札側の明細は元ファイル(CSV)配列の(60)から
        'EG20 V30.1.0.1 DEL END
        PrintHeader intJprFile, "通過データ出力"
    Else
        'EG20 V30.1.0.1 DEL START （在来と幹線によって設定値が異なるのでイメージファイル処理へ移動）
'        iKomokuMaxCnt = MAX_KOMOKU_NUM_KINGAKU
'        iStartLineKaisatu = 6   '改札側の明細は元ファイル(CSV)配列の(6)から
'        iStartLineShusatu = 25  '集札側の明細は元ファイル(CSV)配列の(60)から
        'EG20 V30.1.0.1 DEL END
        PrintHeader intJprFile, "利用金額データ出力"
    End If

    '設置駅・コーナ名出力
    Print #intJprFile, "設置駅：" & gstrStationName(0)
    
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
'            If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then       '通過データ
'                strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "TUKA" & gstrStationName(intCount) & gstrCornerName(intCount) & "*.csv")
'            ElseIf dwDataKind = Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI Then    '利用金額データ
'                strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "ICRIYO" & gstrStationName(intCount) & gstrCornerName(intCount) & "*.csv")
'            Else
'                JprEdit_TukaData = False
'                Exit Function
'            End If
'
'            '////////////////////////////////////////////////
'            '// 通過データ/利用金額の媒体出力ファイルを取得
'            'ファイル番号取得
'            '駅名称＋コーナ名称yyyymmddhhmmss.csv
'            Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForAppending)
'            j = FsoTS.Line
'            FsoTS.Close
'
'            ReDim strLineCount(j) As String         'CSVファイルを1行ずつ入れておく
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
'            '媒体出力ファイルイメージ構造体にセットする
'            ReDim ReadFileBaitai(j) As BAITAI_OUTPUT_IMAGE_FILE         'ファイル読込用エリア
'            l = 0
'            If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
'
'                For i = 0 To j - 1
'                    Select Case i
'                        Case 0, 1, 2, 3, 4, 57, 58    'TUKAxxxx.csvの1～4行目まではタイトルなので、項目名にセット
'                            ReadFileBaitai(i).strKomokuName = strLineCount(i)
'                        Case Else
'                            'カンマ区切りを1項目ずつ取り出す。
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
'                        Case 0, 1, 2, 3, 4, 22, 23    'ICRIYOxxxx.csvの1～4行目まではタイトルなので、項目名にセット
'                            ReadFileBaitai(i).strKomokuName = strLineCount(i)
'                        Case Else
'                            'カンマ区切りを1項目ずつ取り出す。
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
'            Print #intJprFile, "設置コーナ：" & gstrCornerName(intCount)
'            Print #intJprFile, ""
'
'            If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
'                Print #intJprFile, "【通過データ】"
'            Else
'                Print #intJprFile, "【ＩＣカード利用金額データ】"
'            End If
'            '/////////////////////
'            '改札側データの出力
'            Print #intJprFile, "改札側通過合計"
'
'            For i = 0 To MAX_KOMOKU_NUM_TUKA - 1
'                If RTrim(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName) = "0" Then
'                    '項目名に0がセットされていたら以降は出力しない
'                    Exit For
'                Else
'                    '締切オフラインジャーナルとあわせるための例外処理
'                    If RTrim(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName) = "その他IC (小)" Then
'                        ReadFileBaitai(i + iStartLineKaisatu).strKomokuName = "その他IC(小)" & Space(38)   'スペースを除く
'                    End If
'
'                    Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName, vbFromUnicode), 20), vbUnicode) _
'                    & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLineKaisatu).strGoukei), "#,0"), 10)
'                End If
'            Next i
'            Print #intJprFile, ""
'
'            '/////////////////////
'            '集札側データの出力
'            Print #intJprFile, "集札側通過合計"
'
'            For i = 0 To MAX_KOMOKU_NUM_TUKA - 1
'                If RTrim(ReadFileBaitai(i + iStartLineShusatu).strKomokuName) = "0" Then
'                    '項目名に0がセットされていたら以降は出力しない
'                    Exit For
'                Else
'                    '締切オフラインジャーナルとあわせるための例外処理
'                    If RTrim(ReadFileBaitai(i + iStartLineShusatu).strKomokuName) = "その他IC (小)" Then
'                        ReadFileBaitai(i + iStartLineShusatu).strKomokuName = "その他IC(小)" & Space(38)    'スペースを除く
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
    
'エラー処理
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
    'エラーログの出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
    JprEdit_TukaData = False
                                      
End Function


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : PadLeft
'//  機能名称  : 右寄せ
'//  機能概要  : 指定の文字数になるまで先頭を文字で埋める。
'//
'//              型        名称         意味
'//  引数      : string    strTarget    処理対象文字列
'//              Integer   iLength      文字の長さ
'//              string    chOne        埋める文字(省略時は半角スペース)
'//
'//              型        値        意味
'//  戻り値    : string    右寄せされた文字列
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
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
'//  関数名称  : PadRight
'//  機能名称  : 左寄せ（末尾をスペースで埋める)
'//  機能概要  : 指定の文字数になるまで先頭を文字で埋める。
'//
'//              型        名称         意味
'//  引数      : string    strTarget    処理対象文字列
'//              Integer   iLength      文字の長さ
'//              string    chOne        埋める文字(省略時は半角スペース)
'//
'//              型        値        意味
'//  戻り値    : string    左寄せされた文字列
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
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
'//  関数名称  : PrintHeader
'//  機能名称  : ヘッダ部作成
'//  機能概要  : ヘッダ部を作成する。（ジャーナルの１～４行目)
'//
'//              型        名称         意味
'//  引数      : Integer   iFileNum     ファイル番号
'//              string    strTitle     ジャーナルタイトル
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(EG30 V32.1.0.1 2016-06-14   REVISED BY [TCC] T.Nakajima
'//                 2016年度施策対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub PrintHeader(iFileNum As Integer, strTitle As String)
    Dim lpSystemTime            As SYSTEMTIME               'ローカル時刻を取得
    'EG30 V32.1.0.1 DEL START
    'Print #iFileNum, "*************EG20*************"
    'EG30 V32.1.0.1 DEL END
    Print #iFileNum, strTitle
    
    'ローカル時刻を取得
    Call GetLocalTime(lpSystemTime)
    Print #iFileNum, "印字日時：" & lpSystemTime.wYear & "年" & Format(lpSystemTime.wMonth, "00") & "月" & Format(lpSystemTime.wDay, "00") & "日" _
                            & Format(lpSystemTime.wHour, "00") & ":" & Format(lpSystemTime.wMinute, "00")
    Print #iFileNum, ""
End Sub
'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : PrintHeader2
'//  機能名称  : ヘッダ部作成
'//  機能概要  : ヘッダ部を作成する。（ジャーナルの１～４行目)
'//
'//              型        名称         意味
'//  引数      : Integer   iFileNum     ファイル番号
'//              string    strTitle     ジャーナルタイトル
'//              string    strTitle2    ジャーナルタイトル２行名
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG20 V30.3.0.1) 2014-10-01   CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi07_003_01】
'//     REVISIONS :(EG30 V32.1.0.1 2016-06-14   REVISED BY [TCC] T.Nakajima
'//                 2016年度施策対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub PrintHeader2(iFileNum As Integer, strTitle As String, strTitle2 As String)
    Dim lpSystemTime            As SYSTEMTIME               'ローカル時刻を取得
    
    'EG30 V32.1.0.1 DEL START
    'Print #iFileNum, "*************EG20*************"
    'EG30 V32.1.0.1 DEL END
    Print #iFileNum, strTitle
    Print #iFileNum, strTitle2
    
    'ローカル時刻を取得
    Call GetLocalTime(lpSystemTime)
    Print #iFileNum, "印字日時：" & lpSystemTime.wYear & "年" & Format(lpSystemTime.wMonth, "00") & "月" & Format(lpSystemTime.wDay, "00") & "日" _
                            & Format(lpSystemTime.wHour, "00") & ":" & Format(lpSystemTime.wMinute, "00")
    Print #iFileNum, "保存日時：" & pfGetSaveDate(0)    'コーナ0の保存日時  'EG30 V32.1.0.1 ADD
    Print #iFileNum, ""
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  関数名称  : PrintHeader3
'//  機能名称  : ヘッダ部作成
'//  機能概要  : ヘッダ部を作成する。
'//
'//              型        名称         意味
'//  引数      : Integer   iFileNum     ファイル番号
'//              string    strTitle     ジャーナルタイトル
'//              string    strSaveDate  保存日時
'//
'//              型        値        意味
'//  戻り値    : なし
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-22   CODED   BY [TCC] T.Nakajima
'//                 2016年度施策対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub PrintHeader3(iFileNum As Integer, strTitle As String, strSaveDate As String)
    Dim lpSystemTime            As SYSTEMTIME               'ローカル時刻を取得
    
    Print #iFileNum, strTitle
    'ローカル時刻を取得
    Call GetLocalTime(lpSystemTime)
    Print #iFileNum, "印字日時：" & lpSystemTime.wYear & "年" & Format(lpSystemTime.wMonth, "00") & "月" & Format(lpSystemTime.wDay, "00") & "日" _
                            & Format(lpSystemTime.wHour, "00") & ":" & Format(lpSystemTime.wMinute, "00")
    Print #iFileNum, "保存日時：" & strSaveDate
    Print #iFileNum, ""
End Sub


'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : pfCornerGokiCheck
'//  機能名称  : コーナ号機チェック
'//  機能概要  : 画面でチェックされたコーナ号機が存在するか確認する
'//
'//              型        名称         意味
'//  引数      : Integer   iCorner      コーナ(1～6)
'//              Integer   iGoki        号機（省略可能：省略時は号機はチェックしない) 1～16
'//
'//              型        値           意味
'//  戻り値    : Boolean   true/false   true:設置されている   false:設置されていない
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfCornerGokiCheck(iCornerNo As Integer, Optional iGoki As Integer = 0) As Boolean
    Dim i       As Integer
    Dim bRet    As Boolean
    bRet = False
    
    If gudtSettiCorner(iCornerNo - 1).intGokiNum > 0 Then
        ' 指定したコーナは設置されている
        ' パラメータで指定された号機は設置されているか？
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
        '指定されたコーナは設置されていない
    End If
    
    pfCornerGokiCheck = bRet
    Exit Function
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2016 All Rights Reserved
'//
'//  関数名称  : pfCornerGokiToGateNo
'//  機能名称  : コーナ号機→論理号機番号に変換
'//  機能概要  : 画面でチェックされたコーナ号機が存在するか確認し、論理号機番号を返す。
'//
'//              型        名称         意味
'//  引数      : Integer   iCorner      コーナ(1～6)
'//              Integer   iGoki        号機（省略可能：省略時は号機はチェックしない) 1～16
'//              Integer   iGateNo      論理号機(1～32)
'//
'//              型        値           意味
'//  戻り値    : Boolean   true/false   true:設置されている   false:設置されていない
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-28   CODED   BY [TCC] T.Nakajima
'//                 2016年度施策対応
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfCornerGokiToGateNo(iCornerNo As Integer, iGoki As Integer, ByRef iGateNo As Integer) As Boolean
    Dim i       As Integer
    Dim bRet    As Boolean
    bRet = False
    iGateNo = 0
    If gudtSettiCorner(iCornerNo - 1).intGokiNum > 0 Then
        ' 指定したコーナは設置されている
        ' パラメータで指定された号機は設置されているか？
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
        '指定されたコーナは設置されていない
    End If
    
    pfCornerGokiToGateNo = bRet
    Exit Function
End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2013 All Rights Reserved
'//
'//  関数名称  : pfSettingaCheck
'//  機能名称  : コーナ号機の設置確認
'//  機能概要  : ジャーナルに出力するコーナ号機が設置されているか確認する。
'//
'//              型        名称         意味
'//  引数      : Boolean   bGokiCheck   号機チェック有無
'//
'//
'//              型        値           意味
'//  戻り値    : Boolean   true/false   true:設置されている   false:設置されていない
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfSettingCheck(Optional bGokiCheck As Boolean = True) As Boolean
    Dim i   As Integer
    Dim j   As Integer
    Dim k   As Integer
    
    '画面で設定されたうちどれかひとつでも設置されているコーナ号機があればOKとする
    For i = 0 To udtJprPrintSetteingInfo.iCornerCount - 1
        If gudtSettiCorner(udtJprPrintSetteingInfo.iCorner(i) - 1).intGokiNum > 0 Then
            'そのコーナは設置されている
            ' チェックされた号機はそのコーナに存在しているか？(号機チェックありの場合)
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
'//  関数名称  : MidByte
'//  機能名称  : コーナ号機の設置確認
'//  機能概要  : ジャーナルに出力するコーナ号機が設置されているか確認する。
'//
'//              型        名称         意味
'//  引数      : String    strTarget     対象文字列
'//              long      iStart       開始位置(1バイト～)
'//              Variant   ibyteCount   長さ
'//
'//
'//              型        値           意味
'//  戻り値    :String                  抽出された文字列
'//
'//     ORIGINAL  :(EG20 V7.2.0.1) 2013-06-26   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
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
'//  関数名称  : psMakeTukaRiyoImageFile
'//  機能名称  : 通過データ/利用金額データジャーナルのイメージファイル作成（在来用）
'//  機能概要  : 通過データおよび利用金額データジャーナルのイメージファイルを作成する。
'//
'//              型        名称         意味
'//  引数      : Integer   iCornerIdx   コーナインデックス
'//              Long      dwDataKind   データ種別（通過データ、利用金額データ）
'//
'//
'//              型        値           意味
'//  戻り値    : 無し
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：      幹線コーナ用のイメージファイル作成処理が別途必要となったため、
'//              JprEdit_TukaData()からサブルーチン化
'///////////////////////////////////////////////////////////////////
Private Sub psMakeTukaRiyoImageFile(iCornerIdx As Integer, dwDataKind As Long)
    
    Dim strBaitaiFileName   As String                       '媒体出力ファイル TUKAコーナ名YYYYMMDDhhmmss.csv ICRIYOコーナ名YYYYMMDDhhmmss.csv
    Dim ReadFileBaitai()    As BAITAI_OUTPUT_IMAGE_FILE     '媒体出力ファイル
    Dim intOutFile          As Integer                      '出力ファイル番号
    Dim strLineCount()      As String
    Dim fso                 As New FileSystemObject
    Dim FsoTS               As TextStream
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    Dim l                   As Integer
    Dim strCammaArray()     As String                       'カンマ区切りで1項目ずつ取り出したデータ
    Dim iKomokuMaxCnt       As Integer                      ' 集計データ項目の最大数
    Dim iStartLineKaisatu   As Integer                      ' 改札側データの開始行（ＣＳＶファイルの）
    Dim iStartLineShusatu   As Integer                      ' 集札側データの開始行（ＣＳＶファイルの）
            
    On Error GoTo Err_handler
    
    If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then       '通過データ
        
        iKomokuMaxCnt = MAX_KOMOKU_NUM_TUKA
        iStartLineKaisatu = 6   '改札側の明細は元ファイル(CSV)配列の(6)から
        iStartLineShusatu = 60  '集札側の明細は元ファイル(CSV)配列の(60)から
        
        strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "TUKA" & gstrStationName(iCornerIdx) & gstrCornerName(iCornerIdx) & "*.csv")
    ElseIf dwDataKind = Ml_DT_SHU_KIND.ML_DT_KINGAKU_BAITAI Then    '利用金額データ
        
        iKomokuMaxCnt = MAX_KOMOKU_NUM_KINGAKU
        iStartLineKaisatu = 6   '改札側の明細は元ファイル(CSV)配列の(6)から
        iStartLineShusatu = 25  '集札側の明細は元ファイル(CSV)配列の(60)から
        
        strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "ICRIYO" & gstrStationName(iCornerIdx) & gstrCornerName(iCornerIdx) & "*.csv")
    End If
           
    '////////////////////////////////////////////////
    '// 通過データ/利用金額の媒体出力ファイルを取得
    'ファイル番号取得
    '駅名称＋コーナ名称yyyymmddhhmmss.csv
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForAppending)
    j = FsoTS.Line
    FsoTS.Close
           
    ReDim strLineCount(j) As String         'CSVファイルを1行ずつ入れておく
           
    i = 0
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForReading)
    Do Until FsoTS.AtEndOfStream = True
        strLineCount(i) = FsoTS.ReadLine
        i = i + 1
    Loop
    FsoTS.Close
    Set fso = Nothing
    
    '媒体出力ファイルイメージ構造体にセットする
    ReDim ReadFileBaitai(j) As BAITAI_OUTPUT_IMAGE_FILE         'ファイル読込用エリア
    l = 0
    If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
    
        For i = 0 To j - 1
            Select Case i
                Case 0, 1, 2, 3, 4, 57, 58    'TUKAxxxx.csvの1～4行目まではタイトルなので、項目名にセット
                    ReadFileBaitai(i).strKomokuName = strLineCount(i)
                Case Else
                    'カンマ区切りを1項目ずつ取り出す。
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
                Case 0, 1, 2, 3, 4, 22, 23    'ICRIYOxxxx.csvの1～4行目まではタイトルなので、項目名にセット
                    ReadFileBaitai(i).strKomokuName = strLineCount(i)
                Case Else
                    'カンマ区切りを1項目ずつ取り出す。
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

    Print #intJprFile, "設置コーナ：" & gstrCornerName(iCornerIdx)
    Print #intJprFile, ""
    
    If dwDataKind = Ml_DT_SHU_KIND.ML_DT_TUKA_BAITAI Then
        Print #intJprFile, "【通過データ】"
    Else
        Print #intJprFile, "【ＩＣカード利用金額データ】"
    End If
    '/////////////////////
    '改札側データの出力
    Print #intJprFile, "改札側通過合計"
    
    For i = 0 To MAX_KOMOKU_NUM_TUKA - 1
        If RTrim(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName) = "0" Then
            '項目名に0がセットされていたら以降は出力しない
            Exit For
        Else
            '締切オフラインジャーナルとあわせるための例外処理
            If RTrim(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName) = "その他IC (小)" Then
                ReadFileBaitai(i + iStartLineKaisatu).strKomokuName = "その他IC(小)" & Space(38)   'スペースを除く
            End If
        
            Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLineKaisatu).strKomokuName, vbFromUnicode), 20), vbUnicode) _
            & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLineKaisatu).strGoukei), "#,0"), 10)
        End If
    Next i
    Print #intJprFile, ""
    
    '/////////////////////
    '集札側データの出力
    Print #intJprFile, "集札側通過合計"
    
    For i = 0 To MAX_KOMOKU_NUM_TUKA - 1
        If RTrim(ReadFileBaitai(i + iStartLineShusatu).strKomokuName) = "0" Then
            '項目名に0がセットされていたら以降は出力しない
            Exit For
        Else
            '締切オフラインジャーナルとあわせるための例外処理
            If RTrim(ReadFileBaitai(i + iStartLineShusatu).strKomokuName) = "その他IC (小)" Then
                ReadFileBaitai(i + iStartLineShusatu).strKomokuName = "その他IC(小)" & Space(38)    'スペースを除く
            End If
        
            Print #intJprFile, StrConv(LeftB(StrConv(ReadFileBaitai(i + iStartLineShusatu).strKomokuName, vbFromUnicode), 20), vbUnicode) _
            & PadLeft(Format(CLng(ReadFileBaitai(i + iStartLineShusatu).strGoukei), "#,0"), 10)
        End If
    Next i
    Print #intJprFile, ""
        
    
    'Print #intJprFile, FOOTER_STRING
    'Close #intJprFile
    
    Exit Sub
    
'エラー処理
Err_handler:

    If intOutFile > 0 Then
        Close #intOutFile
    End If
    
    If intJprFile > 0 Then
        Close #intJprFile
    End If

    Set fso = Nothing
    'エラーログの出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : psMakeTukaImageFileKan
'//  機能名称  : 通過データジャーナルのイメージファイル作成（幹線用）
'//  機能概要  : 通過データジャーナルのイメージファイルを作成する。
'//
'//              型        名称         意味
'//  引数      : Integer   iCornerIdx   コーナインデックス
'//
'//
'//              型        値           意味
'//  戻り値    : 無し
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psMakeTukaImageFileKan(iCornerIdx As Integer)
    
    Dim strBaitaiFileName   As String                       '媒体出力ファイル TUKAコーナ名YYYYMMDDhhmmss.csv ICRIYOコーナ名YYYYMMDDhhmmss.csv
    Dim ReadFileBaitai()    As BAITAI_OUTPUT_IMAGE_FILE_KAN '媒体出力ファイル
    Dim intOutFile          As Integer                      '出力ファイル番号
    Dim strLineCount()      As String
    Dim fso                 As New FileSystemObject
    Dim FsoTS               As TextStream
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    Dim l                   As Integer
    Dim strCammaArray()     As String                       'カンマ区切りで1項目ずつ取り出したデータ
    Dim iKomokuMaxCnt       As Integer                      ' 集計データ項目の最大数
    Dim iStartLine          As Integer                      '各集計ブロックの開始行
                                                                
    On Error GoTo Err_handler
    
    '各集計項目の出力開始位置を取得（INIファイルにより出力有無が指定できるため、開始位置は可変になる）
    
    strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "TUKA" & gstrStationName(iCornerIdx) & gstrCornerName(iCornerIdx) & "*.csv")
           
    '////////////////////////////////////////////////
    '// 通過データ/利用金額の媒体出力ファイルを取得
    'ファイル番号取得
    '駅名称＋コーナ名称yyyymmddhhmmss.csv
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForAppending)
    j = FsoTS.Line
    FsoTS.Close
           
    ReDim strLineCount(j) As String         'CSVファイルを1行ずつ入れておく
           
    i = 0
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForReading)
    Do Until FsoTS.AtEndOfStream = True
        strLineCount(i) = FsoTS.ReadLine
        i = i + 1
    Loop
    FsoTS.Close
    Set fso = Nothing
    
    '媒体出力ファイルイメージ構造体にセットする
    ReDim ReadFileBaitai(j) As BAITAI_OUTPUT_IMAGE_FILE_KAN     'ファイル読込用エリア
    l = 0
    
    For i = 0 To j - 1
        If InStr(strLineCount(i), ",") = 0 Then
            'カンマ区切りになっていない行は項目名にとりあえずデータをセット
            ReadFileBaitai(i).strKomokuName = strLineCount(i)
        Else
            'カンマ区切りを1項目ずつ取り出す。
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

    Print #intJprFile, "設置コーナ：" & gstrCornerName(iCornerIdx)
    Print #intJprFile, ""
    
    Print #intJprFile, "【ＪＲ東新幹線通過データ】"
    
    '//////////////////////////////////////////////////////////
    '改札側 新幹線通過データの出力
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_KAISATU_KANSEN_TUKA)
        Print #intJprFile, "改札側　新幹線通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_TUKA_KAN - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '集札側　新幹線通過データの出力
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_SHUSATU_KANSEN_TUKA)
        
        Print #intJprFile, "集札側　新幹線通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_TUKA_KAN - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '運行不能通過データの出力
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_UNKOU_FUNOU) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_IC_UNKO_FUNOU)
    
        Print #intJprFile, "運行不能通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_UNKOU_FUNOU - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '幹―在乗換　在来線通過データの出力
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA_KAN_ZAI) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_KAN_ZAI_TUKA)

        Print #intJprFile, "幹－在乗換　在来線通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_NORIKAE_TUKA - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '在―幹乗換　在来線通過データの出力
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA_ZAI_KAN) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_ZAI_KAN_TUKA)
        
        Print #intJprFile, "在－幹乗換　在来線通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_NORIKAE_TUKA - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '自駅入場救済データの出力
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_KYUSAI) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_JIEKI_KYUSAI)
    
        Print #intJprFile, "自駅入場救済通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_JIEKI_KYUSAI - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '磁気券回収中止通過データの出力
    '//////////////////////////////////////////////////////////
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_KAISHU_CHUSI) = CMN_ON Then
        iStartLine = pfGetStartLineTuka(mintTukaShukeiKoumoku.SHUKEI_KAISHU_CHUSHI)
    
        Print #intJprFile, "磁気券回収中止通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_JIKI_KAISHU_CHUSHI - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    
'エラー処理
Err_handler:

    If intOutFile > 0 Then
        Close #intOutFile
    End If
    
    If intJprFile > 0 Then
        Close #intJprFile
    End If

    Set fso = Nothing
    'エラーログの出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : psMakeRiyoImageFileKan
'//  機能名称  : 利用金額データジャーナルのイメージファイル作成（幹線用）
'//  機能概要  : 利用金額データジャーナルのイメージファイルを作成する。
'//
'//              型        名称         意味
'//  引数      : Integer   iCornerIdx   コーナインデックス
'//
'//
'//              型        値           意味
'//  戻り値    : 無し
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Sub psMakeRiyoImageFileKan(iCornerIdx As Integer)
    
    Dim strBaitaiFileName   As String                       '媒体出力ファイル TUKAコーナ名YYYYMMDDhhmmss.csv ICRIYOコーナ名YYYYMMDDhhmmss.csv
    Dim ReadFileBaitai()    As BAITAI_OUTPUT_IMAGE_FILE_KAN '媒体出力ファイル
    Dim intOutFile          As Integer                      '出力ファイル番号
    Dim strLineCount()      As String
    Dim fso                 As New FileSystemObject
    Dim FsoTS               As TextStream
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    Dim l                   As Integer
    Dim strCammaArray()     As String                       'カンマ区切りで1項目ずつ取り出したデータ
    Dim iKomokuMaxCnt       As Integer                      ' 集計データ項目の最大数
    Dim iStartLine          As Integer                      '各集計ブロックの開始行
                                                                
    On Error GoTo Err_handler
    
    '各集計項目の出力開始位置を取得（INIファイルにより出力有無が指定できるため、開始位置は可変になる）
    
    strBaitaiFileName = PATH_SHUKEI_SEND & Dir(PATH_SHUKEI_SEND & "ICRIYO" & gstrStationName(iCornerIdx) & gstrCornerName(iCornerIdx) & "*.csv")
           
    '////////////////////////////////////////////////
    '// 通過データ/利用金額の媒体出力ファイルを取得
    'ファイル番号取得
    '駅名称＋コーナ名称yyyymmddhhmmss.csv
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForAppending)
    j = FsoTS.Line
    FsoTS.Close
           
    ReDim strLineCount(j) As String         'CSVファイルを1行ずつ入れておく
           
    i = 0
    Set FsoTS = fso.OpenTextFile(strBaitaiFileName, ForReading)
    Do Until FsoTS.AtEndOfStream = True
        strLineCount(i) = FsoTS.ReadLine
        i = i + 1
    Loop
    FsoTS.Close
    Set fso = Nothing
    
    '媒体出力ファイルイメージ構造体にセットする
    ReDim ReadFileBaitai(j) As BAITAI_OUTPUT_IMAGE_FILE_KAN     'ファイル読込用エリア
    l = 0
    
    For i = 0 To j - 1
        If InStr(strLineCount(i), ",") = 0 Then
            'カンマ区切りになっていない行は項目名にとりあえずデータをセット
            ReadFileBaitai(i).strKomokuName = strLineCount(i)
        Else
            'カンマ区切りを1項目ずつ取り出す。
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

    Print #intJprFile, "設置コーナ：" & gstrCornerName(iCornerIdx)
    Print #intJprFile, ""
    
    Print #intJprFile, "【ＪＲ東新幹線金額データ】"
    
    '//////////////////////////////////////////////////////////
    '改札側 大人 新幹線スイカ通過合計の出力
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAI_OTONA_SUICA_RIYO)
        Print #intJprFile, "改札側 大人 新幹線ｽｲｶ通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '集札側 大人 新幹線スイカ通過合計の出力
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_SHU_OTONA_SUICA_RIYO)
        Print #intJprFile, "集札側 大人 新幹線ｽｲｶ通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '改札側 小児 新幹線スイカ通過合計の出力
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAI_SHONI_SUICA_RIYO)
        Print #intJprFile, "改札側 小児 新幹線ｽｲｶ通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '集札側 小児 新幹線スイカ通過合計の出力
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_SHU_SHONI_SUICA_RIYO)
        Print #intJprFile, "集札側 小児 新幹線ｽｲｶ通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    'スイカ会社間精算運賃支払い通過合計
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_CARD_SHIHARAI) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_SEISAN_SHIHARAI)
        Print #intJprFile, "ｽｲｶ会社間精算運賃支払通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_SEISAN - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '改札側オートチャージ通過合計
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_AUTO_CHARGE) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAI_AUTOCHARGE)
        Print #intJprFile, "改札側 オートチャージ通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_AUTOCHARGE - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '集札側オートチャージ通過合計
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_AUTO_CHARGE) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_SHU_AUTOCHARGE)
        Print #intJprFile, "集札側 オートチャージ通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_AUTOCHARGE - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '新幹線運賃 大人　スイカ通過合計
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_KIKAN_KANSEN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAN_OTONA_SUICA_RIYO)
        Print #intJprFile, "幹線運賃 大人 ｽｲｶ通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '新幹線運賃 小児　スイカ通過合計
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_KIKAN_KANSEN) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_KAN_SHONI_SUICA_RIYO)
        Print #intJprFile, "幹線運賃 小児 ｽｲｶ通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '乗換在来運賃 大人　スイカ通過合計
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_KIKAN_ZAIRAI) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_NORI_ZAI_OTONA_SUICA_RIYO)
        Print #intJprFile, "乗換在来運賃 大人 ｽｲｶ通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    '乗換在来運賃 小児　スイカ通過合計
    '//////////////////////////////////////////////////////////
    'INIで出力有に設定されていれば出力する
    If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_KIKAN_ZAIRAI) = CMN_ON Then
        iStartLine = pfGetStartLineKingaku(mintKingakuShukeiKoumoku.SHUKEI_NORI_ZAI_SHONI_SUICA_RIYO)
        Print #intJprFile, "乗換在来運賃 小児 ｽｲｶ通過合計"
        
        For i = 0 To MAX_KOMOKU_NUM_SUICA_RIYO - 1
            If RTrim(ReadFileBaitai(i + iStartLine).strKomokuName) = "" Then
                '項目名に0がセットされていたら出力しない
            Else
                '項目名が20桁に収まらない場合は半角スペース一つ入れて数値を出力（位置はそろえない)
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
    
'エラー処理
Err_handler:

    If intOutFile > 0 Then
        Close #intOutFile
    End If
    
    If intJprFile > 0 Then
        Close #intJprFile
    End If

    Set fso = Nothing
    'エラーログの出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_FILE, JPR_PRINT_OUTPUT_ERR, 0)
    
End Sub

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : pfGetStartLineTuka
'//  機能名称  : 指定した集計項目の印字開始位置取得
'//  機能概要  : 指定した集計項目の印字開始位置をGAIBU_OUTPUT.INIに従って求める。
'//
'//              型        名称         意味
'//  引数      : Integer   intShukeiKoumoku     集計項目
'//
'//
'//              型        値           意味
'//  戻り値    : Integer                開始位置
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetStartLineTuka(intShukeiKoumoku As Integer) As Integer

    Dim intCount            As Integer
    Dim intIniIdx           As Integer  'GAIBU_OUTPUT.INIのキーに対するインデックス
    
    Dim intStartLine        As Integer  '通過データの開始行数（CSV上）
    
    Dim intNextBlockLine    As Integer  '次の集計ブロックのデータがある位置（CSV上）
    
    Dim intNowLine           As Integer  'INIファイルの出力有無に従って、CSVファイルを上から順に見ていったときの現在行
    
    intNowLine = 0
    intNextBlockLine = 6
    intIniIdx = 0
    
    For intCount = 0 To MAX_TUKA_SHUKEI_KOUMOKU - 1
    
        Select Case intCount
            Case mintTukaShukeiKoumoku.SHUKEI_KAISATU_KANSEN_TUKA              '【改札側 新幹線通過データ】
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_TUKA_KAN + 2
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_SHUSATU_KANSEN_TUKA              '【集札側　新幹線通過データ】
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_TUKA_KAN + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_IC_UNKO_FUNOU                    '【運行不能データ】
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_UNKOU_FUNOU) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_UNKOU_FUNOU + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_KAN_ZAI_TUKA                    '【幹-在 乗換通過データ】
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA_KAN_ZAI) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_NORIKAE_TUKA + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_ZAI_KAN_TUKA                    '【在-幹 乗換通過データ】
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_TUKA_ZAI_KAN) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_NORIKAE_TUKA + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_JIEKI_KYUSAI                    '【自駅入場救済通過データ】
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_KYUSAI) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_JIEKI_KYUSAI + 3
                End If
            Case mintTukaShukeiKoumoku.SHUKEI_KAISHU_CHUSHI                  '【磁気券回収中止通過データ】
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_KAISHU_CHUSI) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_JIKI_KAISHU_CHUSHI + 3
                End If
            Case Else   '上記以外は金額データに関する設定のためスキップ
                
        End Select
        If intCount <> 0 Then
            intIniIdx = intIniIdx + 1
        End If
        
        '求めたい開始位置だったら、その行数を返す
        If intShukeiKoumoku = intCount Then
            pfGetStartLineTuka = intNowLine
            Exit Function
        End If
    Next
    
    '上記のFor文を最大回数まで回って終了したということは、求めたい開始位置が求められなかった。
    pfGetStartLineTuka = intNowLine

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : pfGetStartLineKingaku
'//  機能名称  : 指定した集計項目の印字開始位置取得
'//  機能概要  : 指定した集計項目の印字開始位置をGAIBU_OUTPUT.INIに従って求める。
'//
'//              型        名称         意味
'//  引数      : Integer   intShukeiKoumoku     集計項目
'//
'//
'//              型        値           意味
'//  戻り値    : Integer                開始位置
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetStartLineKingaku(intShukeiKoumoku As Integer) As Integer

    Dim intCount            As Integer
    Dim intIniIdx           As Integer  'GAIBU_OUTPUT.INIのキーに対するインデックス
    
    Dim intStartLine        As Integer  '通過データの開始行数（CSV上）
    
    Dim intNextBlockLine    As Integer  '次の集計ブロックのデータがある位置（CSV上）
    
    Dim intNowLine           As Integer  'INIファイルの出力有無に従って、CSVファイルを上から順に見ていったときの現在行
    
    intNowLine = 0
    intNextBlockLine = 6
    intIniIdx = 0
    
    For intCount = 0 To MAX_KINGAKU_SHUKEI_KOUMOKU - 1
    
        Select Case intCount
            '【改札側　大人　新幹線スイカ利用合計金額】
            '【集札側　大人　新幹線スイカ利用合計金額】
            '【改札側　小児　新幹線スイカ利用合計金額】
            
            '【幹線運賃　大人　スイカ利用合計金額】
            '【乗換在来運賃　大人　スイカ利用合計金額】
            Case mintKingakuShukeiKoumoku.SHUKEI_KAI_OTONA_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_SHU_OTONA_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_KAI_SHONI_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_KAN_OTONA_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_NORI_ZAI_OTONA_SUICA_RIYO
                
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_SUICA_RIYO + 2
                End If
            '【集札側　小児　新幹線スイカ利用合計金額】
            '【幹線運賃　小児　スイカ利用合計金額】
            '【乗換在来運賃　小児　スイカ利用合計金額】
            Case mintKingakuShukeiKoumoku.SHUKEI_SHU_SHONI_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_KAN_SHONI_SUICA_RIYO, _
                 mintKingakuShukeiKoumoku.SHUKEI_NORI_ZAI_SHONI_SUICA_RIYO
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_ICSF_KIKAN) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_SUICA_RIYO + 3
                End If
            '【スイカ会社間精算データ　運賃支払額】
            Case mintKingakuShukeiKoumoku.SHUKEI_SEISAN_SHIHARAI
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_IC_CARD_SHIHARAI) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_SUICA_SEISAN + 3
                End If
            '【改札側　オートチャージデータ】
            Case mintKingakuShukeiKoumoku.SHUKEI_KAI_AUTOCHARGE
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_AUTO_CHARGE) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_AUTOCHARGE + 2
                End If
            '【集札側　オートチャージデータ】
            Case mintKingakuShukeiKoumoku.SHUKEI_SHU_AUTOCHARGE
                If gintShukeiOutFlg(mintGaibuOutputKey.GAIBU_INI_AUTO_CHARGE) = CMN_ON Then
                    intNowLine = intNextBlockLine
                    intNextBlockLine = intNowLine + MAX_KOMOKU_NUM_AUTOCHARGE + 3
                End If
            Case Else   '上記以外は金額データに関する設定のためスキップ
                
        End Select
        If intCount <> 0 Then
            intIniIdx = intIniIdx + 1
        End If
        
        '求めたい開始位置だったら、その行数を返す
        If intShukeiKoumoku = intCount Then
            pfGetStartLineKingaku = intNowLine
            Exit Function
        End If
    Next
    
    '上記のFor文を最大回数まで回って終了したということは、求めたい開始位置が求められなかった。
    pfGetStartLineKingaku = intNowLine

End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : pfGetSubGateCsv
'//  機能名称  : 自改補助情報取得
'//  機能概要  : 指定したコーナの自改補助CSVファイルを取得する。
'//
'//              型        名称         意味
'//  引数      : Integer   intCornerNo   コーナ番号
'//
'//
'//              型        値           意味
'//  戻り値    : Integer                取得レコード数
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(EG20 V30.3.0.1) 2014-09-19   CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi07_008_01】
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Function pfGetSubGateCsv(intCornerNo As Integer) As Integer                                            ' EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 DEL
Private Function pfGetSubGateCsv(intCornerNo As Integer, intGokiNo As Integer, intKomoku As Integer) As Integer 'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD

    Dim intFileNumber            As Integer
    Dim i                        As Integer
    Dim ReadBuf                  As JIKAIINFO_IMAGE_FILE    '読み込みバッファ
        
    'Erase ReadSetteiSubGate        'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 DEL
    
    'エラールーチンを宣言
    On Error GoTo Err_handler      'EG20 V30.3.0.1 ADD
    
    'ファイル番号取得
    intFileNumber = FreeFile
    
    'CSVファイルオープン
    Open EKI_TUDO_CHK_SUBGATE_FILE For Input As #intFileNumber
    
    '一致するコーナ番号のレコードをエリアに保存していく
    i = 0
    Do While Not EOF(intFileNumber)
                
        Input #intFileNumber, ReadBuf.strBunrui_Dai, ReadBuf.strBunrui_Tyu, _
            ReadBuf.srtBunrui_Sho, ReadBuf.strCorner, ReadBuf.strKomoku, _
            ReadBuf.strKubun, ReadBuf.strData, ReadBuf.strSetShosai
        
        If CInt(ReadBuf.strCorner) = intCornerNo Then
            If CInt(ReadBuf.strBunrui_Tyu) = intGokiNo Then     'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD
                If CInt(ReadBuf.srtBunrui_Sho) = intKomoku Then     'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD
                    'ReDim Preserve ReadSetteiSubGate(i) As JIKAIINFO_IMAGE_FILE    'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 DEL
                    'ReadSetteiSubGate(i) = ReadBuf                                 'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD
                    ReadSetteiSubGate((intGokiNo - 1) * SUBGATE_ITEM_NUM + (intKomoku - 1)) = ReadBuf
                    i = i + 1
                    Exit Do     '号機、項目番号で絞り込むようにしたので、戻り値となるレコード数は０か１どちらかになる。 EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD
                End If      'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD
            End If      'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD
        End If
    Loop
    
    'CSVファイルクローズ
    Close #intFileNumber
    pfGetSubGateCsv = i
    
'EG20 V30.3.0.1 ADD START
    Exit Function
Err_handler:
    If intFileNumber > 0 Then
        Close #intFileNumber
    End If
    '異常ログ出力
    Call sLogTraceReq(LTYP_ERROR, L3AN_ETC, JPR_PRINT_OUTPUT_ERR, 0)
    
    pfGetSubGateCsv = 0
'EG20 V30.3.0.1 ADD END


End Function

'///////////////////////////////////////////////////////////////////
'//  (C) Copyright TOSHIBA Corporation 2014 All Rights Reserved
'//
'//  関数名称  : pfOutPutSubGate
'//  機能名称  : 自改補助情報出力
'//  機能概要  : 指定したコーナの自改補助内容をジャーナル形式で出力する
'//
'//              型        名称         意味
'//  引数      : Integer   intCornerNo   コーナ番号
'//              Integer   intFileNumber ファイル番号
'//
'//
'//              型        値           意味
'//  戻り値    : Integer                取得レコード数
'//
'//  ORIGINAL  :(EG20 V30.1.0.1) 2014-03-28   CODED   BY [TCC] T.Nakajima
'//  REVISIONS :(EG20 V30.3.0.1) 2014-09-19   CODED   BY [TCC] T.Nakajima
'//                 北陸新幹線フェーズ２対応 【HKRK_Kansi07_003_01】、【HKRK_Kansi07_008_01】
'//  REVISIONS :(EG30 V32.1.0.1) 2016-06-16   CODED   BY [TCC] T.Nakajima
'//                 2016年度施策対応
'//  REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
'Private Sub pfOutPutSubGate(intCornerNo As Integer, intFileNumber As Integer)  'EG20 V30.3.0.1 DEL
Private Function pfOutPutSubGate(intCornerNo As Integer, intFileNumber As Integer) As Boolean   'EG20 V30.3.0.1 ADD
    Dim intTitleFlg             As Integer                  '自改補助の大見出しの出力フラグ
    Dim intSubGateCnt           As Integer                  '自改補助1コーナ分のレコード数
    Dim i                       As Integer
    Dim intGokiLoop             As Integer                  '号機1～32 EG20 V30.3.0.1       【HKRK_Kansi07_008_01】 ADD
    Dim intKomokuLoop           As Integer                  '小項目①～⑥ EG20 V30.3.0.1    【HKRK_Kansi07_008_01】 ADD
    Dim intRet                  As Integer                  ' EG20 V30.3.0.1 ADD
    
    'EG30 V32.1.0.1 ADD START
    Dim bRet                    As Boolean
    Dim lErrCode                As Long
    Dim strEkiSettiBefPath      As String           '現在駅設定データ（変更前保存）
    Dim strGetValue             As String * 64
    Dim strCompValue            As String           '設定値（変更前保存）
    Dim strChangeFlg            As String           '変更印
    Dim intValueLen             As Integer          '取得した設定値の長さ
    'EG30 V32.1.0.1 ADD END


    'そのコーナの自改補助データを取得
    intTitleFlg = 0
    'intSubGateCnt = pfGetSubGateCsv(intCornerNo)    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】DEL
    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD START
    'SUB_GATE_KAN.INIからコーナがなくなったため、コーナは0固定で号機、項目①～⑥の順でEKI_DATA.CSVから検索
    intSubGateCnt = 0                       'EG20 V30.3.0.1 【HKRK_Kansi07_008_01】 ADD
    For intGokiLoop = 0 To 31
        For intKomokuLoop = 0 To 5
            intRet = pfGetSubGateCsv(0, intGokiLoop + 1, intKomokuLoop + 1)
            If intRet = 0 Then
                ' CSVからの取得件数が0件の場合はエラーとする。
                pfOutPutSubGate = False
                Exit Function
            Else
                intSubGateCnt = intSubGateCnt + intRet
            End If
        Next
    Next
    
    'EG30 V32.1.0.1 ADD START
    'コーナ０の変更前保存された駅都度データと比較する。
    'そのコーナの変更前データ保存されたデータをメモリ上に展開する
    strEkiSettiBefPath = Replace(EKI_SETTI_FILE_BEF, "#", 0)
    Call dllGetEkiIniDataBefore(strEkiSettiBefPath, lErrCode)
    'EG30 V32.1.0.1 ADD END
    
    'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD END
    For i = 0 To intSubGateCnt - 1
        ' EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL START
        ' 指定したコーナ、号機に対応するレコードを出力する必要がなくなり、1～32号機固定になったためIf文を削除
        'If IsTaisyoGoki(CInt(ReadSetteiSubGate(i).strCorner), CInt(ReadSetteiSubGate(i).strBunrui_Tyu)) = True Then
        ' EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL END
        If intTitleFlg = 0 Then
            Print #intFileNumber, ""
            'Print #intFileNumber, "【改札機　設置条件　自社】統合" 'EG30 V32.1.0.1 DEL
            Print #intFileNumber, "　【改札機　設置条件　自社】統合"    'EG30 V32.1.0.1 ADD
            intTitleFlg = 1
        End If
        
        'EG30 V32.1.0.1 ADD START
        '変更前データ保存された設定値と比較する
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
        
        'ReadSetteiSubGate(i).strData = Format(ReadSetteiSubGate(i).strData, "0#")      'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 DEL
        ReadSetteiSubGate(i).strData = Format(ReadSetteiSubGate(i).strData, "00#")      'EG20 V30.3.0.1 【HKRK_Kansi07_003_01】 ADD
        Select Case CInt(ReadSetteiSubGate(i).srtBunrui_Sho)
            Case 1
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "号機 " & "FM券 ｺｰﾅｰ番号" & " " & ReadSetteiSubGate(i).strData  'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "号機 " & "FM券 ｺｰﾅｰ番号" & " " & ReadSetteiSubGate(i).strData    'EG30 V32.1.0.1 ADD
            Case 2
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "号機 " & "FM券 号機番号" & " " & ReadSetteiSubGate(i).strData  'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "号機 " & "FM券 号機番号" & " " & ReadSetteiSubGate(i).strData    'EG30 V32.1.0.1 ADD
            Case 3
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "号機 " & "新幹線IC ｺｰﾅｰ番号" & " " & ReadSetteiSubGate(i).strData  'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "号機 " & "新幹線IC ｺｰﾅｰ番号" & " " & ReadSetteiSubGate(i).strData    'EG30 V32.1.0.1 ADD
            Case 4
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "号機 " & "新幹線IC 号機番号" & " " & ReadSetteiSubGate(i).strData  'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "号機 " & "新幹線IC 号機番号" & " " & ReadSetteiSubGate(i).strData    'EG30 V32.1.0.1 ADD
            Case 5
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "号機 " & "NRZ券 ｺｰﾅｰ番号" & " " & ReadSetteiSubGate(i).strData 'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "号機 " & "NRZ券 ｺｰﾅｰ番号" & " " & ReadSetteiSubGate(i).strData   'EG30 V32.1.0.1 ADD
            Case 6
                'Print #intFileNumber, ReadSetteiSubGate(i).strBunrui_Tyu & "号機 " & "NRZ券 号機番号" & " " & ReadSetteiSubGate(i).strData 'EG30 V32.1.0.1 DEL
                Print #intFileNumber, strChangeFlg & ReadSetteiSubGate(i).strBunrui_Tyu & "号機 " & "NRZ券 号機番号" & " " & ReadSetteiSubGate(i).strData   'EG30 V32.1.0.1 ADD
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
'//  関数名称  : pfGetSaveDate
'//  機能名称  : 変更前データ保存日付取得処理
'//  機能概要  : コーナごとに保存されているSaveDate.datの更新日付を取得する
'//
'//              型        名称      意味
'//  引数      : Integer    intCorner   取得するコーナ番号
'//
'//              型        値        意味
'//  戻り値    : String     更新日付    YYYY年MM月DD日HH:MM
'//
'//     ORIGINAL  :(EG30 V32.1.0.1) 2016-06-14   CODED   BY [TCC] T.Nakajima
'//     REVISIONS :(X.X.X.X) ----------   REVISED BY []
'//  備考：
'///////////////////////////////////////////////////////////////////
Private Function pfGetSaveDate(intCorner As Integer) As String
    Dim strFileName(0 To 1)     As String           '作成日時
    Dim intCnt                  As Integer          'カウンタ
    Dim lngHandle               As Long             'ハンドル

    Dim lpCreatTime             As FILETIME         '作成日時
    Dim lpAccessTime            As FILETIME         '最終アクセス日時
    Dim lpLastwTime             As FILETIME         '更新日時
    Dim lpLocalTime             As FILETIME         'ローカル日時
    Dim lpSystemTime            As SYSTEMTIME       'システム時刻
    Dim bRet                    As Boolean          '戻り値
    
    Dim strSaveFile             As String
    
    On Error Resume Next

           
    '保存ファイルの日付を取得
    strSaveFile = PATH_OPERATE & "CORNER" & CStr(intCorner) & "\\SETTEI_BEF\\" & SET_BEF_DATE_FILE
    If Dir(strSaveFile) = "" Then
        pfGetSaveDate = "    年  月  日  :  "
        Exit Function
    Else
        'ファイルをオープン
        lngHandle = CreateFile(strSaveFile, GENERIC_READ, FILE_SHARE_READ, _
                                    0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)

        'ファイルオープンが正常に行われたか？
        If lngHandle = INVALID_HANDLE_VALUE Then GoTo ErrorHandler
            'ファイルタイムをGET
            bRet = GetFileTime(lngHandle, lpCreatTime, lpAccessTime, lpLastwTime)
            If bRet = False Then GoTo APIError                          '取得が正常に行われたか？
        
            'ファイルタイムをローカルタイムに変換
            bRet = FileTimeToLocalFileTime(lpLastwTime, lpLocalTime)    'EG20 V2.1.0.1 ADD 【Mainte_03_01】
            If bRet = False Then GoTo APIError                          '変換が正常に行われたか？
        
            'ローカルタイムをシステムタイムに変換
            bRet = FileTimeToSystemTime(lpLocalTime, lpSystemTime)
            If bRet = False Then GoTo APIError                          '変換が正常に行われたか？
                
            'ハンドルのクローズ
            Call CloseHandle(lngHandle)
        
            '作成日付を表示する (YYYY年MM月DD日hh:mm)
            pfGetSaveDate = lpSystemTime.wYear & "年" & _
                                Format(lpSystemTime.wMonth, "00") & "月" & _
                                Format(lpSystemTime.wDay, "00") & "日" & _
                                Format(lpSystemTime.wHour, "00") & ":" & _
                                Format(lpSystemTime.wMinute, "00")
    End If
            
    Exit Function

APIError:
    Call CloseHandle(lngHandle)             'ハンドルのクローズ

ErrorHandler:
    pfGetSaveDate = "    年  月  日  :  "
    
End Function
'EG30 V32.1.0.1 ADD END
