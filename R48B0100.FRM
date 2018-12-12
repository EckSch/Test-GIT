VERSION 5.00
Object = "{473956E6-78C8-4D4F-BDA1-9F35065342E6}#2.1#0"; "R48Checkbox.ocx"
Object = "{DAEC4B0D-F7F9-4873-A2E2-7006412CE831}#2.0#0"; "mbaToolbar.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6219FA60-20A6-41DD-B7F5-BE33795101DA}#1.0#0"; "R48TBox.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{33130918-9A80-4A1E-A420-06F4DCBA163B}#1.0#0"; "dataButton.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Betriebsabteilungen"
   ClientHeight    =   7035
   ClientLeft      =   4005
   ClientTop       =   2895
   ClientWidth     =   7500
   HelpContextID   =   50100000
   Icon            =   "R48B0100.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   7035
   ScaleWidth      =   7500
   Begin dataButton.DataBtn cmdOK 
      Height          =   255
      Left            =   3180
      TabIndex        =   28
      Top             =   660
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   32768
      Text            =   "myData"
   End
   Begin VB.Frame Frame6 
      Caption         =   "Planstellenzuordnung zu Betriebsabteilungen"
      Height          =   4815
      Left            =   2460
      TabIndex        =   20
      Top             =   1800
      Width           =   4935
      Begin R48CheckBox.R48ChkBox SSCheck1 
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   344
         Caption         =   "nur zugeordnete Planstellen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Default         =   0   'False
      End
      Begin R48CheckBox.R48ChkBox SSCheck9 
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   540
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   344
         Caption         =   "alle nicht zugeordneten Planstellen markieren"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Default         =   0   'False
      End
      Begin TrueDBGrid80.TDBGrid grdPS 
         Height          =   3780
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   6668
         _LayoutType     =   4
         _RowHeight      =   13
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Planstelle"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Betriebsabteilung"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).ScrollBars=   2
         Splits(0).AllowColSelect=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=556"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=476"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3731"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3651"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
         BorderStyle     =   0
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&HC0C0C0&"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HC0C0C0&"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H808080&"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H80&"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HBDFDFC&"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Named:id=33:Normal"
         _StyleDefs(43)  =   ":id=33,.parent=0"
         _StyleDefs(44)  =   "Named:id=34:Heading"
         _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(46)  =   ":id=34,.wraptext=-1"
         _StyleDefs(47)  =   "Named:id=35:Footing"
         _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   "Named:id=36:Selected"
         _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=39:EvenRow"
         _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(57)  =   "Named:id=40:OddRow"
         _StyleDefs(58)  =   ":id=40,.parent=33"
         _StyleDefs(59)  =   "Named:id=41:RecordSelector"
         _StyleDefs(60)  =   ":id=41,.parent=34"
         _StyleDefs(61)  =   "Named:id=42:FilterBar"
         _StyleDefs(62)  =   ":id=42,.parent=33"
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Unten ausrichten
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   6660
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "17.08.2017"
         EndProperty
      EndProperty
   End
   Begin VB.Timer StatusTimer 
      Interval        =   2000
      Left            =   5460
      Top             =   420
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Oben ausrichten
      Height          =   435
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7500
      _Version        =   65536
      _ExtentX        =   13229
      _ExtentY        =   767
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin mbaToolbar.mbaBtnToolbar cmdToolbarLast 
         Height          =   315
         Left            =   3420
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Letzter Datensatz"
         Top             =   60
         Width           =   315
         _ExtentX        =   609
         _ExtentY        =   609
         Picture         =   "R48B0100.frx":058A
         BackColor       =   16777215
         Object.Width           =   345
         Object.Height          =   345
      End
      Begin mbaToolbar.mbaBtnToolbar cmdToolbarFirst 
         Height          =   315
         Left            =   2280
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Erster Datensatz"
         Top             =   60
         Width           =   315
         _ExtentX        =   609
         _ExtentY        =   609
         Picture         =   "R48B0100.frx":066C
         BackColor       =   16777215
         Object.Width           =   345
         Object.Height          =   345
      End
      Begin mbaToolbar.mbaBtnToolbar cmdToolbarNew 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Neuer Datensatz"
         Top             =   60
         Width           =   315
         _ExtentX        =   609
         _ExtentY        =   609
         Picture         =   "R48B0100.frx":074E
         BackColor       =   16777215
         Object.Width           =   345
         Object.Height          =   345
      End
      Begin mbaToolbar.mbaBtnToolbar cmdToolbarNext 
         Height          =   315
         HelpContextID   =   30000006
         Index           =   1
         Left            =   3060
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Nächster Datensatz"
         Top             =   60
         Width           =   315
         _ExtentX        =   609
         _ExtentY        =   609
         Picture         =   "R48B0100.frx":0870
         BackColor       =   16777215
         Object.Width           =   345
         Object.Height          =   345
      End
      Begin mbaToolbar.mbaBtnToolbar cmdToolbarPrev 
         Height          =   315
         HelpContextID   =   30000006
         Index           =   0
         Left            =   2640
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Vorheriger Datensatz"
         Top             =   60
         Width           =   315
         _ExtentX        =   609
         _ExtentY        =   609
         Picture         =   "R48B0100.frx":0952
         BackColor       =   16777215
         Object.Width           =   345
         Object.Height          =   345
      End
      Begin mbaToolbar.mbaBtnToolbar cmdToolbarPrint 
         Height          =   315
         HelpContextID   =   30000005
         Index           =   3
         Left            =   1200
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Drucken"
         Top             =   60
         Width           =   315
         _ExtentX        =   609
         _ExtentY        =   609
         Picture         =   "R48B0100.frx":0A34
         BackColor       =   16777215
         Object.Width           =   345
         Object.Height          =   345
      End
      Begin mbaToolbar.mbaBtnToolbar cmdToolbarDelete 
         Height          =   315
         HelpContextID   =   30000004
         Index           =   2
         Left            =   840
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Datensatz löschen"
         Top             =   60
         Width           =   315
         _ExtentX        =   609
         _ExtentY        =   609
         Picture         =   "R48B0100.frx":112E
         BackColor       =   16777215
         Object.Width           =   345
         Object.Height          =   345
      End
      Begin mbaToolbar.mbaBtnToolbar cmdToolbarSave 
         Height          =   315
         HelpContextID   =   30000003
         Index           =   1
         Left            =   480
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Datensatz speichern"
         Top             =   60
         Width           =   315
         _ExtentX        =   609
         _ExtentY        =   609
         Picture         =   "R48B0100.frx":1210
         BackColor       =   16777215
         Object.Width           =   345
         Object.Height          =   345
      End
      Begin MSComDlg.CommonDialog dlg1 
         Left            =   8220
         Top             =   -60
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageCombo imcVerknuepfung 
         Height          =   375
         Left            =   4740
         TabIndex        =   23
         Top             =   30
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "imcVerknuepfung"
         ImageList       =   "imlVerknuepfung"
      End
   End
   Begin R48TBox.Text Text1 
      Height          =   255
      Index           =   3
      Left            =   1740
      TabIndex        =   8
      Top             =   1680
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      TextRechts      =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   3
      TextNumerisch   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      MaxLength       =   3
   End
   Begin R48TBox.Text Text1 
      Height          =   255
      HelpContextID   =   50100000
      Index           =   2
      Left            =   1740
      TabIndex        =   7
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   3
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      MaxLength       =   3
   End
   Begin R48TBox.Text Text1 
      Height          =   255
      HelpContextID   =   50100000
      Index           =   1
      Left            =   1740
      TabIndex        =   5
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   30
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      MaxLength       =   30
   End
   Begin R48TBox.Text Text1 
      Height          =   255
      HelpContextID   =   50100000
      Index           =   0
      Left            =   1740
      TabIndex        =   4
      Top             =   660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   450
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   8
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      MaxLength       =   8
   End
   Begin Threed.SSCommand cmdMC 
      Height          =   225
      HelpContextID   =   30000009
      Left            =   2850
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Matchcode"
      Top             =   690
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "R48B0100.frx":190A
   End
   Begin MSComctlLib.ImageList imlVerknuepfung 
      Left            =   6540
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin R48TBox.Text Text1 
      Height          =   255
      HelpContextID   =   50100000
      Index           =   4
      Left            =   1740
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   8
      TextNumerisch   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      MaxLength       =   8
   End
   Begin Threed.SSCommand cmdAbtLtr 
      Height          =   225
      HelpContextID   =   30000009
      Left            =   2940
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Matchcode"
      Top             =   1200
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "R48B0100.frx":19EC
   End
   Begin VB.Label lblAbtLtr 
      Caption         =   "Abteilungsleiter"
      Height          =   225
      Left            =   3240
      TabIndex        =   27
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Abteilungsleiter"
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   26
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Max. abwes."
      Height          =   225
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Fibu-Mandant"
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Bezeichnung"
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Betriebsabteilung"
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   660
      Width           =   1275
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Beenden"
      HelpContextID   =   50100000
      Index           =   0
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "B&earbeiten"
      Begin VB.Menu mnucopy 
         Caption         =   "Feld &kopieren"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Feld &einfügen"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSepAddFav 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddFav 
         Caption         =   "Favoriten hinzufügen"
      End
      Begin VB.Menu mnuDelFav 
         Caption         =   "Favoriten entfernen"
      End
      Begin VB.Menu mnuSepAddFav1 
         Caption         =   "-"
      End
      Begin VB.Menu mnucopystammsatz 
         Caption         =   "Betriebsabteilung kopieren"
      End
   End
   Begin VB.Menu mnuMain1 
      Caption         =   "&Hilfe"
      HelpContextID   =   50100000
      Begin VB.Menu mnuMainHelp 
         Caption         =   "&Inhalt"
         HelpContextID   =   50100000
         Index           =   0
      End
      Begin VB.Menu mnuMainHelp 
         Caption         =   "I&nfo..."
         HelpContextID   =   50100000
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ABT As cBetriebsabteilungen
'Private MC_ABT As Object ' cR48MCTAB
Private sBetrAbt As String

Private mbStatusFixed As Boolean

Dim nRandLinks%
Dim nRandRechts%
Dim nRandOben%
Dim nRandUnten%
Dim Abbruch As Boolean
Dim index%
Dim SZ%
Dim arr As New XArrayDB

Private Sub cmdAbtLtr_Click()
    Dim s$, v As Variant
    Dim MC As Object 'cR48MCTAB
    Set MC = CreateObject("R48MCArtikel.cR48MCArtikel")
    v = MC.execute(ConstMC.MC_Personalstamm): If Not IsNull(v) Then s = v(0)
    Set MC = Nothing
    Form1.Refresh
    If (s <> "") Then
        Text1(4).Text = s
        readAbteilungsleiter
    End If
    
    Screen.MousePointer = 0
End Sub

Private Sub cmdMC_Click()
    Dim s$
    Dim MC_ABT As Object 'cR48MCTAB
    Set MC_ABT = CreateObject("R48MCTABellen.cR48MCTAB")
    s = MC_ABT.MC_Betriebsabteilungen(False)
    Form1.Refresh
    If (s <> "") Then
        Text1(0).Text = s
        readABT
    End If
    
    Screen.MousePointer = 0
End Sub

Private Sub cmdok_Click()
    If Text1(0).Text = "" Then
        Text1(0).SetFocus
        Exit Sub
    End If
    readABT
    
End Sub
Private Function readAbteilungsleiter() As Boolean
    If Val(Text1(4).Text) = 0 Then
        readAbteilungsleiter = True
        lblAbtLtr.Caption = ""
    Else
        Dim rs As New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "select rtrim(a.Name1) +', ' + rtrim(a.Name2) as AbtLeiter from personalstamm as p inner join adressen as a on a.AdressID=p.AdressId where p.PersonalNr=" & Val(Text1(4).Text), conn_main, adOpenForwardOnly
        If Not rs.EOF Then
            lblAbtLtr.Caption = rs!AbtLeiter
            readAbteilungsleiter = True
        Else
            lblAbtLtr.Caption = ""
            readAbteilungsleiter = False
        End If
    End If
End Function
Private Sub readABT()
    If (0 = ABT.execute(BGETEQUAL, Text1(0))) Then
        recAbtEnter
        gnES = ES_EDIT
    ElseIf 4 Then
        gnES = ES_NEW
        gbES_Changed = True
    Else
        gnES = ES_NONE
    End If
    ES_set
End Sub
Private Function RecABTLeave() As Boolean
    Select Case ES_Leave
    Case vbYes
        RecABTLeave = ABTSave
    Case vbNo
        RecABTLeave = True
    Case vbCancel
        RecABTLeave = False
    End Select
End Function

Private Sub recAbtEnter()
    ES_BeginUpdate
    
    Text1(0).Text = ABT.Abteilung
    Text1(1).Text = ABT.Bezeichnung          'ABTBez
    Text1(2).Text = ABT.FIBUMandant          'ABTFibumandant
    Text1(3).Text = ABT.MaxAbwMitarbeiter    'MaxAbwMitarb
    Text1(4).Text = IIf(ABT.Abteilungsleiter = 0, "", ABT.Abteilungsleiter)
    ReadPlanSt False
    CheckDoppelte
    ES_Enter
    ES_set
    SSCheck1.value = False
    SSCheck9.value = False
End Sub

'01.11.2010 sch auf doppelte Verwendung prüfen
Private Sub CheckDoppelte()
    Dim rs As New ADODB.Recordset
    Dim i%
    For i = 0 To arr.UpperBound(1)
        Dim txt$
        txt = ""
        If arr(i, 0) = True Then
            rs.Open "select a.bezeichnung from betriebsabteilungen a join betriebsabteilung_planstellen p on p.abteilung=a.abteilung join planstellen ps on ps.planstelle=p.planstelle where p.abteilung <> '" & Text1(0) & "' and ps.bezeichnung='" & arr(i, 1) & "'", conn_main
            If Not rs.EOF Then
                txt = txt + " " + rs!Bezeichnung
                rs.MoveNext
            End If
            rs.Close
            If txt > "" Then msgbox me,"Die Planstelle " & arr(i, 1) & " ist auch noch folgenden Abteilungen zugeordnet: " & txt & vbCrLf & "Bitte löschen sie alle doppelten Zuordnungen!"
        End If
    Next
End Sub

Private Sub cmdToolbarDelete_Click(index As Integer)
    ABTDelete
End Sub

Private Sub cmdToolbarFirst_Click()
    ABTFirst False
End Sub

Private Sub cmdToolbarLast_Click()
    ABTFirst True
End Sub

Private Sub cmdToolbarNew_Click()
    If RecABTLeave Then
        gnES = ES_NONE
        ES_set
        
        Text1(0).Text = ""
        Text1(0).SetFocus
    End If
End Sub

Private Sub cmdToolbarNext_Click(index As Integer)
    If RecABTLeave Then
        ABTNext True
        Text1(0).SetFocus
    End If
End Sub

Private Sub cmdToolbarPrev_Click(index As Integer)
    If RecABTLeave Then
        ABTNext False
        Text1(0).SetFocus
    End If
End Sub

Private Sub cmdToolbarPrint_Click(index As Integer)
    ABTPrint
End Sub

Private Sub cmdToolbarSave_Click(index As Integer)
    ABTSave
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        Select Case KeyCode
            Case 70
                KeyCode = 0
                Matchcode
            Case 83
                KeyCode = 0
                ABTSave
        End Select
    End If
End Sub

Private Sub Form_Load()
If dvxInit Then
On Local Error Resume Next
    Set cmdToolbarLast.Picture = LoadPicture(Trim(DVXRT.USERPRGVERZEICHNIS) & "ICONS\button_light_gray_rechts.gif")
    Set cmdToolbarFirst.Picture = LoadPicture(Trim(DVXRT.USERPRGVERZEICHNIS) & "ICONS\button_light_gray_links.gif")
    Set cmdToolbarNew.Picture = LoadPicture(Trim(DVXRT.USERPRGVERZEICHNIS) & "ICONS\button_light_gray_new.gif")
    Set cmdToolbarNext(1).Picture = LoadPicture(Trim(DVXRT.USERPRGVERZEICHNIS) & "ICONS\button_light_gray_simplerechts.gif")
    Set cmdToolbarPrev(0).Picture = LoadPicture(Trim(DVXRT.USERPRGVERZEICHNIS) & "ICONS\button_light_gray_simplelinks.gif")
    Set cmdToolbarPrint(3).Picture = LoadPicture(Trim(DVXRT.USERPRGVERZEICHNIS) & "ICONS\button_light_gray_print.gif")
    Set cmdToolbarDelete(2).Picture = LoadPicture(Trim(DVXRT.USERPRGVERZEICHNIS) & "ICONS\button_light_gray_cancel.gif")
    Set cmdToolbarSave(1).Picture = LoadPicture(Trim(DVXRT.USERPRGVERZEICHNIS) & "ICONS\button_light_gray_save.gif")
On Local Error GoTo 0
    CenterForm Me
    Set ABT = New cBetriebsabteilungen
    ABT.Refresh
    Verknuepfungen_Show
    Load frmHidden ' Für Start Externer Programme aus dem MC heraus
    Unload frmStart
    
    gnES = ES_NONE
    ES_set
    Show
    grdPS.Columns(0).ValueItems.Presentation = dbgCheckBox
    grdPS.Array = arr
Else
    Unload Me
    Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (Not RecABTLeave) Then
        Cancel = 1
    Else
        ''' Start für Start Externer Programme aus dem MC heraus
            If gn_hwnd_MC > 0 Then
                On Local Error GoTo errorhandler
                SendMessageSTRING gn_hwnd_MC, &HC, 256, Me.Text1(0).Text
errorhandler:
                On Local Error GoTo 0
            End If
        ''' Ende für Start Externer Programme aus dem MC heraus
        Unload frmStart
        Set ABT = Nothing
        dvxTerminate
    End If
End Sub


Private Sub grdPS_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If arr(grdPS.Bookmark, 2) > "" And Trim(arr(grdPS.Bookmark, 2)) <> Trim(Text1(1)) Then
        msgbox me,"Eine Planstelle kann nur einer Abteilung zugeordnet werden", vbCritical
        Cancel = True
    End If
End Sub

Private Sub grdPS_LostFocus()
    grdPS.Update
End Sub

Private Sub imcVerknuepfung_Click()
    VerknuepfungShell
End Sub

Private Sub mnuAddFav_Click()
    AddFavoriten
End Sub

Private Sub mnucopystammsatz_Click()
 frmGrpCopy.Show 1
End Sub

Private Sub mnuDelFav_Click()
    DelFavoriten
End Sub

Private Sub mnuCopy_Click()
    SendKeys "^C"
End Sub

Private Sub mnuEdit_Click()
On Error GoTo 2
    If ActiveControl.SelText = "" Then
        mnuCopy.Enabled = False
    Else
        mnuCopy.Enabled = True
    End If
2:

On Error GoTo 3
    If Clipboard.GetFormat(vbCFText) Then
        mnuPaste.Enabled = True
    Else
        mnuPaste.Enabled = False
    End If
3:

End Sub

Private Sub MnuMain_Click(index As Integer)
    Unload Me
End Sub

Private Sub mnuMainHelp_Click(index As Integer)
    Select Case index
        Case 0
            dlg1.HelpFile = App.HelpFile
            dlg1.HelpCommand = cdlHelpContext
            dlg1.HelpContext = 50100000
            dlg1.ShowHelp
        Case 1
            frmInfo.Exec Me, App.EXEName, PRG_VERSION
    End Select
End Sub

Private Sub mnuPaste_Click()
    SendKeys "^V"
End Sub

Private Sub SSCheck1_Click(value As Integer)
    ReadPlanSt (value = True)
    If value Then SSCheck9.value = False
End Sub

Private Sub SSCheck1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys vbTab
End If

End Sub


Private Sub StatusTimer_Timer()
    StatusBar1.Panels(2).Text = ""
    StatusTimer.Enabled = False

End Sub

Private Sub SetState(s As String, Optional fixed As Variant)
    If (Not IsMissing(fixed)) Then
        mbStatusFixed = CBool(fixed)
    Else
        mbStatusFixed = False
    End If
    If (Not mbStatusFixed) Then StatusTimer.Enabled = True
    StatusBar1.Panels(2).Text = s
End Sub

Private Sub ClearState()
    StatusBar1.Panels(2).Text = ""
End Sub

Private Sub ES_set()
    Dim i%
    
    Static savES As Integer
    If (savES <> gnES) Then
        Select Case gnES

            Case ES_NONE
            mnucopystammsatz = False
                StatusBar1.Panels(1).Text = ""
                For i = 1 To 4
                    Text1(i).Enabled = False
                    Text1(i).Text = ""
                Next
                cmdAbtLtr.Enabled = False
                lblAbtLtr.Caption = ""
                arr.Clear
                arr.ReDim 0, -1, 0, 2
                grdPS.Enabled = False
                SSCheck1.Enabled = False
                SSCheck9.Enabled = False
             
            Case ES_EDIT, ES_NEW
            
                StatusBar1.Panels(1).Text = IIf(gnES = ES_EDIT, "Ändern", "Erfassen")
                Text1(1).Enabled = True
                Text1(2).Enabled = True
                Text1(3).Enabled = True
                Text1(4).Enabled = True
                cmdAbtLtr.Enabled = True
                grdPS.Enabled = True
                SSCheck1.Enabled = True
                SSCheck9.Enabled = True
                If gnES = ES_NEW Then ReadPlanSt False
                mnucopystammsatz = True
        End Select
        savES = gnES
    End If

End Sub

Private Sub ReadPlanSt(bAbt As Boolean)
    Dim rs As New ADODB.Recordset
    arr.Clear
    arr.ReDim 0, -1, 0, 2
    rs.Open "select s.bezeichnung as pbez,a.bezeichnung as abez,a.abteilung  from planstellen s left join betriebsabteilung_planstellen  as p  on s.planstelle=p.planstelle left join betriebsabteilungen a on a.abteilung=p.abteilung " & IIf(bAbt, "where p.abteilung ='" & Text1(0) & "'", ""), conn_main, adOpenDynamic, adLockOptimistic
    Do While Not rs.EOF
        arr.AppendRows
        arr(arr.UpperBound(1), 1) = Trim(rs!pbez)
        arr(arr.UpperBound(1), 2) = IIf(IsNull(rs!abez), "", Trim(rs!abez))
        If Trim(rs!Abteilung) = Trim(Text1(0)) Then arr(arr.UpperBound(1), 0) = True
        rs.MoveNext
    Loop
    rs.Close
    grdPS.Bookmark = 0
    grdPS.ReBind
End Sub

Private Sub SavePlanSt()
    Dim i%
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    rs.Open "delete from betriebsabteilung_planstellen where abteilung = '" & Text1(0) & "'", conn_main, adOpenDynamic, adLockOptimistic
    rs.Open "select * from betriebsabteilung_planstellen where abteilung = '" & Text1(0) & "'", conn_main, adOpenDynamic, adLockOptimistic
    For i = 0 To arr.UpperBound(1)
        If arr(i, 0) = True Then
            rs2.Open "select * from planstellen where bezeichnung = '" & arr(i, 1) & "'", conn_main, adOpenDynamic, adLockOptimistic
            If Not rs2.EOF Then
                rs.AddNew
                rs!Abteilung = Text1(0)
                rs!planstelle = rs2!planstelle
                rs.Update
            rs2.Close
            End If
        End If
    Next
    rs.Close
End Sub

Private Sub SSCheck9_Click(value As Integer)
    If value = True Then
        SSCheck1.value = False
        ReadPlanSt False
        Dim i%
        For i = 0 To arr.UpperBound(1)
            If arr(i, 2) = "" Then arr(i, 0) = True
        Next i
        grdPS.Refresh
    End If
End Sub



Private Function ABTSave() As Boolean
    Dim BOP%
    ABTSave = False
    Select Case gnES
        Case ES_EDIT
            BOP = BUPDATE
        Case ES_NEW
            BOP = BINSERT
            ABT.Abteilung = sBetrAbt
        Case Else
            Exit Function
    End Select
    ABT.Bezeichnung = Text1(1).Text
    ABT.FIBUMandant = Text1(2).Text
    If Text1(3).Text = "" Then Text1(3).Text = 0
    ABT.MaxAbwMitarbeiter = Text1(3).Text
    ABT.Abteilungsleiter = Val(Text1(4).Text)
    SavePlanSt
    If (ABT.execute(BOP) = 0) Then
        SetState "Daten gespeichert"
        ABTSave = True
        readABT
    End If

End Function

Private Sub ABTPrint()
    SZ = 0
'--- Druck-Dialog
    On Error GoTo lblErr
    dlg1.CancelError = True
    dlg1.Copies = 1
    dlg1.Flags = cdlPDUseDevModeCopies + cdlPDNoSelection
    dlg1.ShowPrinter
    
    Dim MC_ABT As Object 'cR48MCTAB
    Set MC_ABT = CreateObject("R48MCTABellen.cR48MCTAB")
    MC_ABT.print_Betriebsabteilungen
    
'---Einstellungen
    SetState "Drucken..."
    
lblErr:

End Sub

Private Sub ABTDelete()
    If (gnES = ES_EDIT) Then
        Select Case frmmsg.exec(me,App.Title, vbCrLf & vbCrLf & "Datensatz löschen ?", vbYesNo, nmsg_Kegel)
        'Select Case msgbox(me,"Datensatz löschen ?", vbYesNo + vbExclamation)
            Case vbYes
                If (sBetrAbt = ABT.Abteilung) Then
                    Dim rs As New ADODB.Recordset
                    rs.Open "delete from betriebsabteilung_planstellen where abteilung = '" & sBetrAbt & "'", conn_main, adOpenDynamic, adLockOptimistic
                    If (ABT.execute(BDELETE) = 0) Then
                        SetState "Datensatz gelöscht"
                        gnES = ES_NONE
                        ES_set
'                        readABT
                    End If
                Else
                    frmmsg.exec me,App.Title, vbCrLf & vbCrLf & "Interner Fehler. (ABTDelete: BetrAbt <> ABT.Abteilung)"
                    'msgbox me,"Interner Fehler. (ABTDelete: BetrAbt <> ABT.Abteilung)"
                End If
        End Select
    End If

End Sub

Private Sub Text1_Change(index As Integer)
    Select Case index
        Case 0
            If (gnES <> ES_UPDATE) Then
                If RecABTLeave Then
                    gnES = ES_NONE
                    ES_set
                Else
                    ES_BeginUpdate
                    Text1(0).Text = sBetrAbt
                    ES_EndUpdate
                End If
            End If
            sBetrAbt = Text1(0)
        Case 4: readAbteilungsleiter: ES_Changed
        Case Else
            gbES_Changed = True
    End Select
        
End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (index = 0) Then
            If Text1(0) = "" Then
                Text1(0).SetFocus
                Exit Sub
            Else
                readABT
            End If
        End If
        
    End If
End Sub

Private Function ABTNext(index As Boolean) As Boolean
    ABTNext = False
    If (index) Then
        Call ABT.execute(BGETGREATER, Text1(0).Text)
    Else
        Call ABT.execute(BGETLESS, Text1(0).Text)
    End If
    If (ABT.BStatus = 0) Then
        ABTNext = True
        gbES_ReadOnly = False
        recAbtEnter
        Exit Function
    End If
    If gnES = ES_EDIT Then readABT
    If (index) Then
        SetState "Tabellenende"
    Else
        SetState "Tabellenanfang"
    End If

End Function

Private Function ABTFirst(index As Boolean) As Boolean
    ABTFirst = False
    If (index) Then
        Call ABT.execute(BGETLAST)
    Else
        Call ABT.execute(BGETFIRST)
    End If
    If (ABT.BStatus = 0) Then
        ABTFirst = True
        gbES_ReadOnly = False
        recAbtEnter
    End If
    If gnES = ES_EDIT Then readABT
    If (index) Then
        SetState "Tabellenende"
    Else
        SetState "Tabellenanfang"
    End If

End Function

Private Sub Matchcode()
    '------------------------------------------------------------------
    'RHü, 18.07.2001:
    'Dieses MenuItem dient dazu, den ShortKey [STRG+F] zu verwalten.
    'Der Menüeintrag selbst kann ruhig unsichtbar bleiben. Wir suchen
    'uns jetzt einfach das Control, dass gerade beim Drücken dieser
    'Tastenkombination den Focus hat und können entsprechend reagieren!
    '------------------------------------------------------------------
    Dim sTargetControl As String
    Dim col%
    On Error Resume Next
    sTargetControl = Me.ActiveControl.Name & "(" & CStr(Me.ActiveControl.index) & ")"
    On Error GoTo 0
    
     
    Select Case sTargetControl
        Case "Text1(0)"
            cmdMC = True
    End Select
End Sub
Public Function Verknuepfung_Wert(sKey$) As String
    Dim nRow&
    Select Case sKey
        Case "Planstelle"
            Dim rs2 As New ADODB.Recordset
            rs2.Open "select * from planstellen where bezeichnung = '" & arr(grdPS.Bookmark, 1) & "'", conn_main
            If Not rs2.EOF Then
                Verknuepfung_Wert = rs2!planstelle
            Else
                Verknuepfung_Wert = "NEU"
            End If
            rs2.Close
        Case Else
            If sKey > "" Then
                frmmsg.exec me,App.Title, vbCrLf & vbCrLf & "Der Übergabewert " & sKey & " wird von disem Programm nicht unterstützt"
                'msgbox me,"Der Übergabewert " & sKey & " wird von disem Programm nicht unterstützt", vbCritical
            End If
    End Select
End Function



