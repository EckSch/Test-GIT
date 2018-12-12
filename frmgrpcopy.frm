VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{843D7F9A-ED85-4A40-9E35-C9D63E27D1F4}#1.1#0"; "mbaCommand.ocx"
Begin VB.Form frmGrpCopy 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Betriebsabteilung kopieren"
   ClientHeight    =   1560
   ClientLeft      =   4140
   ClientTop       =   3075
   ClientWidth     =   4440
   HelpContextID   =   50100000
   Icon            =   "frmgrpcopy.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Oben ausrichten
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4440
      _Version        =   65536
      _ExtentX        =   7832
      _ExtentY        =   767
      _StockProps     =   15
      BackColor       =   13160660
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
      Begin MSComDlg.CommonDialog dlg1 
         Left            =   5160
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin mbaCommand.Boton cmdCopy 
      Height          =   345
      Left            =   2910
      TabIndex        =   4
      Top             =   630
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      Caption         =   "Kopieren"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmgrpcopy.frx":058A
      BackColor       =   65280
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2655
      Begin VB.TextBox Text1 
         Appearance      =   0  '2D
         Height          =   285
         HelpContextID   =   50100000
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   2
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "neue Betriebsabteilung:"
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin mbaCommand.Boton cmdAbbruch 
      Height          =   345
      Left            =   2910
      TabIndex        =   5
      Top             =   1050
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      Caption         =   "Abbrechen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmgrpcopy.frx":0EF7
      BackColor       =   65280
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Beenden"
   End
   Begin VB.Menu mnuHilfe 
      Caption         =   "Hilfe"
      Begin VB.Menu mnuInhalt 
         Caption         =   "Inhalt"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "frmGrpCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdCopy_Click()
    ' solange das textfeld leer ist wird nichts kopiert
    If Trim(Text1.Text) = "" Then
        frmmsg.exec me,App.Title, vbCrLf & vbCrLf & "Geben Sie eine neue Betriebsabteilung ein."
        'msgbox me,"Geben Sie eine neue Betriebsabteilungen ein.", vbExclamation
        Exit Sub
    End If
    
    ' gibt es denn schon die verwendete Nummer?
    Dim rs As New ADODB.Recordset
    rs.Open "select abteilung from betriebsabteilung_planstellen where abteilung='" & Text1.Text & "'", conn_main, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        frmmsg.exec me,App.Title, vbCrLf & "Die gewählte Betriebsabteilungen existiert bereits. Kopieren nicht möglich", , nmsg_Durchfahrt_Verboten
        'msgbox me,"Die gewählte Betriebsabteilungen existiert bereits. Kopieren nicht möglich", vbExclamation
        rs.Close
        Exit Sub
    End If
    
    'Kopierfunktion
        gnES = ES_UPDATE
        Form1.Text1(0).Text = Text1.Text
         DoEvents
       Form1.ABT.Refresh
        Form1.ABT.execute BGETEQUAL, Text1.Text
        gnES = ES_NEW
        Form1.cmdToolbarSave(1).DoClick
        Unload Me
End Sub

Private Sub cmdAbbruch_Click()
    Unload Me
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCopy_Click
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuInhalt_Click()
    Form1.dlg1.HelpFile = App.HelpFile
    Form1.dlg1.HelpCommand = cdlHelpContext
    Form1.dlg1.HelpContext = 50100000
    Form1.dlg1.ShowHelp
End Sub

Private Sub mnuInfo_Click()
    frmInfo.exec Me, App.EXEName, PRG_VERSION
End Sub

Private Sub Form_Load()
    CenterForm Me
    On Local Error Resume Next
        Set cmdCopy.Picture = LoadPicture(Trim(DVXRT.USERPRGVERZEICHNIS) & "ICONS\button_Green.Gif")
        Set cmdAbbruch.Picture = LoadPicture(Trim(DVXRT.USERPRGVERZEICHNIS) & "ICONS\button_Red.Gif")
    On Local Error GoTo 0
End Sub
