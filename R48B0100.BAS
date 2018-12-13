Attribute VB_Name = "Mmain"
Option Explicit

Das ist branch Master2

Public Const FRM_CAPTION = "Betriebsabteilungen"
'16.12.2003 BSt Ladefortschritt um Art der Anwendung (hier "Pflege") erweitert
'               Tab- und Enterreihenfolge geändert
'28.05.2004 ESt Kompiliert wegen DLL-Änderung
'02.03.2006 ESt FIBU-Mandant in Maske auf 3 Stellen, wie in DB
'18.01.2007 ESt Enterführung und Speichern auf letztem Feld bei ENTER
'18.08.2009 Sch Enterführung korrigiert
'02.10.2009 ESt Neue Standardverknüpfung Betriebsabteilung - Planstelle
'08.06.2010 fs Msgboxen gegen frmMSG ausgetauscht
'01.11.2010 sch auf doppelte Verwendung prüfen
'06.05.2011 ESt Neu kompiliert wegen Fehler TDBG8Registrierung
'02.05.2014 jm  Updated buttons
'05.06.2014 ESt Austausch ssCommand -> mbaToolbar
'09.06.2016 ESt Abteilungsleiter
Public Const PRG_VERSION = "09.06.2016 ESt"

''' Start Für Start Externer Programme aus dem MC heraus
Public gn_hwnd_MC&
Public Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
''' Ende Für Start Externer Programme aus dem MC heraus

Sub Main()
    On Local Error Resume Next
    If (App.PrevInstance) Then
        WinActivatePrevInstance FRM_CAPTION
    Else
        frmStart.exec FRM_CAPTION & "-Pflege"
        App.Title = "DVx: " + FRM_CAPTION
        Load Form1
    End If
End Sub

Public Function dvxInit() As Boolean
    dvxInit = False
    Set DVXRT = CreateObject("DVXR48OLE.CDVXRT")
    Set conn_main = New ADODB.Connection
    
    If (DVXRT.Login(Path:=App.Path)) Then
        If DVXRT.RtTabellenBDE Or DVXRT.RtTabellenPPS Then
            App.HelpFile = DVXRT.GetHelpFile
            
            conn_main.ConnectionString = DVXRT.getConnectionString_Firma
            conn_main.Open
            
            dvxInit = True
        Else
            msgbox nothing,"Bediener ist nicht berechtigt", 48 + vbSystemModal
        End If
    End If
End Function

Public Sub dvxTerminate()
    If conn_main.State = adStateOpen Then conn_main.Close
    Set conn_main = Nothing
    Set DVXRT = Nothing
    End
End Sub


