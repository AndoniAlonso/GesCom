Attribute VB_Name = "GCmain"
Option Explicit

Private Const GesComSectionName = "GesCom"
Private strPersistServer As String
Public Function PERSIST_SERVER() As String
    
    If strPersistServer = vbNullString Then
        WinIniRegister GesComSectionName
        strPersistServer = WinGetString("PERSIST_SERVER", vbNullString)
        'PrivIniRead "GesCom", "PERSIST_SERVER", 0, vbNullString, strPersistServer, Null, "GesCom.INI"
    End If
    
    PERSIST_SERVER = strPersistServer
    
End Function

