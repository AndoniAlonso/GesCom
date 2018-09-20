VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const cnESCP2Port = "COM1"

Public Sub PrintTicket(ByVal Value As String)
'    Printer.Font.Name = "Courier New"
'    Printer.Print Value
'    Printer.EndDoc

    Open cnESCP2Port For Output As #1
    Print #1, Value
    DoEvents
    Close #1
    DoEvents

End Sub

Public Sub OpenDrawer(ByVal Pin2 As Boolean)
Dim strCommand As String

    
    If Pin2 Then
        strCommand = Chr$(&H1B) & Chr$(&H70) & Chr$(&H0) & Chr$(&H3C) & Chr$(&H78)
    Else
        strCommand = Chr$(&H1B) & Chr$(&H70) & Chr$(&H1) & Chr$(&H3C) & Chr$(&H78)
    End If
    
    Open cnESCP2Port For Output As #1
    Print #1, strCommand
    DoEvents
    Close #1
    DoEvents

End Sub
