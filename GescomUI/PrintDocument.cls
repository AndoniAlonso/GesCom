VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolTablas As Collection

Public Sub Add(NombreTabla As String)
    Dim objTablaItem As PrintTabla
    
    Set objTablaItem = New PrintTabla
    mcolTablas.Add objTablaItem, NombreTabla

    Set objTablaItem = Nothing
End Sub

Private Sub Class_Initialize()

    Set mcolTablas = New Collection
    
End Sub

Private Sub Class_Terminate()

    Set mcolTablas = Nothing

End Sub

Public Function Item(ByVal NombreTabla As Variant) As PrintTabla
Attribute Item.VB_UserMemId = 0

    Set Item = mcolTablas(NombreTabla)

End Function

