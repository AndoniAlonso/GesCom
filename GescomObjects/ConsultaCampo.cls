VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsultaCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As ConsultaCampoProps

Public Property Get ConsultaCampoID() As String

    ConsultaCampoID = mudtprops.ConsultaCampoID
  
End Property

Public Property Get Consulta() As String

    Consulta = mudtprops.Consulta
  
End Property

Public Property Get NombreCampo() As String

    NombreCampo = mudtprops.NombreCampo
  
End Property

Public Property Get Alias() As String

    Alias = mudtprops.Alias
  
End Property

Public Property Get TipoCampo() As String

    TipoCampo = mudtprops.TipoCampo
  
End Property

' Devuelve cierto si el tipo de campo es una fecha (para utilizar despues en un picker de fechas).
Public Property Get IsDate() As Boolean

    IsDate = (mudtprops.TipoCampo = "D")
  
End Property

' Devuelve cierto si el tipo de campo es un numerico (para utilizar despues en comparaciones).
Public Property Get IsNumeric() As Boolean

    IsNumeric = (mudtprops.TipoCampo = "N")
  
End Property

Public Sub Load(ConsultaID As Long)

    Dim objPersist As ConsultaCampoPersist
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.ConsultaCampoPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(ConsultaID)
    Set objPersist = Nothing

End Sub

Private Sub SetState(Buffer As String)

    Dim udtData As ConsultaData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
