VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Consultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolDisplay As Collection

Private Sub Class_Initialize()

    Set mcolDisplay = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolDisplay = Nothing
    
End Sub

Public Function Count() As Long

    Count = mcolDisplay.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolDisplay.[_NewEnum]

End Function

Public Function Item(ByVal Index As Variant) As ConsultaDisplay
Attribute Item.VB_UserMemId = 0

    Set Item = mcolDisplay(Index)
  
End Function

Public Sub Load(ByVal Objeto As String)
    Dim objPersist As ConsultasPersist
  
    Set objPersist = CreateObject("GCServerMTS.ConsultasPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(Objeto)
    Set objPersist = Nothing
  
End Sub

Private Sub SetState(Buffer As String)

    Dim objBuffer As Buffer
    Dim objDisplay As ConsultaDisplay
    Dim lngIndex As Long
    Dim udtData As ConsultaData
    Dim udtProps As ConsultaProps
  
    Set objBuffer = New Buffer
    With objBuffer
        .SetState Buffer
        For lngIndex = 1 To .Count
            Set objDisplay = New ConsultaDisplay
            udtData.Buffer = .Item(lngIndex)
            LSet udtProps = udtData
    
            With objDisplay
                .ConsultaID = udtProps.ConsultaID
                .Objeto = udtProps.Objeto
                .Nombre = udtProps.Nombre
                mcolDisplay.Add objDisplay
                Set objDisplay = Nothing
            End With
        Next
    End With

    Set objBuffer = Nothing

End Sub

