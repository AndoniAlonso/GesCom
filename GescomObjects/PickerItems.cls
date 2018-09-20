VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PickerItems"
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

Public Function Item(ByVal Index As Variant) As PickerItemDisplay
Attribute Item.VB_UserMemId = 0

    Set Item = mcolDisplay(Index)
  
End Function

Public Sub Load(ByVal strTabla As String, ByVal Propietario As Long, _
    ByVal Empresa As Long, ByVal Temporada As Long)

    Dim objPersist As PickerItemsPersist
  
    Set objPersist = CreateObject("GCServerMTS.PickerItemsPersist", PERSIST_SERVER)
  
    ' Incluir aqui la llamada con la clausula where de SQL
    SetState objPersist.Fetch(strTabla, Propietario, Empresa, Temporada)
    Set objPersist = Nothing
      
End Sub

Private Sub SetState(Buffer As String)

    Dim objBuffer As Buffer
    Dim objDisplay As PickerItemDisplay
    Dim lngIndex As Long
    Dim udtData As PickerItemData
    Dim udtProps As PickerItemProps
      
    Set objBuffer = New Buffer
    With objBuffer
        .SetState Buffer
        For lngIndex = 1 To .Count
            Set objDisplay = New PickerItemDisplay
            udtData.Buffer = .Item(lngIndex)
            LSet udtProps = udtData

            With objDisplay
                .DocumentoID = udtProps.DocumentoID
                .Nombre = Trim(udtProps.Nombre)
                .Numero = udtProps.Numero
                .Descripcion = Trim(udtProps.Descripcion)
                .Cantidad = udtProps.Cantidad
                .Fecha = udtProps.Fecha
                mcolDisplay.Add objDisplay
                Set objDisplay = Nothing
            End With

        Next

    End With

    Set objBuffer = Nothing
    
End Sub

Public Function Add() As PickerItemDisplay

    Dim objPickerItemDisplay As PickerItemDisplay
  
    Set objPickerItemDisplay = New PickerItemDisplay
  
    Set Add = objPickerItemDisplay
    Set objPickerItemDisplay = Nothing

End Function

Public Sub AddPickerItemDisplay(Child As PickerItemDisplay)

    mcolDisplay.Add Item:=Child

End Sub
