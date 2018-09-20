VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturaVentaItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolFacturaVentaItems As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

    Set mcolFacturaVentaItems = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolFacturaVentaItems = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As FacturaVentaItem
Attribute Item.VB_UserMemId = 0

    Set Item = mcolFacturaVentaItems(Index)

End Function

Public Function Count() As Long

    Count = mcolFacturaVentaItems.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolFacturaVentaItems.[_NewEnum]

End Function

Public Sub BeginEdit()
    Dim objFacturaVentaItem As FacturaVentaItem
    
    If mflgEditing Then Err.Raise 445
    
    For Each objFacturaVentaItem In mcolFacturaVentaItems
      objFacturaVentaItem.ChildBeginEdit
    Next
    
    mflgEditing = True
    
End Sub

Public Sub CancelEdit()
    Dim objFacturaVentaItem As FacturaVentaItem
    Dim lngIndex As Long
    
    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    For lngIndex = mcolFacturaVentaItems.Count To 1 Step -1
        Set objFacturaVentaItem = mcolFacturaVentaItems.Item(lngIndex)
        With objFacturaVentaItem
            If .IsNew Then _
               mcolFacturaVentaItems.Remove lngIndex
            .ChildCancelEdit
        End With
        
    Next lngIndex

End Sub

Public Sub ApplyEdit(Comision As Double)
    
    If Not mflgEditing Then Err.Raise 445
    
    AsignarComision (Comision)
    
    mflgEditing = False
    
End Sub

Public Sub Delete()
    Dim objFacturaVentaItem As FacturaVentaItem
    
    If Not mflgEditing Then Err.Raise 445
    ' code to delete the object's data goes here
    
    For Each objFacturaVentaItem In mcolFacturaVentaItems
        With objFacturaVentaItem
            .BeginEdit
            .Delete
            .ApplyEdit
        End With
        
    Next
    
End Sub

Public Function Add() As FacturaVentaItem
    Dim objFacturaVentaItem As FacturaVentaItem
    
    If Not mflgEditing Then Err.Raise 445
    
    
    Set objFacturaVentaItem = New FacturaVentaItem
    objFacturaVentaItem.SetAsChild Me
    objFacturaVentaItem.ChildBeginEdit
    Set Add = objFacturaVentaItem
    Set objFacturaVentaItem = Nothing
    
End Function

Public Sub AddFacturaVentaItem(Child As FacturaVentaItem)

    mcolFacturaVentaItems.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445
    
    With mcolFacturaVentaItems(Index)
      .BeginEdit
      .Delete
      .ApplyEdit
    End With
    
End Sub

Friend Sub Load(FacturaVentaID As Long)
    Dim objPersist As FacturaVentaItemsPersist
    
    Set objPersist = CreateObject("GCServerMTS.FacturaVentaItemsPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(FacturaVentaID)
    Set objPersist = Nothing
    
End Sub

Public Property Get Bruto() As Double
    Dim objFacturaVentaItem As FacturaVentaItem
    
    Bruto = 0
    For Each objFacturaVentaItem In mcolFacturaVentaItems
        If Not objFacturaVentaItem.IsDeleted Then _
            Bruto = Bruto + objFacturaVentaItem.Bruto
    Next
    
End Property

Public Property Get Cantidad() As Double
    Dim objFacturaVentaItem As FacturaVentaItem
    
    Cantidad = 0
    For Each objFacturaVentaItem In mcolFacturaVentaItems
        If Not objFacturaVentaItem.IsDeleted Then _
            Cantidad = Cantidad + objFacturaVentaItem.Cantidad
    Next
    
End Property

' Existen Albaranes Contabilizados cuando hay algún Factura que tiene Albaranes.
Public Property Get HayAlbaran() As Boolean
    Dim objFacturaVentaItem As FacturaVentaItem
    
    HayAlbaran = False
    For Each objFacturaVentaItem In mcolFacturaVentaItems
        If objFacturaVentaItem.HayAlbaran Then
            HayAlbaran = True
            Exit For
        End If
    Next
    
End Property

Friend Function GetState() As String
    Dim objBuffer As Buffer
    Dim objFacturaVentaItem As FacturaVentaItem
    
    Set objBuffer = New Buffer
    Set objFacturaVentaItem = New FacturaVentaItem
    objBuffer.Initialize Len(objFacturaVentaItem.GetState), 20
    
    For Each objFacturaVentaItem In mcolFacturaVentaItems
        objBuffer.Add objFacturaVentaItem.GetState
    Next
    
    GetState = objBuffer.GetState
    Set objBuffer = Nothing
    
End Function

Friend Sub SetState(Buffer As String)
    Dim lngIndex As Long
    Dim objFacturaVentaItem As FacturaVentaItem
    Dim objBuffer As Buffer
    
    Set mcolFacturaVentaItems = Nothing
    Set mcolFacturaVentaItems = New Collection
    
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    
    For lngIndex = 1 To objBuffer.Count
        Set objFacturaVentaItem = New FacturaVentaItem
        With objFacturaVentaItem
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me
        End With
        mcolFacturaVentaItems.Add Item:=objFacturaVentaItem
    Next
    
End Sub

Public Sub Actualizar()
    Dim objFacturaVentaItem As FacturaVentaItem
    
    For Each objFacturaVentaItem In mcolFacturaVentaItems
        objFacturaVentaItem.Actualizar
    Next
End Sub

Public Sub CambiarComision(Comision As Double)
    Dim objFacturaVentaItem As FacturaVentaItem
    
    For Each objFacturaVentaItem In mcolFacturaVentaItems
        objFacturaVentaItem.BeginEdit
        objFacturaVentaItem.Comision = Comision
        objFacturaVentaItem.ApplyEdit
    Next
End Sub

' Asigna la comision a los items de Factura que son nuevos (que no lo tienen asignado).
Private Sub AsignarComision(Comision As Double)
    Dim objFacturaVentaItem As FacturaVentaItem
    
    For Each objFacturaVentaItem In mcolFacturaVentaItems
'        If objFacturaVentaItem.IsNew Then
         If objFacturaVentaItem.Comision <> Comision Then
            objFacturaVentaItem.BeginEdit
            objFacturaVentaItem.Comision = Comision
            objFacturaVentaItem.ApplyEdit
        End If
    Next
End Sub

' Devuelve el importe de comision de las lineas de factura
Public Property Get ImporteComision() As Double
    Dim objFacturaVentaItem As FacturaVentaItem
    
    ImporteComision = 0
    For Each objFacturaVentaItem In mcolFacturaVentaItems
        If Not objFacturaVentaItem.IsDeleted Then
            ImporteComision = ImporteComision + objFacturaVentaItem.ImporteComision
        End If
    Next
    
End Property

Friend Property Get IsDirty() As Boolean
    Dim objItem As FacturaVentaItem
    
    IsDirty = False
    
    For Each objItem In mcolFacturaVentaItems
        If objItem.IsDirty Or objItem.IsNew Then
            IsDirty = True
            Exit For
        End If
    Next
    
End Property

