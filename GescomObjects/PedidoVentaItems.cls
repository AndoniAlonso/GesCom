VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PedidoVentaItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMoneda As String

Private mcolPedidoVentaItems As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

    Set mcolPedidoVentaItems = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolPedidoVentaItems = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As PedidoVentaItem
Attribute Item.VB_UserMemId = 0

    Set Item = mcolPedidoVentaItems(Index)

End Function

Public Function Count() As Long

    Count = mcolPedidoVentaItems.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolPedidoVentaItems.[_NewEnum]

End Function

Public Sub BeginEdit(Moneda As String)
    Dim objPedidoVentaItem As PedidoVentaItem
    
    mMoneda = Moneda
    If mflgEditing Then Err.Raise 445
    
    For Each objPedidoVentaItem In mcolPedidoVentaItems
        objPedidoVentaItem.ChildBeginEdit Moneda
    Next
    
    mflgEditing = True
    
End Sub

Public Sub CancelEdit()
    Dim objPedidoVentaItem As PedidoVentaItem
    Dim lngIndex As Long
    
    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    For lngIndex = mcolPedidoVentaItems.Count To 1 Step -1
        Set objPedidoVentaItem = mcolPedidoVentaItems.Item(lngIndex)
        With objPedidoVentaItem
            If .IsNew Then _
                mcolPedidoVentaItems.Remove lngIndex
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
    Dim objPedidoVentaItem As PedidoVentaItem
    
    If Not mflgEditing Then Err.Raise 445
    ' code to delete the object's data goes here
    
    For Each objPedidoVentaItem In mcolPedidoVentaItems
        With objPedidoVentaItem
            .BeginEdit mMoneda
            .Delete
            .ApplyEdit
        End With
    
    Next

End Sub

Public Function Add() As PedidoVentaItem
    Dim objPedidoVentaItem As PedidoVentaItem
  
    If Not mflgEditing Then Err.Raise 445
    
    Set objPedidoVentaItem = New PedidoVentaItem
    objPedidoVentaItem.SetAsChild Me, mMoneda
    objPedidoVentaItem.ChildBeginEdit mMoneda
    Set Add = objPedidoVentaItem
    Set objPedidoVentaItem = Nothing

End Function

Friend Sub AddPedidoVentaItem(Child As PedidoVentaItem)

    mcolPedidoVentaItems.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445
    
    With mcolPedidoVentaItems(Index)
        .BeginEdit mMoneda
        .Delete
        .ApplyEdit
    End With
    
End Sub

Friend Sub Load(PedidoVentaID As Long, Moneda As String)
    Dim objPersist As PedidoVentaItemsPersist
    
    mMoneda = Moneda
    Set objPersist = CreateObject("GCServerMTS.PedidoVentaItemsPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(PedidoVentaID)
    Set objPersist = Nothing
    
End Sub
Public Property Get Bruto() As Double
    Dim objPedidoVentaItem As PedidoVentaItem
    
    Bruto = 0
    For Each objPedidoVentaItem In mcolPedidoVentaItems
        If Not objPedidoVentaItem.IsDeleted Then _
            Bruto = Bruto + objPedidoVentaItem.Bruto
    Next
    
End Property

Public Property Get Cantidad() As Double
    Dim objPedidoVentaItem As PedidoVentaItem
    
    Cantidad = 0
    For Each objPedidoVentaItem In mcolPedidoVentaItems
        If Not objPedidoVentaItem.IsDeleted Then _
            Cantidad = Cantidad + objPedidoVentaItem.Cantidad
    Next
    
End Property

Public Property Get Servido() As Double
    Dim objPedidoVentaItem As PedidoVentaItem
    
    Servido = 0
    For Each objPedidoVentaItem In mcolPedidoVentaItems
        If Not objPedidoVentaItem.IsDeleted Then _
            Servido = Servido + objPedidoVentaItem.Servido
    Next
    
End Property

Public Property Get OrdenCorte() As Boolean
    Dim objPedidoVentaItem As PedidoVentaItem
    
    OrdenCorte = 0
    For Each objPedidoVentaItem In mcolPedidoVentaItems
        If Not objPedidoVentaItem.IsDeleted Then _
            OrdenCorte = OrdenCorte Or objPedidoVentaItem.OrdenCorte
    Next
    
End Property

Friend Function GetState() As String
    Dim objBuffer As Buffer
    Dim objPedidoVentaItem As PedidoVentaItem
    
    Set objBuffer = New Buffer
    Set objPedidoVentaItem = New PedidoVentaItem
    objBuffer.Initialize Len(objPedidoVentaItem.GetState), 20
    
    For Each objPedidoVentaItem In mcolPedidoVentaItems
        objBuffer.Add objPedidoVentaItem.GetState
    Next
    
    GetState = objBuffer.GetState
    Set objBuffer = Nothing
    
End Function

Friend Sub SetState(Buffer As String)
    Dim lngIndex As Long
    Dim objPedidoVentaItem As PedidoVentaItem
    Dim objBuffer As Buffer
    
    Set mcolPedidoVentaItems = Nothing
    Set mcolPedidoVentaItems = New Collection
    
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    
    For lngIndex = 1 To objBuffer.Count
        Set objPedidoVentaItem = New PedidoVentaItem
        With objPedidoVentaItem
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me, mMoneda
        End With
        mcolPedidoVentaItems.Add Item:=objPedidoVentaItem
    Next
    
End Sub

Public Sub Actualizar()
    Dim objPedidoVentaItem As PedidoVentaItem
    
    For Each objPedidoVentaItem In mcolPedidoVentaItems
        objPedidoVentaItem.Actualizar
    Next
End Sub

Public Sub CambiarComision(Comision As Double)
    Dim objPedidoVentaItem As PedidoVentaItem
    
    For Each objPedidoVentaItem In mcolPedidoVentaItems
        objPedidoVentaItem.BeginEdit mMoneda
        objPedidoVentaItem.Comision = Comision
        objPedidoVentaItem.ApplyEdit
    Next
End Sub

' Asigna la comision a los items de pedido que son nuevos (que no lo tienen asignado).
Private Sub AsignarComision(Comision As Double)
    Dim objPedidoVentaItem As PedidoVentaItem
    
    For Each objPedidoVentaItem In mcolPedidoVentaItems
        If objPedidoVentaItem.IsNew Then
            objPedidoVentaItem.BeginEdit mMoneda
            objPedidoVentaItem.Comision = Comision
            objPedidoVentaItem.ApplyEdit
        End If
    Next
End Sub

Friend Sub ActualizarPreciosVenta()
    Dim objPedidoVentaItem As PedidoVentaItem
    Dim dblPrecioArticulo As Double
    Dim dblPrecioPedido As Double
    
    For Each objPedidoVentaItem In mcolPedidoVentaItems
        If objPedidoVentaItem.Pendiente > 0 Then
            dblPrecioPedido = objPedidoVentaItem.PrecioVenta
            dblPrecioArticulo = objPedidoVentaItem.objArticuloColor.ObjArticulo.PrecioVenta
            If dblPrecioArticulo <> dblPrecioPedido Then
               objPedidoVentaItem.BeginEdit mMoneda
               objPedidoVentaItem.PrecioVenta = dblPrecioArticulo
               objPedidoVentaItem.ApplyEdit
            End If
        End If
    Next

End Sub

Friend Property Get IsDirty() As Boolean
    Dim objItem As PedidoVentaItem
    
    IsDirty = False
    
    For Each objItem In mcolPedidoVentaItems
        If objItem.IsDirty Or objItem.IsNew Then
            IsDirty = True
            Exit For
        End If
    
    Next
    
End Property
