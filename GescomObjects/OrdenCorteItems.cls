VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrdenCorteItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMoneda As String

Private mcolOrdenCorteItems As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

    Set mcolOrdenCorteItems = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolOrdenCorteItems = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As OrdenCorteItem
Attribute Item.VB_UserMemId = 0

    Set Item = mcolOrdenCorteItems(Index)

End Function

Public Function Count() As Long

    Count = mcolOrdenCorteItems.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolOrdenCorteItems.[_NewEnum]

End Function

Public Sub BeginEdit(Moneda As String)
    Dim objOrdenCorteItem As OrdenCorteItem
    
    mMoneda = Moneda
    If mflgEditing Then Err.Raise 445
    
    For Each objOrdenCorteItem In mcolOrdenCorteItems
      objOrdenCorteItem.ChildBeginEdit Moneda
    Next
    
    mflgEditing = True
    
End Sub

Public Sub CancelEdit()
    Dim objOrdenCorteItem As OrdenCorteItem
    Dim lngIndex As Long
    
    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    For lngIndex = mcolOrdenCorteItems.Count To 1 Step -1
      Set objOrdenCorteItem = mcolOrdenCorteItems.Item(lngIndex)
      With objOrdenCorteItem
        If .IsNew Then _
           mcolOrdenCorteItems.Remove lngIndex
        .ChildCancelEdit
      End With
    
    Next lngIndex
    
End Sub

Public Sub ApplyEdit() 'OrdenCorteID As Long)

    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    
End Sub

Public Sub Delete()

    Dim objOrdenCorteItem As OrdenCorteItem
    
    If Not mflgEditing Then Err.Raise 445
    ' code to delete the object's data goes here
    
    For Each objOrdenCorteItem In mcolOrdenCorteItems
        With objOrdenCorteItem
            .BeginEdit mMoneda
            .Delete
            .ApplyEdit
        End With
    Next
    
End Sub

Public Function Add() As OrdenCorteItem

    Dim objOrdenCorteItem As OrdenCorteItem
    
    If Not mflgEditing Then Err.Raise 445
    
    Set objOrdenCorteItem = New OrdenCorteItem
    objOrdenCorteItem.SetAsChild Me, mMoneda
    objOrdenCorteItem.ChildBeginEdit mMoneda
    Set Add = objOrdenCorteItem
    Set objOrdenCorteItem = Nothing
    
End Function

Friend Sub AddOrdenCorteItem(Child As OrdenCorteItem)

    mcolOrdenCorteItems.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445
    
    With mcolOrdenCorteItems(Index)
      .BeginEdit mMoneda
      .Delete
      .ApplyEdit
    End With

End Sub

Friend Sub Load(OrdenCorteID As Long, Moneda As String)
    Dim objPersist As OrdenCorteItemsPersist
    
    mMoneda = Moneda
    Set objPersist = CreateObject("GCServerMTS.OrdenCorteItemsPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(OrdenCorteID)
    Set objPersist = Nothing
    
End Sub

Public Property Get Cantidad() As Double
    Dim objOrdenCorteItem As OrdenCorteItem
    
    Cantidad = 0
    For Each objOrdenCorteItem In mcolOrdenCorteItems
        Cantidad = Cantidad + objOrdenCorteItem.Cantidad
    Next
    
End Property

' Esta facturado cuando lo esta alguna de los items
Public Property Get OrdenCortada() As Boolean
    Dim objOrdenCorteItem As OrdenCorteItem
    
    OrdenCortada = False
    For Each objOrdenCorteItem In mcolOrdenCorteItems
        OrdenCortada = OrdenCortada Or objOrdenCorteItem.OrdenCortada
    Next
    
End Property

Friend Function GetState() As String
    Dim objBuffer As Buffer
    Dim objOrdenCorteItem As OrdenCorteItem
    
    Set objBuffer = New Buffer
    Set objOrdenCorteItem = New OrdenCorteItem
    objBuffer.Initialize Len(objOrdenCorteItem.GetState), 20
    
    For Each objOrdenCorteItem In mcolOrdenCorteItems
        objBuffer.Add objOrdenCorteItem.GetState
    Next
    
    GetState = objBuffer.GetState
    Set objBuffer = Nothing
    
End Function

Friend Sub SetState(Buffer As String)
    Dim lngIndex As Long
    Dim objOrdenCorteItem As OrdenCorteItem
    Dim objBuffer As Buffer
    
    Set mcolOrdenCorteItems = Nothing
    Set mcolOrdenCorteItems = New Collection
    
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    
    For lngIndex = 1 To objBuffer.Count
        Set objOrdenCorteItem = New OrdenCorteItem
        With objOrdenCorteItem
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me, mMoneda
        End With
        mcolOrdenCorteItems.Add Item:=objOrdenCorteItem
    Next
    
End Sub

Public Sub Actualizar()
    Dim objOrdenCorteItem As OrdenCorteItem
    
    For Each objOrdenCorteItem In mcolOrdenCorteItems
        objOrdenCorteItem.Actualizar
    Next
    
End Sub

Public Sub Desactualizar()
    Dim objOrdenCorteItem As OrdenCorteItem
    
    For Each objOrdenCorteItem In mcolOrdenCorteItems
        objOrdenCorteItem.Desactualizar
    Next
    
End Sub

Friend Property Get IsDirty() As Boolean
    Dim objItem As OrdenCorteItem
    
    IsDirty = False
    
    For Each objItem In mcolOrdenCorteItems
        If objItem.IsDirty Or objItem.IsNew Then
            IsDirty = True
            Exit For
        End If
    Next
    
End Property

' Devuelve el articulo del que están compuestos los items de ordenes de corte.
' Devuelve 0 si no hay items.
' Devuelve error si los items tienen distintos codigos de articulo.
Friend Property Get ArticuloID() As Long
    Dim objItem As OrdenCorteItem
    Dim mlngArtic As Long
    
    mlngArtic = 0
    For Each objItem In mcolOrdenCorteItems
        If Not objItem.IsDeleted Then
            If objItem.ArticuloID <> mlngArtic And mlngArtic <> 0 Then
                Err.Raise vbObjectError + 1001, _
                    "Una órden de corte debe componerse de pedidos del mismo artículo."
                Exit For
            End If
            mlngArtic = objItem.ArticuloID
        End If
    Next
    
    ArticuloID = mlngArtic
    
End Property
