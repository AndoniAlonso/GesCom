VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturaCompraItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMoneda As String

Private mcolFacturaCompraItems As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

    Set mcolFacturaCompraItems = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolFacturaCompraItems = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As FacturaCompraItem
Attribute Item.VB_UserMemId = 0

    Set Item = mcolFacturaCompraItems(Index)

End Function

Public Function Count() As Long

    Count = mcolFacturaCompraItems.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolFacturaCompraItems.[_NewEnum]

End Function

Public Sub BeginEdit(Moneda As String)
    Dim objFacturaCompraItem As FacturaCompraItem
  
    mMoneda = Moneda
    If mflgEditing Then Err.Raise 445
  
    For Each objFacturaCompraItem In mcolFacturaCompraItems
        objFacturaCompraItem.ChildBeginEdit Moneda
    Next

    mflgEditing = True

End Sub

Public Sub CancelEdit()
    Dim objFacturaCompraItem As FacturaCompraItem
    Dim lngIndex As Long
  
    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    For lngIndex = mcolFacturaCompraItems.Count To 1 Step -1
        Set objFacturaCompraItem = mcolFacturaCompraItems.Item(lngIndex)
        With objFacturaCompraItem
            If .IsNew Then _
                mcolFacturaCompraItems.Remove lngIndex
            .ChildCancelEdit
        End With

    Next lngIndex

End Sub

Public Sub ApplyEdit() 'FacturaCompraID As Long)

    If Not mflgEditing Then Err.Raise 445

    mflgEditing = False

End Sub

Public Sub Delete()
    Dim objFacturaCompraItem As FacturaCompraItem
  
    If Not mflgEditing Then Err.Raise 445
    ' code to delete the object's data goes here

    For Each objFacturaCompraItem In mcolFacturaCompraItems
        With objFacturaCompraItem
            .BeginEdit mMoneda
            .Delete
            .ApplyEdit
        End With

    Next

End Sub

Public Function Add() As FacturaCompraItem
    Dim objFacturaCompraItem As FacturaCompraItem
  
    If Not mflgEditing Then Err.Raise 445

    Set objFacturaCompraItem = New FacturaCompraItem
    objFacturaCompraItem.SetAsChild Me, mMoneda
    objFacturaCompraItem.ChildBeginEdit mMoneda
    Set Add = objFacturaCompraItem
    Set objFacturaCompraItem = Nothing

End Function

Public Sub AddFacturaCompraItem(Child As FacturaCompraItem)

    mcolFacturaCompraItems.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445

    With mcolFacturaCompraItems(Index)
        .BeginEdit mMoneda
        .Delete
        .ApplyEdit
    End With

End Sub

Friend Sub Load(FacturaCompraID As Long, Moneda As String)
    Dim objPersist As FacturaCompraItemsPersist
  
    mMoneda = Moneda
    Set objPersist = CreateObject("GCServerMTS.FacturaCompraItemsPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(FacturaCompraID)
    Set objPersist = Nothing

End Sub

Public Property Get Bruto() As Double
    Dim objFacturaCompraItem As FacturaCompraItem

    Bruto = 0
    For Each objFacturaCompraItem In mcolFacturaCompraItems
        If Not objFacturaCompraItem.IsDeleted Then
            Bruto = Bruto + objFacturaCompraItem.Bruto
        End If
    Next

End Property

Public Property Get Cantidad() As Double
    Dim objFacturaCompraItem As FacturaCompraItem

    Cantidad = 0
    For Each objFacturaCompraItem In mcolFacturaCompraItems
        Cantidad = Cantidad + objFacturaCompraItem.Cantidad
    Next

End Property

' Existen Albaranes Contabilizados cuando hay algún Factura que tiene Albaranes.
Public Property Get HayAlbaran() As Boolean
    Dim objFacturaCompraItem As FacturaCompraItem

    HayAlbaran = False
    For Each objFacturaCompraItem In mcolFacturaCompraItems
        HayAlbaran = HayAlbaran Or objFacturaCompraItem.HayAlbaran
    Next

End Property

Public Function GetState() As String
    Dim objBuffer As Buffer
    Dim objFacturaCompraItem As FacturaCompraItem
  
    Set objBuffer = New Buffer
    Set objFacturaCompraItem = New FacturaCompraItem
    objBuffer.Initialize Len(objFacturaCompraItem.GetState), 20
    
    For Each objFacturaCompraItem In mcolFacturaCompraItems
        objBuffer.Add objFacturaCompraItem.GetState
    Next

    GetState = objBuffer.GetState
    Set objBuffer = Nothing

End Function

Public Sub SetState(Buffer As String)
    Dim lngIndex As Long
    Dim objFacturaCompraItem As FacturaCompraItem
    Dim objBuffer As Buffer

    Set mcolFacturaCompraItems = Nothing
    Set mcolFacturaCompraItems = New Collection
    
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer

    For lngIndex = 1 To objBuffer.Count
        Set objFacturaCompraItem = New FacturaCompraItem
        With objFacturaCompraItem
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me, mMoneda
        End With
        mcolFacturaCompraItems.Add Item:=objFacturaCompraItem
    Next

End Sub

Public Sub Actualizar()
    Dim objFacturaCompraItem As FacturaCompraItem

    For Each objFacturaCompraItem In mcolFacturaCompraItems
        objFacturaCompraItem.Actualizar
    Next

End Sub

Friend Property Get IsDirty() As Boolean
    Dim objItem As FacturaCompraItem
    
    IsDirty = False
    
    For Each objItem In mcolFacturaCompraItems
        If objItem.IsDirty Or objItem.IsNew Then
            IsDirty = True
            Exit For
        End If
    Next
    
End Property

