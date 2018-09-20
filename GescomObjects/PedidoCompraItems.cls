VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PedidoCompraItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMoneda As String

Private mcolPedidoCompraItems As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

    Set mcolPedidoCompraItems = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolPedidoCompraItems = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As PedidoCompraItem
Attribute Item.VB_UserMemId = 0

    Set Item = mcolPedidoCompraItems(Index)

End Function

Public Function Count() As Long

    Count = mcolPedidoCompraItems.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolPedidoCompraItems.[_NewEnum]

End Function

Public Sub BeginEdit(Moneda As String)
    Dim objPedidoCompraItem As PedidoCompraItem
    Dim objPedidoCompraItemMaterial As PedidoCompraItemMaterial
    Dim objPedidoCompraItemArticulo As PedidoCompraItemArticulo
  
    mMoneda = Moneda
    
    If mflgEditing Then Err.Raise 445
  
    For Each objPedidoCompraItem In mcolPedidoCompraItems
        Select Case objPedidoCompraItem.Tipo
        Case PEDIDOCOMPRAITEM_MATERIAL
            Set objPedidoCompraItemMaterial = objPedidoCompraItem
            objPedidoCompraItemMaterial.ChildBeginEdit 'Moneda
        Case PEDIDOCOMPRAITEM_ARTICULO
            Set objPedidoCompraItemArticulo = objPedidoCompraItem
            objPedidoCompraItemArticulo.ChildBeginEdit 'Moneda
        Case Else
            Err.Raise vbObjectError + 1001, "PedidoCompraItems BeginEdit", "No existe el tipo de item de pedido de compra:" & objPedidoCompraItem.Tipo & ". Avisar al personal técnico."
        End Select
    Next

    mflgEditing = True

End Sub

Public Sub CancelEdit()
    Dim objPedidoCompraItem As PedidoCompraItem
    Dim objPedidoCompraItemMaterial As PedidoCompraItemMaterial
    Dim objPedidoCompraItemArticulo As PedidoCompraItemArticulo
    Dim lngIndex As Long
  
    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    For lngIndex = mcolPedidoCompraItems.Count To 1 Step -1
        Set objPedidoCompraItem = mcolPedidoCompraItems.Item(lngIndex)
        Select Case objPedidoCompraItem.Tipo
        
        Case PEDIDOCOMPRAITEM_MATERIAL
            Set objPedidoCompraItemMaterial = objPedidoCompraItem
            
            With objPedidoCompraItemMaterial
                If .IsNew Then _
                    mcolPedidoCompraItems.Remove lngIndex
                .ChildCancelEdit
            End With
        Case PEDIDOCOMPRAITEM_ARTICULO
            Set objPedidoCompraItemArticulo = objPedidoCompraItem
            
            With objPedidoCompraItemArticulo
                If .IsNew Then _
                    mcolPedidoCompraItems.Remove lngIndex
                .ChildCancelEdit
            End With
        Case Else
            Err.Raise vbObjectError + 1001, "PedidoCompraItems CancelEdit", "No existe el tipo de item de pedido de compra:" & objPedidoCompraItem.Tipo & ". Avisar al personal técnico."
        End Select

    Next lngIndex

End Sub

Public Sub ApplyEdit() 'PedidoCompraID As Long)
  
    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False

End Sub

Public Sub Delete()

    Dim objPedidoCompraItem As PedidoCompraItem
  
    If Not mflgEditing Then Err.Raise 445
    
    ' code to delete the object's data goes here
    For Each objPedidoCompraItem In mcolPedidoCompraItems
        With objPedidoCompraItem
            .BeginEdit 'mMoneda
            .Delete
            .ApplyEdit
        End With

    Next

End Sub

Public Function Add(Tipo As PedidoCompraItemTipos) As PedidoCompraItem

    Select Case Tipo
    Case PEDIDOCOMPRAITEM_MATERIAL
        Set Add = AddMaterial
    Case PEDIDOCOMPRAITEM_ARTICULO
        Set Add = AddArticulo
    Case Else
        Err.Raise vbObjectError + 1001, "PedidoCompraItems Add", "No existe el tipo de item de pedido de compra:" & Tipo & ". Avisar al personal técnico."
    End Select
    
End Function
    
Private Function AddMaterial() As PedidoCompraItem
    Dim objPedidoCompraItemMaterial As PedidoCompraItemMaterial

    If Not mflgEditing Then Err.Raise 445

    Set objPedidoCompraItemMaterial = New PedidoCompraItemMaterial
    objPedidoCompraItemMaterial.SetAsChild Me ', mMoneda
    objPedidoCompraItemMaterial.ChildBeginEdit 'mMoneda
    Set AddMaterial = objPedidoCompraItemMaterial
    Set objPedidoCompraItemMaterial = Nothing

End Function
    
Private Function AddArticulo() As PedidoCompraItem
    Dim objPedidoCompraItemArticulo As PedidoCompraItemArticulo

    If Not mflgEditing Then Err.Raise 445

    Set objPedidoCompraItemArticulo = New PedidoCompraItemArticulo
    objPedidoCompraItemArticulo.SetAsChild Me   ', mMoneda
    objPedidoCompraItemArticulo.ChildBeginEdit 'mMoneda
    Set AddArticulo = objPedidoCompraItemArticulo
    Set objPedidoCompraItemArticulo = Nothing

End Function

Friend Sub AddPedidoCompraItem(Child As PedidoCompraItem)

    mcolPedidoCompraItems.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445

    With mcolPedidoCompraItems(Index)
        .BeginEdit 'mMoneda
        .Delete
        .ApplyEdit
    End With

End Sub

Friend Sub Load(PedidoCompraID As Long, Moneda As String)
    Dim objPersist As PedidoCompraItemsPersist
  
    mMoneda = Moneda
  
    Set objPersist = CreateObject("GCServerMTS.PedidoCompraItemsPersist", PERSIST_SERVER)
    SetMateriales objPersist.Fetch(PedidoCompraID)
    SetArticulos objPersist.FetchArticulos(PedidoCompraID)
    Set objPersist = Nothing

End Sub

Public Property Get Bruto() As Double
    Dim objPedidoCompraItem As PedidoCompraItem

    Bruto = 0
    For Each objPedidoCompraItem In mcolPedidoCompraItems
        If Not objPedidoCompraItem.IsDeleted Then _
            Bruto = Bruto + objPedidoCompraItem.Bruto
    Next

End Property

Public Property Get Cantidad() As Double
    Dim objPedidoCompraItem As PedidoCompraItem

    Cantidad = 0
    For Each objPedidoCompraItem In mcolPedidoCompraItems
        If Not objPedidoCompraItem.IsDeleted Then _
            Cantidad = Cantidad + objPedidoCompraItem.Cantidad
    Next

End Property

Public Property Get Servido() As Double
    Dim objPedidoCompraItem As PedidoCompraItem

    Servido = 0
    For Each objPedidoCompraItem In mcolPedidoCompraItems
        If Not objPedidoCompraItem.IsDeleted Then _
            Servido = Servido + objPedidoCompraItem.Servido
    Next

End Property

Friend Function GetMateriales() As String
    Dim objBuffer As Buffer
    Dim objPedidoCompraItem As PedidoCompraItem
    Dim objPedidoCompraItemMaterial As PedidoCompraItemMaterial
  
    Set objBuffer = New Buffer
    Set objPedidoCompraItemMaterial = New PedidoCompraItemMaterial
    objBuffer.Initialize Len(objPedidoCompraItemMaterial.GetState), 20
    Set objPedidoCompraItemMaterial = Nothing

    For Each objPedidoCompraItem In mcolPedidoCompraItems
        If objPedidoCompraItem.Tipo = PEDIDOCOMPRAITEM_MATERIAL Then
            Set objPedidoCompraItemMaterial = objPedidoCompraItem
            objBuffer.Add objPedidoCompraItemMaterial.GetState
        End If
    Next

    GetMateriales = objBuffer.GetState
    Set objBuffer = Nothing

End Function

Friend Function GetArticulos() As String
    Dim objBuffer As Buffer
    Dim objPedidoCompraItem As PedidoCompraItem
    Dim objPedidoCompraArticulo As PedidoCompraItemArticulo
  
    Set objBuffer = New Buffer
    Set objPedidoCompraArticulo = New PedidoCompraItemArticulo
    objBuffer.Initialize Len(objPedidoCompraArticulo.GetState), 20
    Set objPedidoCompraArticulo = Nothing

    For Each objPedidoCompraItem In mcolPedidoCompraItems
        If objPedidoCompraItem.Tipo = PEDIDOCOMPRAITEM_ARTICULO Then
            Set objPedidoCompraArticulo = objPedidoCompraItem
            objBuffer.Add objPedidoCompraArticulo.GetState
        End If
    Next

    GetArticulos = objBuffer.GetState
    Set objBuffer = Nothing

End Function

Friend Sub SetMateriales(Buffer As String)
    Dim lngIndex As Long
    Dim objPedidoCompraItemMaterial As PedidoCompraItemMaterial
    Dim objBuffer As Buffer

    Set mcolPedidoCompraItems = Nothing
    Set mcolPedidoCompraItems = New Collection

    Set objBuffer = New Buffer
    objBuffer.SetState Buffer

    For lngIndex = 1 To objBuffer.Count
        Set objPedidoCompraItemMaterial = New PedidoCompraItemMaterial
        With objPedidoCompraItemMaterial
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me  ', mMoneda
        End With
        mcolPedidoCompraItems.Add Item:=objPedidoCompraItemMaterial
    Next

End Sub

Friend Sub SetArticulos(Buffer As String)
    Dim lngIndex As Long
    Dim objPedidoCompraItemArticulo As PedidoCompraItemArticulo
    Dim objBuffer As Buffer

    'Set mcolPedidoCompraItems = Nothing
    'Set mcolPedidoCompraItems = New Collection

    Set objBuffer = New Buffer
    objBuffer.SetState Buffer

    For lngIndex = 1 To objBuffer.Count
        Set objPedidoCompraItemArticulo = New PedidoCompraItemArticulo
        With objPedidoCompraItemArticulo
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me  ', mMoneda
        End With
        mcolPedidoCompraItems.Add Item:=objPedidoCompraItemArticulo
    Next

End Sub


Public Sub Actualizar()
    Dim objPedidoCompraItem As PedidoCompraItem

    For Each objPedidoCompraItem In mcolPedidoCompraItems
        objPedidoCompraItem.Actualizar
    Next
    
End Sub

Friend Property Get IsDirty() As Boolean
    Dim objItem As PedidoCompraItem
    
    IsDirty = False
    
    For Each objItem In mcolPedidoCompraItems
        If objItem.IsDirty Or objItem.IsNew Then
            IsDirty = True
            Exit For
        End If
    
    Next
    
End Property

