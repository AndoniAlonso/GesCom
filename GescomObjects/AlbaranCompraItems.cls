VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlbaranCompraItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMoneda As String

Private mcolAlbaranCompraItems As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

    Set mcolAlbaranCompraItems = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolAlbaranCompraItems = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As AlbaranCompraItem
Attribute Item.VB_UserMemId = 0

    Set Item = mcolAlbaranCompraItems(Index)

End Function

Public Function Count() As Long

    Count = mcolAlbaranCompraItems.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolAlbaranCompraItems.[_NewEnum]

End Function

Public Sub BeginEdit(Moneda As String)
    Dim objAlbaranCompraItem As AlbaranCompraItem
    Dim objAlbaranCompraItemMaterial As AlbaranCompraItemMaterial
    Dim objAlbaranCompraItemArticulo As AlbaranCompraItemArticulo
  
    mMoneda = Moneda
    If mflgEditing Then Err.Raise 445
  
    For Each objAlbaranCompraItem In mcolAlbaranCompraItems
        Select Case objAlbaranCompraItem.Tipo
        Case PEDIDOCOMPRAITEM_MATERIAL
            Set objAlbaranCompraItemMaterial = objAlbaranCompraItem
            objAlbaranCompraItemMaterial.ChildBeginEdit Moneda
        Case PEDIDOCOMPRAITEM_ARTICULO
            Set objAlbaranCompraItemArticulo = objAlbaranCompraItem
            objAlbaranCompraItemArticulo.ChildBeginEdit
        Case Else
            Err.Raise vbObjectError + 1001, "AlbaranCompraItems BeginEdit", "No existe el tipo de item de pedido de compra:" & objAlbaranCompraItem.Tipo & ". Avisar al personal técnico."
        End Select
    Next

    mflgEditing = True

End Sub

Public Sub CancelEdit()

    Dim objAlbaranCompraItem As AlbaranCompraItem
    Dim objAlbaranCompraItemMaterial As AlbaranCompraItemMaterial
    Dim objAlbaranCompraItemArticulo As AlbaranCompraItemArticulo
    Dim lngIndex As Long
  
    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    For lngIndex = mcolAlbaranCompraItems.Count To 1 Step -1
        Set objAlbaranCompraItem = mcolAlbaranCompraItems.Item(lngIndex)
        
        Select Case objAlbaranCompraItem.Tipo
        Case ALBARANCOMPRAITEM_MATERIAL
            Set objAlbaranCompraItemMaterial = objAlbaranCompraItem
        
            With objAlbaranCompraItemMaterial
                If .IsNew Then _
                    mcolAlbaranCompraItems.Remove lngIndex
                .ChildCancelEdit
            End With

        Case ALBARANCOMPRAITEM_ARTICULO
            Set objAlbaranCompraItemArticulo = objAlbaranCompraItem
        
            With objAlbaranCompraItemArticulo
                If .IsNew Then _
                    mcolAlbaranCompraItems.Remove lngIndex
                .ChildCancelEdit
            End With
        Case Else
            Err.Raise vbObjectError + 1001, "AlbaranCompraItems BeginEdit", "No existe el tipo de item de pedido de compra:" & objAlbaranCompraItem.Tipo & ". Avisar al personal técnico."
        End Select

    Next lngIndex

End Sub

Public Sub ApplyEdit() 'AlbaranCompraID As Long)

    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    
End Sub

Public Sub Delete()

    Dim objAlbaranCompraItem As AlbaranCompraItem
  
    If Not mflgEditing Then Err.Raise 445
    ' code to delete the object's data goes here
    
    For Each objAlbaranCompraItem In mcolAlbaranCompraItems
        With objAlbaranCompraItem
            .BeginEdit mMoneda
            .Delete
            .ApplyEdit
        End With

    Next

End Sub

Public Function Add(Tipo As AlbaranCompraItemTipos) As AlbaranCompraItem

    Select Case Tipo
    Case ALBARANCOMPRAITEM_MATERIAL
        Set Add = AddMaterial
    Case ALBARANCOMPRAITEM_ARTICULO
        Set Add = AddArticulo
    Case Else
        Err.Raise vbObjectError + 1001, "AlbaranCompraItems Add", "No existe el tipo de item de albarán de compra:" & Tipo & ". Avisar al personal técnico."
    End Select
    
End Function
    
Public Function AddMaterial() As AlbaranCompraItem
    Dim objAlbaranCompraItemMaterial As AlbaranCompraItemMaterial

    If Not mflgEditing Then Err.Raise 445

    Set objAlbaranCompraItemMaterial = New AlbaranCompraItemMaterial
    objAlbaranCompraItemMaterial.SetAsChild Me, mMoneda
    objAlbaranCompraItemMaterial.ChildBeginEdit mMoneda
    Set AddMaterial = objAlbaranCompraItemMaterial
    Set objAlbaranCompraItemMaterial = Nothing

End Function

Public Function AddArticulo() As AlbaranCompraItem
    Dim objAlbaranCompraItemArticulo As AlbaranCompraItemArticulo

    If Not mflgEditing Then Err.Raise 445

    Set objAlbaranCompraItemArticulo = New AlbaranCompraItemArticulo
    objAlbaranCompraItemArticulo.SetAsChild Me
    objAlbaranCompraItemArticulo.ChildBeginEdit
    Set AddArticulo = objAlbaranCompraItemArticulo
    Set objAlbaranCompraItemArticulo = Nothing

End Function

Public Sub AddAlbaranCompraItem(Child As AlbaranCompraItem)

    mcolAlbaranCompraItems.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)
    If Not mflgEditing Then Err.Raise 445

    With Item(Index)
        .BeginEdit mMoneda
        .Delete
        .ApplyEdit
    End With
    
End Sub

Friend Sub Load(AlbaranCompraID As Long, Moneda As String)
    Dim objPersist As AlbaranCompraItemsPersist
  
    mMoneda = Moneda
    Set objPersist = CreateObject("GCServerMTS.AlbaranCompraItemsPersist", PERSIST_SERVER)
    SetMateriales objPersist.Fetch(AlbaranCompraID)
    SetArticulos objPersist.FetchArticulos(AlbaranCompraID)
    Set objPersist = Nothing
    
End Sub

Public Property Get Bruto() As Double

    Dim objAlbaranCompraItem As AlbaranCompraItem

    Bruto = 0
    For Each objAlbaranCompraItem In mcolAlbaranCompraItems
        If Not objAlbaranCompraItem.IsDeleted Then _
            Bruto = Bruto + objAlbaranCompraItem.Bruto
    Next

End Property

Public Property Get Cantidad() As Double

    Dim objAlbaranCompraItem As AlbaranCompraItem

    Cantidad = 0
    For Each objAlbaranCompraItem In mcolAlbaranCompraItems
        If Not objAlbaranCompraItem.IsDeleted Then _
            Cantidad = Cantidad + objAlbaranCompraItem.Cantidad
    Next

End Property

' Esta facturado cuando lo esta alguna de los items
Public Property Get Facturado() As Boolean
    
    Dim objAlbaranCompraItem As AlbaranCompraItem

    Facturado = False
    For Each objAlbaranCompraItem In mcolAlbaranCompraItems
        Facturado = Facturado Or objAlbaranCompraItem.Facturado
    Next
    
End Property

' Existen pedidos facturados cuando hay algún albaran que tiene pedidos.
Public Property Get HayPedido() As Boolean

    Dim objAlbaranCompraItem As AlbaranCompraItem

    HayPedido = False
    For Each objAlbaranCompraItem In mcolAlbaranCompraItems
        HayPedido = HayPedido Or objAlbaranCompraItem.HayPedido
    Next

End Property

Friend Function GetMateriales() As String   'GetState
    Dim objBuffer As Buffer
    Dim objAlbaranCompraItem As AlbaranCompraItem
    Dim objAlbaranCompraItemMaterial As AlbaranCompraItemMaterial
      
    Set objBuffer = New Buffer
    Set objAlbaranCompraItemMaterial = New AlbaranCompraItemMaterial
    objBuffer.Initialize Len(objAlbaranCompraItemMaterial.GetState), 20
    Set objAlbaranCompraItemMaterial = Nothing

    For Each objAlbaranCompraItem In mcolAlbaranCompraItems
        If objAlbaranCompraItem.Tipo = ALBARANCOMPRAITEM_MATERIAL Then
            Set objAlbaranCompraItemMaterial = objAlbaranCompraItem
            objBuffer.Add objAlbaranCompraItemMaterial.GetState
        End If
    Next
    
    GetMateriales = objBuffer.GetState
    Set objBuffer = Nothing
    Set objAlbaranCompraItem = Nothing
    Set objAlbaranCompraItemMaterial = Nothing

End Function

Friend Function GetArticulos() As String   'GetState
    Dim objBuffer As Buffer
    Dim objAlbaranCompraItem As AlbaranCompraItem
    Dim objAlbaranCompraItemArticulo As AlbaranCompraItemArticulo
      
    Set objBuffer = New Buffer
    Set objAlbaranCompraItemArticulo = New AlbaranCompraItemArticulo
    objBuffer.Initialize Len(objAlbaranCompraItemArticulo.GetState), 20
    Set objAlbaranCompraItemArticulo = Nothing

    For Each objAlbaranCompraItem In mcolAlbaranCompraItems
        If objAlbaranCompraItem.Tipo = ALBARANCOMPRAITEM_ARTICULO Then
            Set objAlbaranCompraItemArticulo = objAlbaranCompraItem
            objBuffer.Add objAlbaranCompraItemArticulo.GetState
        End If
    Next
    
    GetArticulos = objBuffer.GetState
    Set objBuffer = Nothing
    Set objAlbaranCompraItem = Nothing
    Set objAlbaranCompraItemArticulo = Nothing

End Function

Friend Sub SetMateriales(Buffer As String) 'SetState
    Dim lngIndex As Long
    Dim objAlbaranCompraItemMaterial As AlbaranCompraItemMaterial
    Dim objBuffer As Buffer

    Set mcolAlbaranCompraItems = Nothing
    Set mcolAlbaranCompraItems = New Collection

    Set objBuffer = New Buffer
    objBuffer.SetState Buffer

    For lngIndex = 1 To objBuffer.Count
        Set objAlbaranCompraItemMaterial = New AlbaranCompraItemMaterial
        With objAlbaranCompraItemMaterial
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me, mMoneda
        End With
        mcolAlbaranCompraItems.Add Item:=objAlbaranCompraItemMaterial
    Next

End Sub

Friend Sub SetArticulos(Buffer As String) 'SetState
    Dim lngIndex As Long
    Dim objAlbaranCompraItemArticulo As AlbaranCompraItemArticulo
    Dim objBuffer As Buffer

    'Set mcolAlbaranCompraItems = Nothing
    'Set mcolAlbaranCompraItems = New Collection

    Set objBuffer = New Buffer
    objBuffer.SetState Buffer

    For lngIndex = 1 To objBuffer.Count
        Set objAlbaranCompraItemArticulo = New AlbaranCompraItemArticulo
        With objAlbaranCompraItemArticulo
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me
        End With
        mcolAlbaranCompraItems.Add Item:=objAlbaranCompraItemArticulo
    Next

End Sub
Friend Sub Actualizar()
    Dim objAlbaranCompraItem As AlbaranCompraItem

    For Each objAlbaranCompraItem In mcolAlbaranCompraItems
        objAlbaranCompraItem.Actualizar
    Next
    
End Sub

Friend Property Get IsDirty() As Boolean
    Dim objItem As AlbaranCompraItem
    
    IsDirty = False
    
    For Each objItem In mcolAlbaranCompraItems
        If objItem.IsDirty Or objItem.IsNew Then
            IsDirty = True
            Exit For
        End If
    Next
    
End Property

