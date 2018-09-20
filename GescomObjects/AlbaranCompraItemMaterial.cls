VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlbaranCompraItemMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements AlbaranCompraItem

Event Valid(IsValid As Boolean)

' Aqui declaro el ComboBox de la lista de materiales.
Private mobjMateriales As TextList

Private mMoneda As String
Private mTipo As PedidoCompraItemTipos

Private mudtprops As AlbaranCompraItemProps
Private mudtSave As AlbaranCompraItemProps
Private mudtChild As AlbaranCompraItemProps
Private mobjParent As AlbaranCompraItems
Private mobjPedidoCompraItem As PedidoCompraItem
Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChild As Boolean
Private mflgChildEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit(Moneda As String)

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    
    mMoneda = Moneda
    ' save object state
    LSet mudtSave = mudtprops
    mflgEditing = True
    
    ' Si hay pedido lo cargamos para posteriores validaciones
    If HayPedido Then
        Set mobjPedidoCompraItem = New PedidoCompraItem
        mobjPedidoCompraItem.Load mudtprops.PedidoCompraItemID, mTipo
    End If

End Sub

Public Sub CancelEdit()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    
    ' restore object state
    LSet mudtprops = mudtSave
    
    ' Eliminamos el pedido de memoria que habiamos cargado para validaciones.
    If HayPedido Then Set mobjPedidoCompraItem = Nothing

End Sub

Public Sub ApplyEdit()
    Dim objPersist As AlbaranCompraItemPersist
    
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    Set objPersist = CreateObject("GCServerMTS.AlbaranCompraItemPersist", PERSIST_SERVER)
    
    If Not mflgChild Then
        If mflgDeleted And Not mflgNew Then
            ' code to delete the object's data goes here
            objPersist.DeleteObject mudtprops.AlbaranCompraItemID
            mflgNew = True
            mflgDeleted = False
    
        ElseIf mflgDirty Or mflgNew Then
            If Not IsValid Then Err.Raise 445
            ' save object to database if appropriate
            Actualizar
            SetState objPersist.Save(GetState)
            ' save object state
            LSet mudtSave = mudtprops
            mflgNew = False
    
        End If
      
        mflgDirty = False
    
    ElseIf mflgNew Then
        If Not mobjParent Is Nothing Then
            mobjParent.AddAlbaranCompraItem Me
            Set mobjParent = Nothing
        End If
    
    End If
    
    Set objPersist = Nothing
    mflgEditing = False
  
    ' Eliminamos el pedido de memoria que habiamos cargado para validaciones.
    If HayPedido Then Set mobjPedidoCompraItem = Nothing

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    Set mobjValid = New BrokenRules
  
    Set mobjMateriales = New TextList
    mobjMateriales.Load "vNombreMateriales"
    mudtprops.MaterialID = mobjMateriales.Key(mobjMateriales.IndexItem(1))
  
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    'mobjValid.RuleBroken "Descripcion", True
    mobjValid.RuleBroken "Cantidad", True
      
    With mudtprops
        .Situacion = "A"    ' Marco la situacion del Albaran como "A"--> alta,
                        ' pendiente de codificar las posibles situaciones.
        '.Descripcion = vbNullString
        .Cantidad = 0
        .PrecioCostePTA = 0
        .PrecioCosteEUR = 0
        .Descuento = 0
        .BrutoPTA = 0
        .BrutoEUR = 0
        .Comision = 0
        .ActualizarAlta = False
        .DesactualizarAlta = False
        .ActualizarFactura = False
        .DesactualizarFactura = False
    End With
  
    mTipo = PEDIDOCOMPRAITEM_MATERIAL

End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
  
    Set mobjMateriales = Nothing

End Sub

Public Property Get IsValid() As Boolean

    IsValid = (mobjValid.Count = 0)

End Property

Private Sub mobjValid_BrokenRule()

    RaiseEvent Valid(False)

End Sub

Private Sub mobjValid_NoBrokenRules()

    RaiseEvent Valid(True)

End Sub

Public Sub Load(AlbaranCompraItemID As Long, Moneda As String)
    Dim objPersist As AlbaranCompraItemPersist
  
    If mflgChild Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    mMoneda = Moneda
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.AlbaranCompraItemPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(AlbaranCompraItemID)
    Set objPersist = Nothing
  
    mobjValid.RuleBroken "MaterialID", False
    mobjValid.RuleBroken "Descripcion", False
    mobjValid.RuleBroken "Cantidad", False

    mflgNew = False

End Sub

Public Sub Delete()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
  
    If Facturado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han facturado"
  
    mflgDeleted = True
    mflgDirty = True
  
End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty Or mflgNew

End Property

Friend Sub SetAsChild(objAlbaranCompraItems As AlbaranCompraItems, Moneda As String)

    mMoneda = Moneda
    If mflgNew Then Set mobjParent = objAlbaranCompraItems
    mflgChild = True

End Sub

Friend Sub ChildBeginEdit(Moneda As String)

    If mflgChildEditing Then Err.Raise 445
  
    mMoneda = Moneda
    ' save object state
    LSet mudtChild = mudtprops
    mflgChildEditing = True

End Sub

Friend Sub ChildCancelEdit()

    If Not mflgChildEditing Then Err.Raise 445
  
    mflgChildEditing = False
    mflgDeleted = False
    ' restore object state
    LSet mudtprops = mudtChild
    
End Sub

Public Property Get AlbaranCompraItemID() As Long

    AlbaranCompraItemID = mudtprops.AlbaranCompraItemID

End Property

Public Property Let PedidoCompraItemID(Value As Long)

    mudtprops.PedidoCompraItemID = Value

End Property

Public Property Get PedidoCompraItemID() As Long

    PedidoCompraItemID = mudtprops.PedidoCompraItemID

End Property

Public Property Get Situacion() As String
    
    Select Case mudtprops.Situacion
        Case "A"
            Situacion = "Servido"
        Case "F"
            Situacion = "Facturado"
        Case Else
            Err.Raise vbObjectError + 1001, "Situacion de Albaran erronea"
    End Select
    
End Property

Public Property Get MaterialID() As Long

    MaterialID = mudtprops.MaterialID
  
End Property

Public Property Let MaterialID(Value As Long)

    mudtprops.MaterialID = Value

End Property

Public Property Let Material(Value As String)
    Dim objMaterial As Material

    If Not mflgEditing Then Err.Raise 383
      
    If HayPedido Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, está relacionado con un pedido"
  
    If Facturado Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han Facturado entregas"
        
    ' Hay que asignar una descripcion si no hemos seleccionado un material
    'mobjValid.RuleBroken "Descripcion", (Len(mudtProps.Descripcion) = 0 And mobjMateriales.Key(Value) = 0)
    
    If mobjMateriales.Key(Value) = 0 Then Exit Property
      
    Set objMaterial = New Material
      
    objMaterial.Load mobjMateriales.Key(Value), mMoneda
  
    PrecioCoste = objMaterial.PrecioCoste
    
    'Descripcion = objMaterial.Nombre
  
    mudtprops.MaterialID = objMaterial.MaterialID
  
    mflgDirty = True

End Property

Public Property Get Material() As String

    Material = mobjMateriales.Item(mudtprops.MaterialID)

End Property

Public Property Get Materiales() As TextList

    Set Materiales = mobjMateriales

End Property

Public Property Let Cantidad(Value As Double)

    If HayPedido Then
        ValidarCantidadTalla Value, mudtSave.Cantidad, _
            mobjPedidoCompraItem.Pendiente, _
            "Cantidad"
    Else
        ValidarCantidadTalla Value, 0, 0, "Cantidad"
    End If
  
    mudtprops.Cantidad = Round(Value, 2)
  
    ' Se valida la cantidad total de la linea de Albaran
    mobjValid.RuleBroken "Cantidad", (Cantidad = 0)
    CalcularBruto
    mflgDirty = True
      
End Property

Public Property Get Cantidad() As Double

    Cantidad = mudtprops.Cantidad

End Property

Public Property Get Facturado() As Boolean
    
    Facturado = (mudtprops.Situacion = "F")
  
End Property

Public Property Get HayPedido() As Boolean
  
    HayPedido = (mudtprops.PedidoCompraItemID <> 0)
  
End Property

Public Property Let Descuento(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "Descuento", (Value < 0)
  
    mudtprops.Descuento = Round(Value, 2)
  
    mflgDirty = True

End Property

Public Property Get Descuento() As Double

    Descuento = mudtprops.Descuento

End Property

Public Property Let Comision(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "Comision", (Value < 0)
  
    mudtprops.Comision = Round(Value, 2)
  
    mflgDirty = True

End Property

Public Property Get Comision() As Double

    Comision = mudtprops.Comision

End Property

Public Property Let PrecioCoste(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "PrecioCoste", (Value < 0)
  
    If EsEUR(mMoneda) Then
        mudtprops.PrecioCosteEUR = Round(Value, 3)
        mudtprops.PrecioCostePTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.PrecioCostePTA = Round(Value, 0)
        mudtprops.PrecioCosteEUR = PTA2EUR(Value)
    Else
        mudtprops.PrecioCostePTA = 0
        mudtprops.PrecioCosteEUR = 0
    End If
  
    CalcularBruto
    
    mflgDirty = True

End Property

Public Property Get PrecioCoste() As Double

    If EsEUR(mMoneda) Then
        PrecioCoste = mudtprops.PrecioCosteEUR
    ElseIf EsPTA(mMoneda) Then
        PrecioCoste = mudtprops.PrecioCostePTA
    Else
        PrecioCoste = 0
    End If
      
End Property

Public Property Let Bruto(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
        
'    If Value < 0 Then _
'        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
      
    'mobjValid.RuleBroken "Bruto", (Value < 0)
      
    If EsEUR(mMoneda) Then
        mudtprops.BrutoEUR = Round(Value, 2)
        mudtprops.BrutoPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.BrutoPTA = Round(Value, 0)
        mudtprops.BrutoEUR = PTA2EUR(Value)
    Else
        mudtprops.BrutoPTA = 0
        mudtprops.BrutoEUR = 0
    End If
    
    mflgDirty = True

End Property

Public Property Get Bruto() As Double

    If EsEUR(mMoneda) Then
        Bruto = mudtprops.BrutoEUR
    ElseIf EsPTA(mMoneda) Then
        Bruto = mudtprops.BrutoPTA
    Else
        Bruto = 0
    End If
  
End Property

Friend Function GetState() As String

    Dim udtData As AlbaranCompraItemData
  
    With mudtprops
        .IsNew = mflgNew
        .IsDeleted = mflgDeleted
        .IsDirty = mflgDirty
    End With
    
    LSet udtData = mudtprops
    GetState = udtData.Buffer
    'mflgNew = False
    
End Function

Friend Sub SetState(Buffer As String)

    Dim udtData As AlbaranCompraItemData
    
    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    mobjValid.RuleBroken "MaterialID", False
    mobjValid.RuleBroken "Descripcion", False
    mobjValid.RuleBroken "Cantidad", False
    
End Sub

Private Sub CalcularBruto()
    
    Bruto = PrecioCoste * Cantidad
    
End Sub

Private Sub ValidarCantidadTalla(CantidadTalla As Double, Inicial As Double, Pendiente As Double, _
    Propiedad As String)

    If Not mflgEditing Then Err.Raise 383
  
' Las cantidades pueden ser negativas, y se consideran de devolución
'    If CantidadTalla < 0 Then _
'        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    ' Si esta relacionado con un pedido se mira que no se pase sobre el pedido original.
    ' Al ser nuevo se compara con el valor que se pretende asignar.
    ' Si no es nuevo se compara ademas con el inicial.
    If HayPedido Then
        If mflgNew Then
            If CantidadTalla > Pendiente Then
                Err.Raise vbObjectError + 1001, "Cantidad no puede superar la del pedido original"
            End If
        Else
            If CantidadTalla - Inicial > Pendiente Then
                Err.Raise vbObjectError + 1001, "Cantidad no puede superar la del pedido original"
            End If
        End If
    End If
        
    ' No se permiten Cantidades nulas
    mobjValid.RuleBroken Propiedad, (CantidadTalla = 0)
    
End Sub
  
' Indica si hay que Desactualizar o no
' Hay que hacerlo en caso de que no sea nuevo, haya cambiado el articulo,
' o hayan cambiado las cantidades
Private Function HayQueDesactualizar() As Boolean

    If Not mflgDirty Then
        HayQueDesactualizar = False
        Exit Function
    End If
    
    If mflgNew Or (mudtprops.MaterialID = 0) Then
        HayQueDesactualizar = False
        Exit Function
    End If
        
    If mflgDeleted Then
        HayQueDesactualizar = True
        Exit Function
    End If
      
    If mudtSave.MaterialID = mudtprops.MaterialID And _
        mudtSave.Cantidad = mudtprops.Cantidad Then
        HayQueDesactualizar = False
        Exit Function
    End If
     
    HayQueDesactualizar = True
    
End Function

' Indica si hay que Actualizar o no
' Hay que hacerlo en caso de que sea nuevo, haya cambiado el articulo,
' o hayan cambiado las cantidades
Private Function HayQueActualizar() As Boolean

    If Not mflgDirty Then
        HayQueActualizar = False
        Exit Function
    End If
    
    If mflgDeleted Then
        HayQueActualizar = False
        Exit Function
    End If
      
    If mudtprops.MaterialID = 0 Then
        HayQueActualizar = False
        Exit Function
    End If
  
    If mflgNew Then
        HayQueActualizar = True
        Exit Function
    End If
      
    If mudtSave.MaterialID = mudtprops.MaterialID And _
        mudtSave.Cantidad = mudtprops.Cantidad Then
        HayQueActualizar = False
        Exit Function
    End If
     
    HayQueActualizar = True
    
End Function

Public Sub Actualizar()

    mudtprops.ActualizarAlta = HayQueActualizar
    mudtprops.DesactualizarAlta = HayQueDesactualizar
   
End Sub

' Este procedimiento crea los datos de un albaran a partir de los de un pedido.
Public Sub AlbaranDesdePedido(PedidoItemID As Long)
    Dim objPedidoCompraItemMaterial As PedidoCompraItemMaterial

    Set mobjPedidoCompraItem = New PedidoCompraItem
       
    mobjPedidoCompraItem.Load PedidoItemID, mTipo
    
    Select Case mTipo
    Case PEDIDOCOMPRAITEM_MATERIAL
        Set objPedidoCompraItemMaterial = mobjPedidoCompraItem
        Material = objPedidoCompraItemMaterial.Material
    Case PEDIDOCOMPRAITEM_ARTICULO
        ' OJO TRATAR EL CASO DE ARTÍCULOS
    Case Else
        Err.Raise vbObjectError + 1001, "AlbaranCompraItemMaterial AlbaranDesdePedido", "No existe el tipo de item de pedido de compra:" & mTipo & ". Avisar al personal técnico."
    End Select
    
    mudtprops.PedidoCompraItemID = mobjPedidoCompraItem.PedidoCompraItemID
    'mudtProps.ArticuloColorID = mobjpedidoCompraitem.ArticuloColorID
    ' aqui hay que cargar el nombre del articulo (o quizas nada).
    ' .Descripcion = vbNullString
    Cantidad = mobjPedidoCompraItem.Pendiente
      
    PrecioCoste = mobjPedidoCompraItem.PrecioCoste
    Descuento = mobjPedidoCompraItem.Descuento
    Bruto = mobjPedidoCompraItem.Bruto
    Comision = mobjPedidoCompraItem.Comision
       
    Set mobjPedidoCompraItem = Nothing
    
End Sub

'IMPLEMENTAR EL INTERFAZ
Public Sub AlbaranCompraItem_BeginEdit(Moneda As String)

    BeginEdit Moneda

End Sub

Public Sub AlbaranCompraItem_CancelEdit()

    CancelEdit

End Sub

Public Sub AlbaranCompraItem_ApplyEdit()

    ApplyEdit

End Sub

Public Property Get AlbaranCompraItem_IsValid() As Boolean

    AlbaranCompraItem_IsValid = IsValid

End Property

Public Sub AlbaranCompraItem_Delete()

    Delete

End Sub

Public Property Get AlbaranCompraItem_IsDeleted() As Boolean

    AlbaranCompraItem_IsDeleted = IsDeleted

End Property

Public Property Get AlbaranCompraItem_IsNew() As Boolean

    AlbaranCompraItem_IsNew = IsNew

End Property

Public Property Get AlbaranCompraItem_IsDirty() As Boolean

    AlbaranCompraItem_IsDirty = IsDirty

End Property

Public Sub AlbaranCompraItem_Load(AlbaranCompraItemID As Long, Moneda As String, Tipo As AlbaranCompraItemTipos)

    If Tipo <> ALBARANCOMPRAITEM_MATERIAL Then _
        Err.Raise vbObjectError + 1001, "AlbaranCompraItemMaterial Load", "No se corresponde el tipo de albaran de compra: " & Tipo & " con el valor que se esperaba : ALBARANCOMPRAITEM_MATERIAL"
    
    Load AlbaranCompraItemID, Moneda

End Sub

Public Property Get AlbaranCompraItem_Tipo() As AlbaranCompraItemTipos

    AlbaranCompraItem_Tipo = ALBARANCOMPRAITEM_MATERIAL

End Property

Public Property Get AlbaranCompraItem_AlbaranCompraItemID() As Long

    AlbaranCompraItem_AlbaranCompraItemID = AlbaranCompraItemID

End Property

Public Property Let AlbaranCompraItem_PedidoCompraItemID(Value As Long)

    PedidoCompraItemID = Value

End Property

Public Property Get AlbaranCompraItem_PedidoCompraItemID() As Long

    AlbaranCompraItem_PedidoCompraItemID = PedidoCompraItemID

End Property

Public Property Get AlbaranCompraItem_Situacion() As String

    AlbaranCompraItem_Situacion = Situacion
    
End Property

Public Property Let AlbaranCompraItem_Cantidad(Value As Double)

    Cantidad = Value

End Property

Public Property Get AlbaranCompraItem_Cantidad() As Double

    AlbaranCompraItem_Cantidad = Cantidad

End Property

Public Property Get AlbaranCompraItem_Facturado() As Boolean

    AlbaranCompraItem_Facturado = Facturado
  
End Property

Public Property Get AlbaranCompraItem_HayPedido() As Boolean

    AlbaranCompraItem_HayPedido = HayPedido
  
End Property

Public Property Let AlbaranCompraItem_Descuento(Value As Double)

    Descuento = Value

End Property

Public Property Get AlbaranCompraItem_Descuento() As Double

    AlbaranCompraItem_Descuento = Descuento

End Property

Public Property Let AlbaranCompraItem_Comision(Value As Double)

    Comision = Value

End Property

Public Property Get AlbaranCompraItem_Comision() As Double

    AlbaranCompraItem_Comision = Comision

End Property

Public Property Let AlbaranCompraItem_PrecioCoste(Value As Double)

    PrecioCoste = Value
  
End Property

Public Property Get AlbaranCompraItem_PrecioCoste() As Double

    AlbaranCompraItem_PrecioCoste = PrecioCoste

End Property

Public Property Let AlbaranCompraItem_Bruto(Value As Double)

    Bruto = Value

End Property

Public Property Get AlbaranCompraItem_Bruto() As Double

    AlbaranCompraItem_Bruto = Bruto

End Property

Public Sub AlbaranCompraItem_Actualizar()

    Actualizar

End Sub

Public Sub AlbaranCompraItem_AlbaranDesdePedido(PedidoItemID As Long)

    AlbaranDesdePedido (PedidoItemID)

End Sub

