VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PedidoCompraItemMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements PedidoCompraItem

Event Valid(IsValid As Boolean)

' Aqui declaro el ComboBox de la lista de materiales
Private mobjMateriales As TextList

'Private mMoneda As String

Private mudtprops As PedidoCompraItemProps
Private mudtSave As PedidoCompraItemProps
Private mudtChild As PedidoCompraItemProps
Private mobjParent As PedidoCompraItems

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChild As Boolean
Private mflgChildEditing As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit() 'Moneda As String)

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
  
    'mMoneda = Moneda
    
    ' save object state
    LSet mudtSave = mudtprops
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    
    ' restore object state
    LSet mudtprops = mudtSave
  
End Sub

Public Sub ApplyEdit()
    Dim objPersist As PedidoCompraItemPersist
  
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445

    If Not mflgChild Then
        Set objPersist = CreateObject("GCServerMTS.PedidoCompraItemMaterialPersist", PERSIST_SERVER)
        If mflgDeleted And Not mflgNew Then
            ' code to delete the object's data goes here
            objPersist.DeleteObject mudtprops.PedidoCompraItemID
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
        Set objPersist = Nothing

    ElseIf mflgNew Then
        If Not mobjParent Is Nothing Then
            mobjParent.AddPedidoCompraItem Me
            Set mobjParent = Nothing
        End If
        
    End If
    
    mflgEditing = False

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
    mobjValid.RuleBroken "MaterialID", True
    mobjValid.RuleBroken "Cantidad", True
      
    With mudtprops
        .Situacion = "A"     ' Marco la situacion del pedido como "A"--> alta,
                        ' pendiente de codificar las posibles situaciones.
        .Cantidad = 0
        .Servido = 0
        .PrecioCostePTA = 0
        .PrecioCosteEUR = 0
        .Descuento = 0
        .BrutoPTA = 0
        .BrutoEUR = 0
        .Comision = 0
        .ActualizarAlta = False
        .DesactualizarAlta = False
        .ActualizarAlbaran = False
        .DesactualizarAlbaran = False
    End With
  
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

Public Sub Load(PedidoCompraItemMaterialID As Long) ', Moneda As String)
    Dim objPersist As PedidoCompraItemPersist
  
    If mflgChild Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    'mMoneda = Moneda
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.PedidoCompraItemMaterialPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(PedidoCompraItemMaterialID)
    Set objPersist = Nothing
    mobjValid.RuleBroken "MaterialID", False
    mobjValid.RuleBroken "Cantidad", False
  
    mflgNew = False

End Sub

Public Sub Delete()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
  
    If Servido Then _
        Err.Raise vbObjectError + 1001, "No se puede borrar, ya se han recibido entregas"
    
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

Friend Sub SetAsChild(objPedidoCompraItems As PedidoCompraItems) ', Moneda As String)

    'mMoneda = Moneda
    
    If mflgNew Then Set mobjParent = objPedidoCompraItems
    mflgChild = True

End Sub

Friend Sub ChildBeginEdit() 'Moneda As String)

    If mflgChildEditing Then Err.Raise 445
  
'    mMoneda = Moneda
  
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

Public Property Get PedidoCompraItemID() As Long

    PedidoCompraItemID = mudtprops.PedidoCompraItemID

End Property

Public Property Get Situacion() As String
  
    Select Case mudtprops.Situacion
        Case "A"
            Situacion = "Pendiente"
        Case "C"
            Situacion = "Servido"
        Case Else
            Err.Raise vbObjectError + 1001, "Situacion de pedido erronea"
    End Select

End Property

Public Property Get MaterialID() As Long

    MaterialID = mudtprops.MaterialID
  
End Property

Public Property Let Material(Value As String)
    Dim objMaterial As Material

    If Not mflgEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "MaterialID", (mobjMateriales.Key(Value) = 0)
  
    If mobjMateriales.Key(Value) = 0 Then Exit Property
  
    Set objMaterial = New Material
  
    objMaterial.Load mobjMateriales.Key(Value), "EUR"
  
    PrecioCoste = objMaterial.PrecioCoste
  
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

    If Not mflgEditing Then Err.Raise 383
  
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    If Servido Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
  
    ' No se permiten Cantidades negativas
    ' y la cantidad debe ser mayor que cero
    mobjValid.RuleBroken "Cantidad", (Value < 0)
  
    mudtprops.Cantidad = Round(Value, 2)
  
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True

End Property

Public Property Get Cantidad() As Double

    Cantidad = mudtprops.Cantidad

End Property

Public Property Let Servido(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "La cantidad servida del pedido no puede ser negativa."
  
    ' No se permiten Cantidades negativas ni que sea mayor que lo pedido
    mobjValid.RuleBroken "Servido", (Value < 0) Or (Value > mudtprops.Cantidad)
  
    ' Las cantidades son enteras
    mudtprops.Servido = Round(Value, 0)
  
    mflgDirty = True

End Property

Public Property Get Servido() As Double

    Servido = mudtprops.Servido

End Property

Public Property Get Pendiente() As Double

    Pendiente = Cantidad - Servido
    
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
  
    mudtprops.PrecioCosteEUR = Round(Value, 3)
    mudtprops.PrecioCostePTA = EUR2PTA(Value)
  
    CalcularBruto
    
    mflgDirty = True

End Property

Public Property Get PrecioCoste() As Double

    PrecioCoste = mudtprops.PrecioCosteEUR

End Property

Public Property Let Bruto(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "Bruto", (Value < 0)
  
    mudtprops.BrutoEUR = Round(Value, 2)
    mudtprops.BrutoPTA = EUR2PTA(Value)
    
    mflgDirty = True

End Property

Public Property Get Bruto() As Double

    Bruto = mudtprops.BrutoEUR
  
End Property

Friend Function GetState() As String
    Dim udtData As PedidoCompraItemData
  
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
    Dim udtData As PedidoCompraItemData

    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    mobjValid.RuleBroken "MaterialID", False
    mobjValid.RuleBroken "Cantidad", False
    
End Sub

Private Sub CalcularBruto()
    
    Bruto = PrecioCoste * Cantidad
    
End Sub

' Indica si hay que desactualizar o no
' Hay que hacerlo en caso de que no sea nuevo, haya cambiado el articulo,
' o hayan cambiado las cantidades
Private Function HayQueDesactualizar() As Boolean

    If Not mflgDirty Then
        HayQueDesactualizar = False
        Exit Function
    End If
  
    If mflgNew Then
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

' IMPLEMENTAR EL INTERFAZ
'Public Sub PedidoCompraItem_Load(PedidoCompraItemID As Long, Moneda As String, Tipo As PedidoCompraItemTipos)
Public Sub PedidoCompraItem_Load(PedidoCompraItemID As Long, Tipo As PedidoCompraItemTipos)

    If Tipo <> PEDIDOCOMPRAITEM_MATERIAL Then _
        Err.Raise vbObjectError + 1001, "PedidoCompraItemMaterial Load", "No se corresponde el tipo de pedido de compra: " & Tipo & " con el valor que se esperaba : PEDIDOCOMPRAITEM_MATERIAL"
    
    Load PedidoCompraItemID

End Sub

Public Property Get PedidoCompraItem_Tipo() As PedidoCompraItemTipos

    PedidoCompraItem_Tipo = PEDIDOCOMPRAITEM_MATERIAL
    
End Property

Private Sub PedidoCompraItem_BeginEdit() 'Moneda As String)

  BeginEdit

End Sub

Private Sub PedidoCompraItem_ApplyEdit()

  ApplyEdit

End Sub

Private Sub PedidoCompraItem_CancelEdit()

  CancelEdit

End Sub

Private Sub PedidoCompraItem_Delete()

  Delete

End Sub


Private Property Get PedidoCompraItem_IsDeleted() As Boolean

  PedidoCompraItem_IsDeleted = IsDeleted

End Property

Private Property Get PedidoCompraItem_IsDirty() As Boolean

  PedidoCompraItem_IsDirty = IsDirty

End Property

Private Property Get PedidoCompraItem_IsNew() As Boolean

  PedidoCompraItem_IsNew = IsNew

End Property

Private Property Get PedidoCompraItem_IsValid() As Boolean

  PedidoCompraItem_IsValid = IsValid

End Property


Public Property Get PedidoCompraItem_Cantidad() As Double

    PedidoCompraItem_Cantidad = Cantidad

End Property

Public Sub PedidoCompraItem_Actualizar()
    
    Actualizar
    
End Sub

Public Property Get PedidoCompraItem_PedidoCompraItemID() As Long

    PedidoCompraItem_PedidoCompraItemID = PedidoCompraItemID

End Property

Public Property Get PedidoCompraItem_Situacion() As String
  
    PedidoCompraItem_Situacion = Situacion
    
End Property

Public Property Let PedidoCompraItem_Servido(Value As Double)

    Servido = Value

End Property

Public Property Get PedidoCompraItem_Servido() As Double

    PedidoCompraItem_Servido = Servido

End Property

Public Property Get PedidoCompraItem_Pendiente() As Double

    PedidoCompraItem_Pendiente = Pendiente
    
End Property


Public Property Let PedidoCompraItem_Descuento(Value As Double)

    Descuento = Value
    
End Property

Public Property Get PedidoCompraItem_Descuento() As Double

    PedidoCompraItem_Descuento = Descuento

End Property

Public Property Let PedidoCompraItem_Comision(Value As Double)

    Comision = Value
  
End Property

Public Property Get PedidoCompraItem_Comision() As Double

    PedidoCompraItem_Comision = Comision

End Property

Public Property Let PedidoCompraItem_PrecioCoste(Value As Double)
  
    PrecioCoste = Value
  
End Property

Public Property Get PedidoCompraItem_PrecioCoste() As Double

    PedidoCompraItem_PrecioCoste = PrecioCoste
    
End Property

Public Property Let PedidoCompraItem_Bruto(Value As Double)

    Bruto = Value
  
End Property

Public Property Get PedidoCompraItem_Bruto() As Double

    PedidoCompraItem_Bruto = Bruto

End Property

