VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlbaranVentaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private mobjArticuloColor As ArticuloColor

Private mstrCodigoArticuloColor As String

Private mudtprops As AlbaranVentaItemProps
Private mudtSave As AlbaranVentaItemProps
Private mudtChild As AlbaranVentaItemProps
Private mobjParent As AlbaranVentaItems
Private mobjPedidoVentaItem As PedidoVentaItem
Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChild As Boolean
Private mflgChildEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    
    ' save object state
    LSet mudtSave = mudtprops
    mflgEditing = True
    
    ' Si hay pedido lo cargamos para posteriores validaciones
    If HayPedido Then
        Set mobjPedidoVentaItem = New PedidoVentaItem
        mobjPedidoVentaItem.Load mudtprops.PedidoVentaItemID, "EUR"
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
    If HayPedido Then Set mobjPedidoVentaItem = Nothing

End Sub

Public Sub ApplyEdit()
    Dim objPersist As AlbaranVentaItemPersist
    
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    Set objPersist = CreateObject("GCServerMTS.AlbaranVentaItemPersist", PERSIST_SERVER)
    
    If Not mflgChild Then
        If mflgDeleted And Not mflgNew Then
            ' code to delete the object's data goes here
            objPersist.DeleteObject mudtprops.AlbaranVentaItemID
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
            mobjParent.AddAlbaranVentaItem Me
            Set mobjParent = Nothing
        End If
    
    End If
      
    Set objPersist = Nothing
    mflgEditing = False
    
    ' Eliminamos el pedido de memoria que habiamos cargado para validaciones.
    If HayPedido Then Set mobjPedidoVentaItem = Nothing
    
End Sub

Private Sub Class_Initialize()

    mflgNew = True
    Set mobjValid = New BrokenRules
    
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    'mobjValid.RuleBroken "ArticuloColorID", True
    mobjValid.RuleBroken "Descripcion", True
    mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "Cantidad", True
    
    With mudtprops
        .Situacion = "A"    ' Marco la situacion del Albaran como "A"--> alta,
                            ' pendiente de codificar las posibles situaciones.
        .ArticuloColorID = 0
        .Descripcion = vbNullString
        .CantidadT36 = 0
        .CantidadT38 = 0
        .CantidadT40 = 0
        .CantidadT42 = 0
        .CantidadT44 = 0
        .CantidadT46 = 0
        .CantidadT48 = 0
        .CantidadT50 = 0
        .CantidadT52 = 0
        .CantidadT54 = 0
        .CantidadT56 = 0
        .PrecioVentaPTA = 0
        .PrecioVentaEUR = 0
        .Descuento = 0
        .BrutoPTA = 0
        .BrutoEUR = 0
        .Comision = 0
        .TemporadaID = 0
        .AlmacenID = 0
        .ActualizarAlta = False
        .DesactualizarAlta = False
        .ActualizarFactura = False
        .DesactualizarFactura = False
    End With
  
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    
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

Public Sub Load(AlbaranVentaItemID As Long)
    Dim objPersist As AlbaranVentaItemPersist
  
    If mflgChild Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.AlbaranVentaItemPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(AlbaranVentaItemID)
    Set objPersist = Nothing
    
    mobjValid.RuleBroken "ArticuloColorID", False
    mobjValid.RuleBroken "Descripcion", False
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "Cantidad", False
    
    mflgNew = False
    
End Sub

Public Sub Delete()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    If Facturado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han Facturado entregas"
    
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

Friend Sub SetAsChild(objAlbaranVentaItems As AlbaranVentaItems)

    If mflgNew Then Set mobjParent = objAlbaranVentaItems
    mflgChild = True
    
End Sub

Friend Sub ChildBeginEdit()

    If mflgChildEditing Then Err.Raise 445
    
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

Public Property Get AlbaranVentaID() As Long

    AlbaranVentaID = mudtprops.AlbaranVentaID

End Property

Public Property Get AlbaranVentaItemID() As Long

    AlbaranVentaItemID = mudtprops.AlbaranVentaItemID

End Property

Public Property Let PedidoVentaItemID(Value As Long)

    mudtprops.PedidoVentaItemID = Value

End Property

Public Property Get PedidoVentaItemID() As Long

    PedidoVentaItemID = mudtprops.PedidoVentaItemID

End Property

Public Property Get Situacion() As String
  
    Select Case mudtprops.Situacion
    Case "A"
        Situacion = "Servido"
    Case "F"
        Situacion = "Facturado"
    Case Else
        Err.Raise vbObjectError + 1001, "Situacion de albaran erronea"
    End Select

End Property

Public Property Get ArticuloColorID() As Long

    ArticuloColorID = mudtprops.ArticuloColorID
  
End Property

Public Property Let ArticuloColorID(Value As Long)

    If Not mflgEditing Then Err.Raise 383
    
    If HayPedido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, está relacionado con un pedido"
    
    If Facturado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han Facturado entregas"
    
    Set mobjArticuloColor = New ArticuloColor
    
    mobjArticuloColor.Load Value, vbNullString, 0
    
    mobjValid.RuleBroken "ArticuloColorID", False
    
    PrecioVenta = mobjArticuloColor.objArticulo.PrecioVenta
    mudtprops.Descripcion = Trim(mstrCodigoArticuloColor) & "-" & _
                            mobjArticuloColor.objArticulo.NombreCompleto & _
                            " " & mobjArticuloColor.NombreColor
    
    
    mudtprops.ArticuloColorID = Value

    mflgDirty = True
    
End Property

Public Property Let CodigoArticuloColor(Value As String)

    If Not mflgEditing Then Err.Raise 383
    
    If HayPedido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, está relacionado con un pedido"
    
    If Facturado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han Facturado entregas"
    
    If Len(Trim(Value)) > 8 Then Exit Property
    
    mstrCodigoArticuloColor = UCase(Trim(Value))
     
    If Len(mstrCodigoArticuloColor) < 8 Then Exit Property
    
    Set mobjArticuloColor = New ArticuloColor
    
    mobjArticuloColor.LoadCodigo Trim(mstrCodigoArticuloColor), mudtprops.TemporadaID
    
    mobjValid.RuleBroken "ArticuloColorID", False
    
    PrecioVenta = mobjArticuloColor.objArticulo.PrecioVenta
    mudtprops.Descripcion = Trim(mstrCodigoArticuloColor) & "-" & _
                            mobjArticuloColor.objArticulo.NombreCompleto & _
                            " " & mobjArticuloColor.NombreColor
    
    mudtprops.ArticuloColorID = mobjArticuloColor.ArticuloColorID
    
    mflgDirty = True
    
End Property

Public Property Get CodigoArticuloColor() As String
    
    CodigoArticuloColor = mstrCodigoArticuloColor
    
End Property

Public Property Get objArticuloColor() As ArticuloColor
    
    If mudtprops.ArticuloColorID = 0 Then
        Set objArticuloColor = Nothing
        Exit Property
    End If
    
    If mobjArticuloColor Is Nothing Then _
        Set mobjArticuloColor = New ArticuloColor
        
    If mobjArticuloColor.ArticuloColorID = 0 Then
        mobjArticuloColor.Load mudtprops.ArticuloColorID, "EUR"
    End If
    Set objArticuloColor = mobjArticuloColor

End Property

Public Property Get ArticuloColor() As String

    ArticuloColor = Trim(mudtprops.Descripcion)

End Property

Public Property Let Descripcion(Value As String)

    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.Descripcion) Then _
      Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
      
    ' Hay que asignar una descripcion si no hemos seleccionado un articulo
    mobjValid.RuleBroken "Descripcion", (Len(Trim(Value)) = 0 And mudtprops.ArticuloColorID = 0)
    
    mudtprops.Descripcion = Trim(Value)
    mflgDirty = True
    
End Property

Public Property Get Descripcion() As String

    Descripcion = Trim(mudtprops.Descripcion)

End Property

Public Property Let CantidadT36(Value As Double)

    If HayPedido Then
       ValidarCantidadTalla Value, mudtSave.CantidadT36, _
                    mobjPedidoVentaItem.PendienteT36
                    ', _
                    '"CantidadT36"
    Else
       ValidarCantidadTalla Value, 0, 0 ', "CantidadT36"
    End If
    
    ' Las cantidades son enteras
    mudtprops.CantidadT36 = Round(Value, 0)
    
    ' Que el total de cantidades por talla no sea cero
    ValidarCantidad
    
End Property

Public Property Get CantidadT36() As Double

    CantidadT36 = mudtprops.CantidadT36

End Property

Public Property Let CantidadT38(Value As Double)

    If HayPedido Then
       ValidarCantidadTalla Value, mudtSave.CantidadT38, _
                    mobjPedidoVentaItem.PendienteT38 ', _
                    '"CantidadT38"
    Else
       ValidarCantidadTalla Value, 0, 0 ', "CantidadT38"
    End If
    
    ' Las cantidades son enteras
    mudtprops.CantidadT38 = Round(Value, 0)
    
    ' Que el total de cantidades por talla no sea cero
    ValidarCantidad
    
End Property

Public Property Get CantidadT38() As Double

    CantidadT38 = mudtprops.CantidadT38

End Property

Public Property Let CantidadT40(Value As Double)

    If HayPedido Then
       ValidarCantidadTalla Value, mudtSave.CantidadT40, _
                    mobjPedidoVentaItem.PendienteT40 ', _
                    '"CantidadT40"
    Else
       ValidarCantidadTalla Value, 0, 0 ', "CantidadT40"
    End If
    
    ' Las cantidades son enteras
    mudtprops.CantidadT40 = Round(Value, 0)
    
    ' Que el total de cantidades por talla no sea cero
    ValidarCantidad
    
End Property

Public Property Get CantidadT40() As Double

    CantidadT40 = mudtprops.CantidadT40

End Property

Public Property Let CantidadT42(Value As Double)

  If HayPedido Then
     ValidarCantidadTalla Value, mudtSave.CantidadT42, _
                  mobjPedidoVentaItem.PendienteT42 ', _
                  '"CantidadT42"
  Else
     ValidarCantidadTalla Value, 0, 0 ', "CantidadT42"
  End If
  
  ' Las cantidades son enteras
  mudtprops.CantidadT42 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT42() As Double

  CantidadT42 = mudtprops.CantidadT42

End Property
Public Property Let CantidadT44(Value As Double)

  If HayPedido Then
     ValidarCantidadTalla Value, mudtSave.CantidadT44, _
                  mobjPedidoVentaItem.PendienteT44 ', _
                  '"CantidadT44"
  Else
     ValidarCantidadTalla Value, 0, 0 ', "CantidadT44"
  End If
  
  ' Las cantidades son enteras
  mudtprops.CantidadT44 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT44() As Double

  CantidadT44 = mudtprops.CantidadT44

End Property
Public Property Let CantidadT46(Value As Double)

  If HayPedido Then
     ValidarCantidadTalla Value, mudtSave.CantidadT46, _
                  mobjPedidoVentaItem.PendienteT46 ', _
                  '"CantidadT46"
  Else
     ValidarCantidadTalla Value, 0, 0 ', "CantidadT46"
  End If
  
  ' Las cantidades son enteras
  mudtprops.CantidadT46 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT46() As Double

  CantidadT46 = mudtprops.CantidadT46

End Property
Public Property Let CantidadT48(Value As Double)

  If HayPedido Then
     ValidarCantidadTalla Value, mudtSave.CantidadT48, _
                  mobjPedidoVentaItem.PendienteT48 ', _
                  '"CantidadT48"
  Else
     ValidarCantidadTalla Value, 0, 0 ', "CantidadT48"
  End If
  
  ' Las cantidades son enteras
  mudtprops.CantidadT48 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT48() As Double

  CantidadT48 = mudtprops.CantidadT48

End Property
Public Property Let CantidadT50(Value As Double)

  If HayPedido Then
     ValidarCantidadTalla Value, mudtSave.CantidadT50, _
                  mobjPedidoVentaItem.PendienteT50 ', _
                  '"CantidadT50"
  Else
     ValidarCantidadTalla Value, 0, 0 ', "CantidadT50"
  End If
  
  ' Las cantidades son enteras
  mudtprops.CantidadT50 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT50() As Double

  CantidadT50 = mudtprops.CantidadT50

End Property
Public Property Let CantidadT52(Value As Double)

  If HayPedido Then
     ValidarCantidadTalla Value, mudtSave.CantidadT52, _
                  mobjPedidoVentaItem.PendienteT52 ', _
                  '"CantidadT52"
  Else
     ValidarCantidadTalla Value, 0, 0 ', "CantidadT52"
  End If
  
  ' Las cantidades son enteras
  mudtprops.CantidadT52 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT52() As Double

  CantidadT52 = mudtprops.CantidadT52

End Property
Public Property Let CantidadT54(Value As Double)

  If HayPedido Then
     ValidarCantidadTalla Value, mudtSave.CantidadT54, _
                  mobjPedidoVentaItem.PendienteT54 ', _
                  '"CantidadT54"
  Else
     ValidarCantidadTalla Value, 0, 0 ', "CantidadT54"
  End If
  
  ' Las cantidades son enteras
  mudtprops.CantidadT54 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT54() As Double

  CantidadT54 = mudtprops.CantidadT54

End Property
Public Property Let CantidadT56(Value As Double)

  If HayPedido Then
     ValidarCantidadTalla Value, mudtSave.CantidadT56, _
                  mobjPedidoVentaItem.PendienteT56 ', _
                  '"CantidadT56"
  Else
     ValidarCantidadTalla Value, 0, 0 ', "CantidadT56"
  End If
  
  ' Las cantidades son enteras
  mudtprops.CantidadT56 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT56() As Double

  CantidadT56 = mudtprops.CantidadT56

End Property

Public Property Get Cantidad() As Double
  With mudtprops
  Cantidad = .CantidadT36 + .CantidadT38 + .CantidadT40 + .CantidadT42 + _
             .CantidadT44 + .CantidadT46 + .CantidadT48 + .CantidadT50 + _
             .CantidadT52 + .CantidadT54 + .CantidadT56
  End With

End Property

Public Property Get Facturado() As Boolean
  
  Facturado = (mudtprops.Situacion = "F") Or mudtprops.FacturadoAB
  
End Property

Public Property Get HayPedido() As Boolean
  
  HayPedido = (mudtprops.PedidoVentaItemID <> 0)
  
End Property

Public Property Let Descuento(Value As Double)

  If Not mflgEditing Then Err.Raise 383
  
  mobjValid.RuleBroken "Descuento", (Value < 0 Or Value > 100)
  
  mudtprops.Descuento = Round(Value, 2)
  
  CalcularBruto
  
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

Public Property Let TemporadaID(Value As Long)
    
    If Not mflgEditing Then Err.Raise 383
    
    mobjValid.RuleBroken "TemporadaID", (Value = 0)
    
    mudtprops.TemporadaID = Value
    
    mflgDirty = True

End Property

Public Property Get TemporadaID() As Long
    
    TemporadaID = mudtprops.TemporadaID
    
End Property

Public Property Let PrecioVenta(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
    If Value < 0 Then _
       Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
    
    mobjValid.RuleBroken "PrecioVenta", (Value < 0)
    
    mudtprops.PrecioVentaEUR = Round(Value, 2)
    mudtprops.PrecioVentaPTA = EUR2PTA(Value)
    
    CalcularBruto
      
    mflgDirty = True
    
End Property

Public Property Get PrecioVenta() As Double

    PrecioVenta = mudtprops.PrecioVentaEUR
    
End Property

Public Property Let Bruto(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
    'If Value < 0 Then _
    '   Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
    
    'mobjValid.RuleBroken "Bruto", (Value < 0)
    
    mudtprops.BrutoEUR = Round(Value, 2)
    mudtprops.BrutoPTA = EUR2PTA(Value)
      
    mflgDirty = True
    
End Property

Public Property Get Bruto() As Double

    Bruto = mudtprops.BrutoEUR
  
End Property

Public Property Let AlmacenID(Value As Long)
  
    If Not mflgEditing Then Err.Raise 383
      
    mudtprops.AlmacenID = Value
      
    mflgDirty = True
    
End Property

' Hay que eliminar este procedimiento, se hace en la capa de datos.
Public Sub FacturarAB()

'    If mflgEditing Then Err.Raise 383
'
'    If Me.Facturado Then _
'       Err.Raise vbObjectError + 1001, "Albarán ya facturado"
'
'    mudtprops.FacturadoAB = True
'    mudtprops.Situacion = "F"
'    mudtprops.IsDirty = True
'    mflgDirty = True

End Sub

Friend Function GetState() As String
    Dim udtData As AlbaranVentaItemData
    
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
    Dim udtData As AlbaranVentaItemData
    
    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    
    UnbrokeRules
    
    mstrCodigoArticuloColor = Left(mudtprops.Descripcion, 8)

End Sub

Private Sub UnbrokeRules()
    
    mobjValid.RuleBroken "ArticuloColorID", False
    mobjValid.RuleBroken "Descripcion", False
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "Cantidad", False

End Sub

' Esta funcion "encapsula" en un único string el estado de un objeto (lo serializa)
Public Function GetSuperState() As Byte()
    Dim objPB As PropertyBag
    
    If Not IsValid Then Err.Raise 445
    
    Set objPB = New PropertyBag
    
    With objPB
        .WriteProperty "State", GetState                            'Almacenar el estado del padre
        .WriteProperty "New", mflgNew
        .WriteProperty "Deleted", mflgDeleted
        .WriteProperty "Dirty", mflgDirty
        .WriteProperty "Editing", mflgEditing
        .WriteProperty "Child", mflgChild
        .WriteProperty "ChildEditing", mflgChildEditing
        GetSuperState = .Contents
    End With
    
    Set objPB = Nothing
End Function

' Esta funcion "deserializa" un objeto, a partir de un churro carga el estado.
Public Function SetSuperState(Buffer() As Byte)
    Dim objPB As PropertyBag
    
    Set objPB = New PropertyBag
    With objPB
        ' restaurar los contenidos del propertybag
        .Contents = Buffer
        SetState .ReadProperty("State")
        mflgNew = .ReadProperty("New")
        mflgDeleted = .ReadProperty("Deleted")
        mflgDirty = .ReadProperty("Dirty")
        mflgEditing = .ReadProperty("Editing")
        mflgChild = .ReadProperty("Child")
        mflgChildEditing = .ReadProperty("ChildEditing")
    End With
    
    UnbrokeRules
   
    Set objPB = Nothing
End Function

Private Sub CalcularBruto()
    Dim curImporteSinDescuento As Currency
    
    curImporteSinDescuento = PrecioVenta * Cantidad
    
    Bruto = curImporteSinDescuento - (curImporteSinDescuento * Descuento / 100)

End Sub

' Esta funcion calcula el descuento equivalente al importe bruto asignado.
' Si se asigna un precio determinado a la linea del albarán se calcula el descuento que resulta de ese precio final.
Public Sub CalcularDescuentoEquivalente()
    Dim curImporteSinDescuento As Currency
    
    curImporteSinDescuento = PrecioVenta * Cantidad
    
    mudtprops.Descuento = Round(100 * (curImporteSinDescuento - Bruto) / curImporteSinDescuento, 0)

End Sub

Private Sub ValidarCantidadTalla(CantidadTalla As Double, Inicial As Double, Pendiente As Double)
                            ', _
                            'Propiedad As String)

    If Not mflgEditing Then Err.Raise 383
    
' Puede ser negativo --> Albaranes de devolucion.
'    If CantidadTalla < 0 Then _
'      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
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
          
    ' No se permiten Cantidades negativas
'    mobjValid.RuleBroken Propiedad, (CantidadTalla < 0)
End Sub
  
Private Sub ValidarCantidad()
    ' Se valida la cantidad total de la linea de Albaranes
'    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    mobjValid.RuleBroken "Cantidad", (Cantidad = 0)
    CalcularBruto
    mflgDirty = True
End Sub

' Indica si hay que Desactualizar o no
' Hay que hacerlo en caso de que no sea nuevo, haya cambiado el articulo,
' o hayan cambiado las cantidades
Private Function HayQueDesactualizar() As Boolean

    If Not mflgDirty Then
        HayQueDesactualizar = False
        Exit Function
    End If
    
    If mflgNew Or (mudtprops.ArticuloColorID = 0) Then
        HayQueDesactualizar = False
        Exit Function
    End If
       
    If mflgDeleted Then
        HayQueDesactualizar = True
        Exit Function
    End If
    
    If mudtSave.ArticuloColorID = mudtprops.ArticuloColorID And _
       mudtSave.CantidadT36 = mudtprops.CantidadT36 And _
       mudtSave.CantidadT38 = mudtprops.CantidadT38 And _
       mudtSave.CantidadT40 = mudtprops.CantidadT40 And _
       mudtSave.CantidadT42 = mudtprops.CantidadT42 And _
       mudtSave.CantidadT44 = mudtprops.CantidadT44 And _
       mudtSave.CantidadT46 = mudtprops.CantidadT46 And _
       mudtSave.CantidadT48 = mudtprops.CantidadT48 And _
       mudtSave.CantidadT50 = mudtprops.CantidadT50 And _
       mudtSave.CantidadT52 = mudtprops.CantidadT52 And _
       mudtSave.CantidadT54 = mudtprops.CantidadT54 And _
       mudtSave.CantidadT56 = mudtprops.CantidadT56 Then
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
    
    If mudtprops.ArticuloColorID = 0 Then
       HayQueActualizar = False
       Exit Function
    End If
    
    If mflgNew Then
       HayQueActualizar = True
       Exit Function
    End If
    
    If mudtSave.ArticuloColorID = mudtprops.ArticuloColorID And _
       mudtSave.CantidadT36 = mudtprops.CantidadT36 And _
       mudtSave.CantidadT38 = mudtprops.CantidadT38 And _
       mudtSave.CantidadT40 = mudtprops.CantidadT40 And _
       mudtSave.CantidadT42 = mudtprops.CantidadT42 And _
       mudtSave.CantidadT44 = mudtprops.CantidadT44 And _
       mudtSave.CantidadT46 = mudtprops.CantidadT46 And _
       mudtSave.CantidadT48 = mudtprops.CantidadT48 And _
       mudtSave.CantidadT50 = mudtprops.CantidadT50 And _
       mudtSave.CantidadT52 = mudtprops.CantidadT52 And _
       mudtSave.CantidadT54 = mudtprops.CantidadT54 And _
       mudtSave.CantidadT56 = mudtprops.CantidadT56 Then
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
    
    Set mobjPedidoVentaItem = New PedidoVentaItem
    
    mobjPedidoVentaItem.Load PedidoItemID, "EUR"
      
    'ArticuloColor = mobjPedidoVentaItem.ArticuloColor
    mudtprops.TemporadaID = mobjPedidoVentaItem.TemporadaID
    CodigoArticuloColor = mobjPedidoVentaItem.CodigoArticuloColor
    mudtprops.PedidoVentaItemID = mobjPedidoVentaItem.PedidoVentaItemID
    'mudtProps.ArticuloColorID = mobjpedidoventaitem.ArticuloColorID
    ' aqui hay que cargar el nombre del articulo (o quizas nada).
    ' .Descripcion = vbNullString
    Descripcion = mobjPedidoVentaItem.ArticuloColor
    PrecioVenta = mobjPedidoVentaItem.PrecioVenta
    Descuento = mobjPedidoVentaItem.Descuento
    Bruto = mobjPedidoVentaItem.Bruto
    Comision = mobjPedidoVentaItem.Comision
    
    CantidadT36 = mobjPedidoVentaItem.PendienteT36
    CantidadT38 = mobjPedidoVentaItem.PendienteT38
    CantidadT40 = mobjPedidoVentaItem.PendienteT40
    CantidadT42 = mobjPedidoVentaItem.PendienteT42
    CantidadT44 = mobjPedidoVentaItem.PendienteT44
    CantidadT46 = mobjPedidoVentaItem.PendienteT46
    CantidadT48 = mobjPedidoVentaItem.PendienteT48
    CantidadT50 = mobjPedidoVentaItem.PendienteT50
    CantidadT52 = mobjPedidoVentaItem.PendienteT52
    CantidadT54 = mobjPedidoVentaItem.PendienteT54
    CantidadT56 = mobjPedidoVentaItem.PendienteT56
    
    Set mobjPedidoVentaItem = Nothing
End Sub

