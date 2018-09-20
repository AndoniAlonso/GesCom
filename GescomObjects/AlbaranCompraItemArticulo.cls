VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlbaranCompraItemArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements AlbaranCompraItem

Event Valid(IsValid As Boolean)

Private mobjArticuloColor As ArticuloColor

Private mstrCodigoArticuloColor As String

Private mudtprops As AlbaranCompraArticuloProps
Private mudtSave As AlbaranCompraArticuloProps
Private mudtChild As AlbaranCompraArticuloProps
Private mobjParent As AlbaranCompraItems
Private mobjPedidoCompraItemArticulo As PedidoCompraItemArticulo
Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChild As Boolean
Private mflgChildEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit()
    Dim objParametroAplicacion As ParametroAplicacion

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    
    ' save object state
    LSet mudtSave = mudtprops
    mflgEditing = True
    
    ' Si es nuevo asignamos el valor predeterminado del almacen
    If Me.IsNew Then
        Set objParametroAplicacion = New ParametroAplicacion
        mudtprops.AlmacenID = objParametroAplicacion.AlmacenPredeterminado
        Set objParametroAplicacion = Nothing
    End If
    
    ' Si hay pedido lo cargamos para posteriores validaciones
    If HayPedido Then
        Set mobjPedidoCompraItemArticulo = New PedidoCompraItemArticulo
        mobjPedidoCompraItemArticulo.Load mudtprops.PedidoCompraArticuloID ', "EUR"
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
    If HayPedido Then Set mobjPedidoCompraItemArticulo = Nothing

End Sub

Public Sub ApplyEdit()
    Dim objPersist As AlbaranCompraArtPersist
    
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    Set objPersist = CreateObject("GCServerMTS.AlbaranCompraArtPersist", PERSIST_SERVER)
    
    If Not mflgChild Then
        If mflgDeleted And Not mflgNew Then
            ' code to delete the object's data goes here
            objPersist.DeleteObject mudtprops.AlbaranCompraArticuloID
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
    If HayPedido Then Set mobjPedidoCompraItemArticulo = Nothing
    
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
        .PrecioCompraEUR = 0
        .Descuento = 0
        .BrutoEUR = 0
        .Comision = 0
        .TemporadaID = 0
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

Public Sub Load(AlbaranCompraArticuloID As Long)
    Dim objPersist As AlbaranCompraArtPersist
  
    If mflgChild Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.AlbaranCompraArtPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(AlbaranCompraArticuloID)
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

Friend Sub SetAsChild(objAlbaranCompraItems As AlbaranCompraItems)

    If mflgNew Then Set mobjParent = objAlbaranCompraItems
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

Public Property Get AlbaranCompraID() As Long

    AlbaranCompraID = mudtprops.AlbaranCompraID

End Property

Public Property Get AlbaranCompraArticuloID() As Long

    AlbaranCompraArticuloID = mudtprops.AlbaranCompraArticuloID

End Property

Public Property Let PedidoCompraArticuloID(Value As Long)

    mudtprops.PedidoCompraArticuloID = Value

End Property

Public Property Get PedidoCompraArticuloID() As Long

    PedidoCompraArticuloID = mudtprops.PedidoCompraArticuloID

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

    mudtprops.ArticuloColorID = Value

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
    
    PrecioCompra = mobjArticuloColor.objArticulo.PrecioCompra
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
        
    If mudtprops.ArticuloColorID <> mobjArticuloColor.ArticuloColorID Then
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
                    mobjPedidoCompraItemArticulo.PendienteT36 ', _
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
                    mobjPedidoCompraItemArticulo.PendienteT38 ', _
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
                    mobjPedidoCompraItemArticulo.PendienteT40 ', _
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
                  mobjPedidoCompraItemArticulo.PendienteT42 ', _
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
                  mobjPedidoCompraItemArticulo.PendienteT44 ', _
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
                  mobjPedidoCompraItemArticulo.PendienteT46 ', _
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
                  mobjPedidoCompraItemArticulo.PendienteT48 ', _
                  ' "CantidadT48"
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
                  mobjPedidoCompraItemArticulo.PendienteT50 ', _
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
                  mobjPedidoCompraItemArticulo.PendienteT52 ', _
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
                  mobjPedidoCompraItemArticulo.PendienteT54 ', _
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
                  mobjPedidoCompraItemArticulo.PendienteT56 ', _
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
  
  HayPedido = (mudtprops.PedidoCompraArticuloID <> 0)
  
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

Public Property Let PrecioCompra(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
    If Value < 0 Then _
       Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
    
    mobjValid.RuleBroken "PrecioCompra", (Value < 0)
    
    mudtprops.PrecioCompraEUR = Round(Value, 2)
    
    CalcularBruto
      
    mflgDirty = True
    
End Property

Public Property Get PrecioCompra() As Double

    PrecioCompra = mudtprops.PrecioCompraEUR
    
End Property

Public Property Let Bruto(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
    'If Value < 0 Then _
    '   Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
    
    'mobjValid.RuleBroken "Bruto", (Value < 0)
    
    mudtprops.BrutoEUR = Round(Value, 2)
      
    mflgDirty = True
    
End Property

Public Property Get Bruto() As Double

    Bruto = mudtprops.BrutoEUR
  
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
    Dim udtData As AlbaranCompraArticuloData
    
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
    Dim udtData As AlbaranCompraArticuloData
    
    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    mobjValid.RuleBroken "ArticuloColorID", False
    mobjValid.RuleBroken "Descripcion", False
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "Cantidad", False
    mstrCodigoArticuloColor = Left(mudtprops.Descripcion, 8)

End Sub

Private Sub CalcularBruto()
    Dim curImporteSinDescuento As Currency
    
    curImporteSinDescuento = PrecioCompra * Cantidad
    
    Bruto = curImporteSinDescuento - (curImporteSinDescuento * Descuento / 100)

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
    
    Set mobjPedidoCompraItemArticulo = New PedidoCompraItemArticulo
    
    mobjPedidoCompraItemArticulo.Load PedidoItemID ', "EUR"
      
    'ArticuloColor = mobjPedidoCompraItemArticulo.ArticuloColor
    mudtprops.TemporadaID = mobjPedidoCompraItemArticulo.TemporadaID
    CodigoArticuloColor = mobjPedidoCompraItemArticulo.CodigoArticuloColor
    mudtprops.PedidoCompraArticuloID = mobjPedidoCompraItemArticulo.PedidoCompraArticuloID
    'mudtProps.ArticuloColorID = mobjPedidoCompraItemArticulo.ArticuloColorID
    ' aqui hay que cargar el nombre del articulo (o quizas nada).
    ' .Descripcion = vbNullString
    Descripcion = mobjPedidoCompraItemArticulo.ArticuloColor
    PrecioCompra = mobjPedidoCompraItemArticulo.PrecioCompra
    Descuento = mobjPedidoCompraItemArticulo.Descuento
    Bruto = mobjPedidoCompraItemArticulo.Bruto
    Comision = mobjPedidoCompraItemArticulo.Comision
    
    CantidadT36 = mobjPedidoCompraItemArticulo.PendienteT36
    CantidadT38 = mobjPedidoCompraItemArticulo.PendienteT38
    CantidadT40 = mobjPedidoCompraItemArticulo.PendienteT40
    CantidadT42 = mobjPedidoCompraItemArticulo.PendienteT42
    CantidadT44 = mobjPedidoCompraItemArticulo.PendienteT44
    CantidadT46 = mobjPedidoCompraItemArticulo.PendienteT46
    CantidadT48 = mobjPedidoCompraItemArticulo.PendienteT48
    CantidadT50 = mobjPedidoCompraItemArticulo.PendienteT50
    CantidadT52 = mobjPedidoCompraItemArticulo.PendienteT52
    CantidadT54 = mobjPedidoCompraItemArticulo.PendienteT54
    CantidadT56 = mobjPedidoCompraItemArticulo.PendienteT56
    
    Set mobjPedidoCompraItemArticulo = Nothing
End Sub

'IMPLEMENTAR EL INTERFAZ
Public Sub AlbaranCompraItem_BeginEdit(Moneda As String)

    BeginEdit

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

    If Tipo <> ALBARANCOMPRAITEM_ARTICULO Then _
        Err.Raise vbObjectError + 1001, "AlbaranCompraItemArticulo Load", "No se corresponde el tipo de albaran de compra: " & Tipo & " con el valor que se esperaba : ALBARANCOMPRAITEM_ARTICULO"
    
    Load AlbaranCompraItemID

End Sub

Public Property Get AlbaranCompraItem_Tipo() As AlbaranCompraItemTipos

    AlbaranCompraItem_Tipo = ALBARANCOMPRAITEM_ARTICULO

End Property

Public Property Get AlbaranCompraItem_AlbaranCompraItemID() As Long

    AlbaranCompraItem_AlbaranCompraItemID = AlbaranCompraArticuloID

End Property

Public Property Let AlbaranCompraItem_PedidoCompraItemID(Value As Long)

    PedidoCompraArticuloID = Value

End Property

Public Property Get AlbaranCompraItem_PedidoCompraItemID() As Long

    AlbaranCompraItem_PedidoCompraItemID = PedidoCompraArticuloID

End Property

Public Property Get AlbaranCompraItem_Situacion() As String

    AlbaranCompraItem_Situacion = Situacion
    
End Property

Public Property Let AlbaranCompraItem_Cantidad(Value As Double)
    ' Devolvemos error porque sólo se puede asignar la cantidad si es un material, para artículos hay que indicar la talla.
    Err.Raise vbObjectError + 1001, "No se puede asignar cantidades a una linea de albaran de compra de artículos, hay que indicar la talla."
    
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

    PrecioCompra = Value
  
End Property

Public Property Get AlbaranCompraItem_PrecioCoste() As Double

    AlbaranCompraItem_PrecioCoste = PrecioCompra

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


