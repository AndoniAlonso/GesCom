VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PedidoVentaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private mobjArticuloColor As ArticuloColor

Private mMoneda As String

Private mstrCodigoArticuloColor As String

Private mudtprops As PedidoVentaItemProps
Private mudtSave As PedidoVentaItemProps
Private mudtChild As PedidoVentaItemProps
Private mobjParent As PedidoVentaItems
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
    Dim objPersist As PedidoVentaItemPersist
    
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    Set objPersist = CreateObject("GCServerMTS.PedidoVentaItemPersist", PERSIST_SERVER)
    
    If Not mflgChild Then
        If mflgDeleted And Not mflgNew Then
            ' code to delete the object's data goes here
            objPersist.DeleteObject mudtprops.PedidoVentaItemID
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
            mobjParent.AddPedidoVentaItem Me
            Set mobjParent = Nothing
        End If
    
    End If
      
    Set objPersist = Nothing
    mflgEditing = False
    
End Sub

Private Sub Class_Initialize()
    
    mflgNew = True
    Set mobjValid = New BrokenRules
    
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    mobjValid.RuleBroken "ArticuloColorID", True
    mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "Cantidad", True
    
    With mudtprops
        .Situacion = "A"    ' Marco la situacion del pedido como "A"--> alta,
                            ' pendiente de codificar las posibles situaciones.
        .SituacionCorte = "A"   ' Marco la situacion de corte del pedido como "A"--> alta,
                                ' pendiente de codificar las posibles situaciones.
        .ArticuloColorID = 0
        .NombreArticuloColor = vbNullString
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
        .ServidoT36 = 0
        .ServidoT38 = 0
        .ServidoT40 = 0
        .ServidoT42 = 0
        .ServidoT44 = 0
        .ServidoT46 = 0
        .ServidoT48 = 0
        .ServidoT50 = 0
        .ServidoT52 = 0
        .ServidoT54 = 0
        .ServidoT56 = 0
        .PrecioVentaPTA = 0
        .PrecioVentaEUR = 0
        .Descuento = 0
        .BrutoPTA = 0
        .BrutoEUR = 0
        .Comision = 0
        .TemporadaID = 0
        .Observaciones = vbNullString
        .ActualizarAlta = False
        .DesactualizarAlta = False
        .ActualizarAlbaran = False
        .DesactualizarAlbaran = False
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

Public Sub Load(PedidoVentaItemID As Long, Moneda As String)
    Dim objPersist As PedidoVentaItemPersist
    
    If mflgChild Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    
    mMoneda = Moneda
    
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.PedidoVentaItemPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(PedidoVentaItemID)
    Set objPersist = Nothing
    
    mobjValid.RuleBroken "ArticuloColorID", False
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "Cantidad", False
    
    mflgNew = False
    
End Sub

Public Sub Delete()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    If mudtprops.SituacionCorte = "C" Then _
        Err.Raise vbObjectError + 1001, "No se puede borrar, ya se ha generado orden de corte"
    
    If Servido Then _
        Err.Raise vbObjectError + 1001, "No se puede borrar, ya se han servido entregas"
    
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

Friend Sub SetAsChild(objPedidoVentaItems As PedidoVentaItems, Moneda As String)

    mMoneda = Moneda
    If mflgNew Then Set mobjParent = objPedidoVentaItems
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

Public Property Get PedidoVentaItemID() As Long

    PedidoVentaItemID = mudtprops.PedidoVentaItemID

End Property

Public Property Get PedidoVentaID() As Long

    PedidoVentaID = mudtprops.PedidoVentaID

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

Public Property Get SituacionCorte() As String
    
    Select Case mudtprops.SituacionCorte
    Case "A"
        SituacionCorte = "Alta"
    Case "C"
        SituacionCorte = "Mandado a Orden de corte"
    Case Else
        Err.Raise vbObjectError + 1001, "Situacion de corte erronea"
    End Select
    
End Property

Public Property Get OrdenCorte() As Double
  
    OrdenCorte = (mudtprops.SituacionCorte = "C")

End Property

Public Property Get ArticuloColorID() As Long

    ArticuloColorID = mudtprops.ArticuloColorID
  
End Property

Public Property Let CodigoArticuloColor(Value As String)

    If Not mflgEditing Then Err.Raise 383
    
    If Len(Trim(Value)) > 8 Then Exit Property
    
    mstrCodigoArticuloColor = UCase(Trim(Value))
     
    If Len(mstrCodigoArticuloColor) < 8 Then Exit Property
    
    Set mobjArticuloColor = New ArticuloColor
    
    mobjArticuloColor.LoadCodigo Trim(mstrCodigoArticuloColor), mudtprops.TemporadaID
    
    mobjValid.RuleBroken "ArticuloColorID", False

    PrecioVenta = mobjArticuloColor.objArticulo.PrecioVenta
    'mudtprops.NombreArticuloColor = mobjArticuloColor.Nombre
    mudtprops.NombreArticuloColor = mobjArticuloColor.objArticulo.NombreCompleto & " " & mobjArticuloColor.NombreColor

    mudtprops.ArticuloColorID = mobjArticuloColor.ArticuloColorID
 
    mflgDirty = True
 
End Property

Public Property Get CodigoArticuloColor() As String
    
    CodigoArticuloColor = mstrCodigoArticuloColor
    
End Property

Public Property Get NombreArticuloColor() As String
    
    NombreArticuloColor = mudtprops.NombreArticuloColor
    
End Property

Public Property Get objArticuloColor() As ArticuloColor
    
    If mudtprops.ArticuloColorID = 0 Then
        Set objArticuloColor = Nothing
        Exit Property
    End If
    
    If mobjArticuloColor Is Nothing Then _
        Set mobjArticuloColor = New ArticuloColor
        
    If mudtprops.ArticuloColorID <> mobjArticuloColor.ArticuloColorID Then
        mobjArticuloColor.Load mudtprops.ArticuloColorID, mMoneda
    End If
    
    Set objArticuloColor = mobjArticuloColor

End Property

Public Property Get ArticuloColor() As String

    ArticuloColor = Trim(mudtprops.NombreArticuloColor)

End Property

Public Property Let CantidadT36(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If Servido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
    If mudtprops.SituacionCorte = "C" Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT36", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT36 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
    
End Property

Public Property Get CantidadT36() As Double

    CantidadT36 = mudtprops.CantidadT36

End Property

Public Property Let CantidadT38(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If Servido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
    If mudtprops.SituacionCorte = "C" Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT38", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT38 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
    
End Property

Public Property Get CantidadT38() As Double

    CantidadT38 = mudtprops.CantidadT38

End Property

Public Property Let CantidadT40(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If Servido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
    If mudtprops.SituacionCorte = "C" Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT40", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT40 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
    
End Property

Public Property Get CantidadT40() As Double

    CantidadT40 = mudtprops.CantidadT40

End Property

Public Property Let CantidadT42(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If Servido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
    If mudtprops.SituacionCorte = "C" Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT42", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT42 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
    
End Property

Public Property Get CantidadT42() As Double

    CantidadT42 = mudtprops.CantidadT42

End Property

Public Property Let CantidadT44(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If Servido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
    If mudtprops.SituacionCorte = "C" Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT44", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT44 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
    
End Property

Public Property Get CantidadT44() As Double

    CantidadT44 = mudtprops.CantidadT44

End Property

Public Property Let CantidadT46(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If Servido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
    If mudtprops.SituacionCorte = "C" Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT46", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT46 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
    
End Property

Public Property Get CantidadT46() As Double

    CantidadT46 = mudtprops.CantidadT46

End Property

Public Property Let CantidadT48(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If Servido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
    If mudtprops.SituacionCorte = "C" Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT48", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT48 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
    
End Property

Public Property Get CantidadT48() As Double

    CantidadT48 = mudtprops.CantidadT48

End Property

Public Property Let CantidadT50(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If Servido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
    If mudtprops.SituacionCorte = "C" Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT50", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT50 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
    
End Property

Public Property Get CantidadT50() As Double

    CantidadT50 = mudtprops.CantidadT50

End Property

Public Property Let CantidadT52(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If Servido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
    If mudtprops.SituacionCorte = "C" Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT52", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT52 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
    
End Property

Public Property Get CantidadT52() As Double

    CantidadT52 = mudtprops.CantidadT52

End Property

Public Property Let CantidadT54(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If Servido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
    If mudtprops.SituacionCorte = "C" Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT54", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT54 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
    
End Property

Public Property Get CantidadT54() As Double

    CantidadT54 = mudtprops.CantidadT54

End Property

Public Property Let CantidadT56(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If Servido Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
    If mudtprops.SituacionCorte = "C" Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT56", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT56 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
    
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
Public Property Let ServidoT36(Value As Double)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
  ' No se permiten Cantidades negativas ni que sea mayor que lo pedido
  mobjValid.RuleBroken "ServidoT36", (Value < 0) Or (Value > mudtprops.CantidadT36)
  
  ' Las cantidades son enteras
  mudtprops.ServidoT36 = Round(Value, 0)
  
  mflgDirty = True

End Property

Public Property Get ServidoT36() As Double

  ServidoT36 = mudtprops.ServidoT36

End Property
Public Property Let ServidoT38(Value As Double)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
  ' No se permiten Cantidades negativas ni que sea mayor que lo pedido
  mobjValid.RuleBroken "ServidoT38", (Value < 0) Or (Value > mudtprops.CantidadT38)
  
  ' Las cantidades son enteras
  mudtprops.ServidoT38 = Round(Value, 0)
  
  mflgDirty = True

End Property

Public Property Get ServidoT38() As Double

  ServidoT38 = mudtprops.ServidoT38

End Property
Public Property Let ServidoT40(Value As Double)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
  ' No se permiten Cantidades negativas ni que sea mayor que lo pedido
  mobjValid.RuleBroken "ServidoT40", (Value < 0) Or (Value > mudtprops.CantidadT40)
  
  ' Las cantidades son enteras
  mudtprops.ServidoT40 = Round(Value, 0)
  
  mflgDirty = True

End Property

Public Property Get ServidoT40() As Double

  ServidoT40 = mudtprops.ServidoT40

End Property
Public Property Let ServidoT42(Value As Double)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
  ' No se permiten Cantidades negativas ni que sea mayor que lo pedido
  mobjValid.RuleBroken "ServidoT42", (Value < 0) Or (Value > mudtprops.CantidadT42)
  
  ' Las cantidades son enteras
  mudtprops.ServidoT42 = Round(Value, 0)
  
  mflgDirty = True

End Property

Public Property Get ServidoT42() As Double

  ServidoT42 = mudtprops.ServidoT42

End Property
Public Property Let ServidoT44(Value As Double)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
  ' No se permiten Cantidades negativas ni que sea mayor que lo pedido
  mobjValid.RuleBroken "ServidoT44", (Value < 0) Or (Value > mudtprops.CantidadT44)
  
  ' Las cantidades son enteras
  mudtprops.ServidoT44 = Round(Value, 0)
  
  mflgDirty = True

End Property

Public Property Get ServidoT44() As Double

  ServidoT44 = mudtprops.ServidoT44

End Property
Public Property Let ServidoT46(Value As Double)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
  ' No se permiten Cantidades negativas ni que sea mayor que lo pedido
  mobjValid.RuleBroken "ServidoT46", (Value < 0) Or (Value > mudtprops.CantidadT46)
  
  ' Las cantidades son enteras
  mudtprops.ServidoT46 = Round(Value, 0)
  
  mflgDirty = True

End Property

Public Property Get ServidoT46() As Double

  ServidoT46 = mudtprops.ServidoT46

End Property
Public Property Let ServidoT48(Value As Double)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
  ' No se permiten Cantidades negativas ni que sea mayor que lo pedido
  mobjValid.RuleBroken "ServidoT48", (Value < 0) Or (Value > mudtprops.CantidadT48)
  
  ' Las cantidades son enteras
  mudtprops.ServidoT48 = Round(Value, 0)
  
  mflgDirty = True

End Property

Public Property Get ServidoT48() As Double

  ServidoT48 = mudtprops.ServidoT48

End Property
Public Property Let ServidoT50(Value As Double)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
  ' No se permiten Cantidades negativas ni que sea mayor que lo pedido
  mobjValid.RuleBroken "ServidoT50", (Value < 0) Or (Value > mudtprops.CantidadT50)
  
  ' Las cantidades son enteras
  mudtprops.ServidoT50 = Round(Value, 0)
  
  mflgDirty = True

End Property

Public Property Get ServidoT50() As Double

  ServidoT50 = mudtprops.ServidoT50

End Property
Public Property Let ServidoT52(Value As Double)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
  ' No se permiten Cantidades negativas ni que sea mayor que lo pedido
  mobjValid.RuleBroken "ServidoT52", (Value < 0) Or (Value > mudtprops.CantidadT52)
  
  ' Las cantidades son enteras
  mudtprops.ServidoT52 = Round(Value, 0)
  
  mflgDirty = True

End Property

Public Property Get ServidoT52() As Double

  ServidoT52 = mudtprops.ServidoT52

End Property
Public Property Let ServidoT54(Value As Double)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
  ' No se permiten Cantidades negativas ni que sea mayor que lo pedido
  mobjValid.RuleBroken "ServidoT54", (Value < 0) Or (Value > mudtprops.CantidadT54)
  
  ' Las cantidades son enteras
  mudtprops.ServidoT54 = Round(Value, 0)
  
  mflgDirty = True

End Property

Public Property Get ServidoT54() As Double

  ServidoT54 = mudtprops.ServidoT54

End Property
Public Property Let ServidoT56(Value As Double)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
  ' No se permiten Cantidades negativas ni que sea mayor que lo pedido
  mobjValid.RuleBroken "ServidoT56", (Value < 0) Or (Value > mudtprops.CantidadT56)
  
  ' Las cantidades son enteras
  mudtprops.ServidoT56 = Round(Value, 0)
  
  mflgDirty = True

End Property

Public Property Get ServidoT56() As Double

  ServidoT56 = mudtprops.ServidoT56

End Property
Public Property Get Servido() As Double
  With mudtprops
  Servido = .ServidoT36 + .ServidoT38 + .ServidoT40 + .ServidoT42 + _
             .ServidoT44 + .ServidoT46 + .ServidoT48 + .ServidoT50 + _
             .ServidoT52 + .ServidoT54 + .ServidoT56
  End With

End Property

Public Property Get PendienteT36() As Double

  PendienteT36 = mudtprops.CantidadT36 - mudtprops.ServidoT36

End Property
Public Property Get PendienteT38() As Double

  PendienteT38 = mudtprops.CantidadT38 - mudtprops.ServidoT38

End Property
Public Property Get PendienteT40() As Double

  PendienteT40 = mudtprops.CantidadT40 - mudtprops.ServidoT40

End Property
Public Property Get PendienteT42() As Double

  PendienteT42 = mudtprops.CantidadT42 - mudtprops.ServidoT42

End Property
Public Property Get PendienteT44() As Double

  PendienteT44 = mudtprops.CantidadT44 - mudtprops.ServidoT44

End Property
Public Property Get PendienteT46() As Double

  PendienteT46 = mudtprops.CantidadT46 - mudtprops.ServidoT46

End Property
Public Property Get PendienteT48() As Double

  PendienteT48 = mudtprops.CantidadT48 - mudtprops.ServidoT48

End Property
Public Property Get PendienteT50() As Double

  PendienteT50 = mudtprops.CantidadT50 - mudtprops.ServidoT50

End Property
Public Property Get PendienteT52() As Double

  PendienteT52 = mudtprops.CantidadT52 - mudtprops.ServidoT52

End Property

Public Property Get PendienteT54() As Double

    PendienteT54 = mudtprops.CantidadT54 - mudtprops.ServidoT54

End Property

Public Property Get PendienteT56() As Double

    PendienteT56 = mudtprops.CantidadT56 - mudtprops.ServidoT56

End Property

Public Property Get Pendiente() As Double

    Pendiente = Cantidad - Servido

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
  
  If EsEUR(mMoneda) Then
     mudtprops.PrecioVentaEUR = Round(Value, 2)
     mudtprops.PrecioVentaPTA = EUR2PTA(Value)
  ElseIf EsPTA(mMoneda) Then
     mudtprops.PrecioVentaPTA = Round(Value, 0)
     mudtprops.PrecioVentaEUR = PTA2EUR(Value)
  Else
     mudtprops.PrecioVentaPTA = 0
     mudtprops.PrecioVentaEUR = 0
  End If
  
  CalcularBruto
    
  mflgDirty = True

End Property
Public Property Get PrecioVenta() As Double

  If EsEUR(mMoneda) Then
     PrecioVenta = mudtprops.PrecioVentaEUR
  ElseIf EsPTA(mMoneda) Then
     PrecioVenta = mudtprops.PrecioVentaPTA
  Else
     PrecioVenta = 0
  End If
  
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

Public Property Let Observaciones(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Observaciones) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.Observaciones = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Observaciones() As String

    Observaciones = Trim(mudtprops.Observaciones)

End Property

Friend Function GetState() As String

  Dim udtData As PedidoVentaItemData
  
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
    Dim udtData As PedidoVentaItemData
    
    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    mobjValid.RuleBroken "ArticuloColorID", False
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "Cantidad", False
    mstrCodigoArticuloColor = Left(mudtprops.NombreArticuloColor, 8)
    
End Sub

Private Sub CalcularBruto()
    Dim curImporteSinDescuento As Currency
        
    curImporteSinDescuento = PrecioVenta * Cantidad
    
    Bruto = curImporteSinDescuento - (curImporteSinDescuento * Descuento / 100)

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
