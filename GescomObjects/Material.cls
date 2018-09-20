VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Material"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As MaterialProps
Private mudtSave As MaterialProps

Private mMoneda As String

' Aqui declaro el ComboBox de unidades de medida
Private mobjUnidadesMedida As TextList

Private mobjMoviMaterial As MoviMaterial
Event Valid(IsValid As Boolean)

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean

' Flag para determinar si se ha introducido algun movimiento de stock
Private mflgHayMovimiento As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit(Moneda As String)

    If mflgEditing Then Err.Raise 445
  
    mMoneda = Moneda
    
    ' save object state
    LSet mudtSave = mudtprops
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    
    ' restore object state
    LSet mudtprops = mudtSave
  
    ' Si hay movimientos pendientes de aplicar tambien se cancelan
    If mflgHayMovimiento Then mobjMoviMaterial.CancelEdit
  
End Sub

Public Sub ApplyEdit()
    Dim objPersist As MaterialPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.MaterialPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.MaterialID
        mflgNew = True
        mflgDeleted = False

    ElseIf mflgDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        ' save object to database if appropriate
        SetState objPersist.Save(GetState)
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False

    End If

    Set objPersist = Nothing
    mflgDirty = False
    mflgEditing = False
  
    If mflgHayMovimiento Then
        With mobjMoviMaterial
            .PrecioCoste = PrecioCoste
            .MaterialID = MaterialID
            .ApplyEdit
        End With
        Set mobjMoviMaterial = Nothing
        mflgHayMovimiento = False
     
    End If

End Sub

Private Sub Class_Initialize()
  
    mflgNew = True
    Set mobjValid = New BrokenRules
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    '
    mobjValid.RuleBroken "Nombre", True
    mobjValid.RuleBroken "Codigo", True
  
    Set mobjUnidadesMedida = New TextList
    mobjUnidadesMedida.Load "UnidadesMedida"
    mudtprops.UnidadMedida = mobjUnidadesMedida.Key(mobjUnidadesMedida.IndexItem(1))
  
    mflgHayMovimiento = False
  
    mudtprops.StockActual = 0
    mudtprops.StockPendiente = 0
    mudtprops.StockMinimo = 0
    mudtprops.StockMaximo = 0
    mudtprops.PrecioCostePTA = 0
    mudtprops.PrecioCosteEUR = 0
    mudtprops.PrecioPonderadoPTA = 0
    mudtprops.PrecioPonderadoEUR = 0
    ' Por defecto se registra como otros materiales y sin anchura de tela
    mudtprops.AnchuraTela = 0
    mudtprops.TipoMaterial = "O"
    mudtprops.Composicion1 = vbNullString
    mudtprops.Composicion2 = vbNullString
    mudtprops.Composicion3 = vbNullString
    mudtprops.Composicion4 = vbNullString
    mudtprops.PorcComposicion1 = 0
    mudtprops.PorcComposicion2 = 0
    mudtprops.PorcComposicion3 = 0
    mudtprops.PorcComposicion4 = 0
    mudtprops.FechaAlta = Date
  
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

Public Sub Load(MaterialID As Long, Moneda As String)
    Dim objPersist As MaterialPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    mMoneda = Moneda
    
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.MaterialPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(MaterialID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Nombre", False
    mobjValid.RuleBroken "Codigo", False
 
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
'    If mudtProps.StockActual <> 0 Then _
'        Err.Raise vbObjectError + 1001, "Error borrando material", "No se puede eliminar el material, el stock debe ser nulo."
    
    If mudtprops.StockPendiente <> 0 Then _
        Err.Raise vbObjectError + 1001, "Error borrando material", "No se puede eliminar el material, el stock pendiente debe ser nulo."
  
    mflgDeleted = True

End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty

End Property

Public Property Get MaterialID() As Long

    MaterialID = mudtprops.MaterialID
  
End Property

Public Property Let Nombre(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Nombre) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    mobjValid.RuleBroken "Nombre", (Len(Trim(Value)) = 0)

    mudtprops.Nombre = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Nombre() As String

    Nombre = Trim(mudtprops.Nombre)

End Property

Public Property Let Codigo(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Codigo) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    mobjValid.RuleBroken "Codigo", (Len(Trim(Value)) = 0)

    mudtprops.Codigo = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Codigo() As String

    Codigo = Trim(mudtprops.Codigo)

End Property

Public Property Let UnidadMedida(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.UnidadMedida = vbNullString
  
    Select Case Value
        Case UMUnidadesTexto
            mudtprops.UnidadMedida = UMUnidades
        Case UMMetrosTexto
            mudtprops.UnidadMedida = UMMetros
        Case UMCajasTexto
            mudtprops.UnidadMedida = UMCajas
        Case UMKilosTexto
            mudtprops.UnidadMedida = UMKilos
        Case UMGramosTexto
            mudtprops.UnidadMedida = UMGramos
        Case Else
            Err.Raise vbObjectError + 1001, "Valor de unidad de medida no permitido"
    End Select
  
    mobjValid.RuleBroken "UnidadMedida", (mudtprops.UnidadMedida = vbNullString)
  
    mflgDirty = True

End Property

Public Property Get UnidadMedida() As String

    Select Case mudtprops.UnidadMedida
        Case UMUnidades
            UnidadMedida = UMUnidadesTexto
        Case UMMetros
            UnidadMedida = UMMetrosTexto
        Case UMCajas
            UnidadMedida = UMCajasTexto
        Case UMKilos
            UnidadMedida = UMKilosTexto
        Case UMGramos
            UnidadMedida = UMGramosTexto
        Case Else
            UnidadMedida = "Valor desconocido"
    End Select
  
End Property

Public Property Get UnidadesMedida() As TextList

    Set UnidadesMedida = mobjUnidadesMedida

End Property

Public Property Get Moneda() As String

    Moneda = Trim(mMoneda)

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

Public Property Get PrecioPonderado() As Double

    If EsEUR(mMoneda) Then
        PrecioPonderado = mudtprops.PrecioPonderadoEUR
    ElseIf EsPTA(mMoneda) Then
        PrecioPonderado = mudtprops.PrecioPonderadoPTA
    Else
        PrecioPonderado = 0
    End If
    
    If PrecioPonderado = 0 Then
        PrecioPonderado = PrecioCoste
    End If
      
End Property


' StockActual es una propiedad de solo lectura. Se mantiene a traves de metodos.
Public Property Get StockActual() As Double

    StockActual = mudtprops.StockActual
  
End Property

' StockPendiente es una propiedad de solo lectura. Se mantiene a traves de metodos.
Public Property Get StockPendiente() As Double

    StockPendiente = mudtprops.StockPendiente
  
End Property

' StockPrevisto es una propiedad de solo lectura. Se calcula a partir del actual y del pendiente.
Public Property Get StockPrevisto() As Double

    StockPrevisto = mudtprops.StockPendiente + mudtprops.StockActual
  
End Property

Public Property Let StockMinimo(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Stock no puede ser negativo"
  
    ' Se valida que no sean valores negativos y que no sea mayor que el stock maximo.
    mobjValid.RuleBroken "StockMinimo", (Value < 0)
    mobjValid.RuleBroken "RangoStocks", (Value > mudtprops.StockMaximo)
  
    mudtprops.StockMinimo = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get StockMinimo() As Double

    StockMinimo = mudtprops.StockMinimo
  
End Property

Public Property Let StockMaximo(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Stock no puede ser negativo"
  
    ' Se valida que no sean valores negativos y que no sea mayor que el stock maximo.
    mobjValid.RuleBroken "StockMaximo", (Value < 0)
    mobjValid.RuleBroken "RangoStocks", (Value < mudtprops.StockMinimo)
  
    mudtprops.StockMaximo = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get StockMaximo() As Double

    StockMaximo = mudtprops.StockMaximo
  
End Property

Public Property Let AnchuraTela(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value > 100 Or Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    ' pongo la asignacion del valor antes porque entra en el testeo de reglas
    mudtprops.AnchuraTela = Round(Value, 2)
    ' se rompe una regla si es menor de cero.
    mobjValid.RuleBroken "AnchuraTela1", (Value < 0)
    
    mflgDirty = True

End Property

Public Property Get AnchuraEstandar() As Double

    AnchuraEstandar = dblAnchuraTelaEstandar

End Property

Public Property Get AnchuraTela() As Double

    AnchuraTela = mudtprops.AnchuraTela

End Property

Public Property Let TipoMaterial(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Value <> "O" And Value <> "T" Then _
        Err.Raise vbObjectError + 1001, "Tipo de material debe ser 'O' o 'T'"
  
    mudtprops.TipoMaterial = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get TipoMaterial() As String

    TipoMaterial = Trim(mudtprops.TipoMaterial)

End Property

Private Function GetState() As String
    Dim udtData As MaterialData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Public Property Let Composicion1(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) > Len(mudtprops.Composicion1) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    mobjValid.RuleBroken "Composicion1", ((Len(Trim(Value)) = 0) And (mudtprops.PorcComposicion1 <> 0))

    mudtprops.Composicion1 = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Composicion1() As String

    Composicion1 = Trim(mudtprops.Composicion1)

End Property

Public Property Let PorcComposicion1(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value > 100 Or Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    ' pongo la asignacion del valor antes porque entra en el testeo de reglas
    mudtprops.PorcComposicion1 = Round(Value, 2)
    ' se rompe una regla si es mayor de cien, si es menor de cero o si la composicion no esta informada.
    ' tampoco se permite que haya un total mayor de cien%.
    mobjValid.RuleBroken "PorcComposicion1", _
        ((Value < 0) Or (Value > 100) Or (Len(Trim(mudtprops.Composicion1)) = 0) Or _
        ((mudtprops.PorcComposicion1 + mudtprops.PorcComposicion2 + mudtprops.PorcComposicion3 + mudtprops.PorcComposicion4) > 100) _
        )

    mflgDirty = True

End Property

Public Property Get PorcComposicion1() As Double

    PorcComposicion1 = mudtprops.PorcComposicion1

End Property

Public Property Let Composicion2(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) > Len(mudtprops.Composicion2) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    mobjValid.RuleBroken "Composicion2", ((Len(Trim(Value)) = 0) And (mudtprops.PorcComposicion2 <> 0))

    mudtprops.Composicion2 = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Composicion2() As String

    Composicion2 = Trim(mudtprops.Composicion2)

End Property

Public Property Let PorcComposicion2(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value > 100 Or Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    ' pongo la asignacion del valor antes porque entra en el testeo de reglas
    mudtprops.PorcComposicion2 = Round(Value, 2)
    ' se rompe una regla si es mayor de cien, si es menor de cero o si la composicion no esta informada.
    ' tampoco se permite que haya un total mayor de cien%.
    mobjValid.RuleBroken "PorcComposicion2", _
        ((Value < 0) Or (Value > 100) Or (Len(Trim(mudtprops.Composicion2)) = 0) Or _
        ((mudtprops.PorcComposicion1 + mudtprops.PorcComposicion2 + mudtprops.PorcComposicion3 + mudtprops.PorcComposicion4) > 100) _
        )

    mflgDirty = True

End Property

Public Property Get PorcComposicion2() As Double

    PorcComposicion2 = mudtprops.PorcComposicion2

End Property

Public Property Let Composicion3(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) > Len(mudtprops.Composicion3) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    mobjValid.RuleBroken "Composicion3", ((Len(Trim(Value)) = 0) And (mudtprops.PorcComposicion3 <> 0))

    mudtprops.Composicion3 = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Composicion3() As String

    Composicion3 = Trim(mudtprops.Composicion3)

End Property

Public Property Let PorcComposicion3(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value > 100 Or Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    ' pongo la asignacion del valor antes porque entra en el testeo de reglas
    mudtprops.PorcComposicion3 = Round(Value, 2)
      
    ' se rompe una regla si es mayor de cien, si es menor de cero o si la composicion no esta informada.
    ' tampoco se permite que haya un total mayor de cien%.
    mobjValid.RuleBroken "PorcComposicion3", _
        ((Value < 0) Or (Value > 100) Or (Len(Trim(mudtprops.Composicion3)) = 0) Or _
        ((mudtprops.PorcComposicion1 + mudtprops.PorcComposicion2 + mudtprops.PorcComposicion3 + mudtprops.PorcComposicion4) > 100) _
        )
    
    mflgDirty = True

End Property

Public Property Get PorcComposicion3() As Double

    PorcComposicion3 = Trim(mudtprops.PorcComposicion3)

End Property

Public Property Let Composicion4(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) > Len(mudtprops.Composicion4) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    mobjValid.RuleBroken "Composicion4", ((Len(Trim(Value)) = 0) And (mudtprops.PorcComposicion4 <> 0))

    mudtprops.Composicion4 = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Composicion4() As String

    Composicion4 = Trim(mudtprops.Composicion4)

End Property

Public Property Let PorcComposicion4(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value > 100 Or Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    ' pongo la asignacion del valor antes porque entra en el testeo de reglas
    mudtprops.PorcComposicion4 = Round(Value, 2)
    ' se rompe una regla si es mayor de cien, si es menor de cero o si la composicion no esta informada.
    ' tampoco se permite que haya un total mayor de cien%.
    mobjValid.RuleBroken "PorcComposicion4", _
        ((Value < 0) Or (Value > 100) Or (Len(Trim(mudtprops.Composicion4)) = 0) Or _
        ((mudtprops.PorcComposicion1 + mudtprops.PorcComposicion2 + mudtprops.PorcComposicion3 + mudtprops.PorcComposicion4) > 100) _
        )

    mflgDirty = True

End Property

Public Property Get PorcComposicion4() As Double

    PorcComposicion4 = mudtprops.PorcComposicion4

End Property

Public Property Let FechaAlta(Value As Variant)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) = 0 Then
        mudtprops.FechaAlta = vbNullString
        mobjValid.RuleBroken "FechaAlta", True

    ElseIf IsDate(Value) Then
        mudtprops.FechaAlta = CDate(Value)
        mobjValid.RuleBroken "FechaAlta", False

    Else
        mobjValid.RuleBroken "FechaAlta", True

    End If

    mflgDirty = True

End Property

Public Property Get FechaAlta() As Variant

    With mudtprops
        FechaAlta = IIf(.FechaAlta = 0, vbNullString, .FechaAlta)
    End With

End Property


Private Sub SetState(Buffer As String)
    Dim udtData As MaterialData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub

' Registra un movimiento de inventario
' No tiene que haber movimientos realizados.
Public Sub StockInventario(Cantidad As Double, Concepto As String, Precio As Double, _
                           DocumentoID As Long, TipoDocumento As String)
    
    If mflgHayMovimiento Then Err.Raise 383
    If mflgNew Then Err.Raise 383
    
    Set mobjMoviMaterial = New MoviMaterial
    
    With mobjMoviMaterial
        .BeginEdit
        .Cantidad = Cantidad
        .Tipo = TMMInventarioTexto
        .Concepto = Concepto
        mudtprops.StockActual = Cantidad
        .StockFinal = mudtprops.StockActual
        .Precio = Precio
        .DocumentoID = DocumentoID
        .TipoDocumento = TipoDocumento
    End With
    
    mflgHayMovimiento = True
    mflgDirty = True
    
End Sub

' Registra un movimiento de Entrada
' No tiene que haber movimientos realizados.
Public Sub StockEntrada(Cantidad As Double, Concepto As String, Precio As Double, _
                           DocumentoID As Long, TipoDocumento As String)
    
    If mflgHayMovimiento Then Err.Raise 383
    If mflgNew Then Err.Raise 383
    
    Set mobjMoviMaterial = New MoviMaterial
    
    With mobjMoviMaterial
        .BeginEdit
        .Cantidad = Cantidad
        .Tipo = TMMEntradaTexto
        .Concepto = Concepto
        mudtprops.StockActual = mudtprops.StockActual + Cantidad
        .StockFinal = mudtprops.StockActual
        .Precio = Precio
        .DocumentoID = DocumentoID
        .TipoDocumento = TipoDocumento
    End With
    
    mflgHayMovimiento = True
    mflgDirty = True
    
End Sub

' Registra un movimiento de Salida
' No tiene que haber movimientos realizados.
Public Sub StockSalida(Cantidad As Double, Concepto As String, Precio As Double, _
                           DocumentoID As Long, TipoDocumento As String)
                           
    If mflgHayMovimiento Then Err.Raise 383
    If mflgNew Then Err.Raise 383
    
    Set mobjMoviMaterial = New MoviMaterial
    
    With mobjMoviMaterial
        .BeginEdit
        .Cantidad = Cantidad
        .Tipo = TMMSalidaTexto
        .Concepto = Concepto
        mudtprops.StockActual = mudtprops.StockActual - Cantidad
        .StockFinal = mudtprops.StockActual
        .Precio = Precio
        .DocumentoID = DocumentoID
        .TipoDocumento = TipoDocumento
    End With
    
    mflgHayMovimiento = True
    mflgDirty = True
    
End Sub

' Registra un movimiento de Entrega de material (recepcion)
' No tiene que haber movimientos realizados.
Public Sub StockEntrega(Cantidad As Double, Concepto As String, Precio As Double)
                            ', _
                           'DocumentoID As Long, TipoDocumento As String)
    
    If mflgHayMovimiento Then Err.Raise 383
    If mflgNew Then Err.Raise 383
    
    Set mobjMoviMaterial = New MoviMaterial
    
    With mobjMoviMaterial
        .BeginEdit
        .Cantidad = Cantidad
        .Tipo = TMMEntregaTexto
        .Concepto = Concepto
        mudtprops.StockPendiente = mudtprops.StockPendiente - Cantidad
        .StockFinal = mudtprops.StockPendiente
        .Precio = Precio
    End With
    
    mflgHayMovimiento = True
    mflgDirty = True
    
End Sub

' Registra un movimiento de Reserva de material (entrega pendiente)
' No tiene que haber movimientos realizados.
Public Sub StockReserva(Cantidad As Double, Concepto As String, Precio As Double, _
                           DocumentoID As Long, TipoDocumento As String)
    
    If mflgHayMovimiento Then Err.Raise 383
    If mflgNew Then Err.Raise 383
    
    Set mobjMoviMaterial = New MoviMaterial
    
    With mobjMoviMaterial
        .BeginEdit
        .Cantidad = Cantidad
        .Tipo = TMMReservaTexto
        .Concepto = Concepto
        mudtprops.StockPendiente = mudtprops.StockPendiente + Cantidad
        .StockFinal = mudtprops.StockPendiente
        .Precio = Precio
        .DocumentoID = DocumentoID
        .TipoDocumento = TipoDocumento
    End With
    
    mflgHayMovimiento = True
    mflgDirty = True
    
End Sub
