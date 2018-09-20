VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArticuloColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As ArticuloColorProps
Private mudtSave As ArticuloColorProps

' Aqui declaro el ComboBox de Articulos
Private mobjArticulos As TextList

' Aqui declaro el ComboBox de unidades de medida
Private mobjUnidadesMedida As TextList

Private mobjMoviArticulo As MoviArticulo

Event Valid(IsValid As Boolean)

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean

' Flag para determinar si se ha introducido algun movimiento de stock
Private mflgHayMovimiento As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit()

    If mflgEditing Then Err.Raise 445
  
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
    If mflgHayMovimiento Then mobjMoviArticulo.CancelEdit
  
End Sub

Public Sub ApplyEdit()

    Dim objPersist As ArticuloColorPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.ArticuloColorPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.ArticuloColorID
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
  
    If mflgHayMovimiento Then MovimientoArticulo
 
End Sub

Private Sub Class_Initialize()
    
    mflgNew = True
    Set mobjValid = New BrokenRules
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    '
    mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "NombreColor", True
    mobjValid.RuleBroken "ArticuloID", True
    mobjValid.RuleBroken "Codigo", True
     
    ' Relleno la lista de Articulos y series cuando asigne la temporada
    ' comento esto porque no tiene sentido, aunque quizas de problemas
'    Set mobjArticulos = New TextList
'    mobjArticulos.Load "vNombreArticulos", , mudtProps.TemporadaID
'    mudtProps.ArticuloID = mobjArticulos.Key(mobjArticulos.IndexItem(1))
'
    mudtprops.Nombre = vbNullString
    mudtprops.NombreColor = vbNullString
    mudtprops.StockActualT36 = 0
    mudtprops.StockActualT38 = 0
    mudtprops.StockActualT40 = 0
    mudtprops.StockActualT42 = 0
    mudtprops.StockActualT44 = 0
    mudtprops.StockActualT46 = 0
    mudtprops.StockActualT48 = 0
    mudtprops.StockActualT50 = 0
    mudtprops.StockActualT52 = 0
    mudtprops.StockActualT54 = 0
    mudtprops.StockActualT56 = 0
    mudtprops.StockPendienteT36 = 0
    mudtprops.StockPendienteT38 = 0
    mudtprops.StockPendienteT40 = 0
    mudtprops.StockPendienteT42 = 0
    mudtprops.StockPendienteT44 = 0
    mudtprops.StockPendienteT46 = 0
    mudtprops.StockPendienteT48 = 0
    mudtprops.StockPendienteT50 = 0
    mudtprops.StockPendienteT52 = 0
    mudtprops.StockPendienteT54 = 0
    mudtprops.StockPendienteT56 = 0
      
    mudtprops.TemporadaID = 0
      
End Sub

Private Sub Class_Terminate()
    
    Set mobjValid = Nothing
    Set mobjArticulos = Nothing
    Set mobjUnidadesMedida = Nothing
    Set mobjMoviArticulo = Nothing
      
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

Public Sub Load(ArticuloColorID As Long, Optional Codigo As String, _
                Optional TemporadaID As Long)

    Dim objPersist As ArticuloColorPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.ArticuloColorPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(ArticuloColorID, Codigo, TemporadaID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "NombreColor", False
    mobjValid.RuleBroken "ArticuloID", False
    mobjValid.RuleBroken "Codigo", False
    
    Set mobjArticulos = Nothing
'    Set mobjArticulos = New TextList
'    mobjArticulos.Load "vNombreArticulos", , mudtprops.TemporadaID
    
    mflgNew = False

End Sub

Public Sub LoadCodigo(Codigo As String, TemporadaID As Long)
    Dim mintArticuloID As Long
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' Si el codigo buscado existe se asigna como esté en la tabla.
    If ExisteCodigo(Codigo, TemporadaID) Then
        ' El articulo-color existe y se ha establecido
        
        ' Si no existe hay que crearlo en base al articulo base.
    Else
        mintArticuloID = CodigoArticuloID(Mid(Codigo, 1, 6), TemporadaID)
        If mintArticuloID <> 0 Then
            Me.BeginEdit
            Me.TemporadaID = TemporadaID
            mudtprops.ArticuloID = mintArticuloID
            mobjValid.RuleBroken "ArticuloID", False
            Me.Codigo = Mid$(Codigo, 7, 2)
            mudtprops.NombreColor = Mid(Codigo, 7)
            Me.ApplyEdit
        Else
            Err.Raise vbObjectError + 1001, Err.Source, Err.Description
        End If
    End If

    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "NombreColor", False
    mobjValid.RuleBroken "ArticuloID", False
    mobjValid.RuleBroken "Codigo", False
    
    Set mobjArticulos = Nothing
'    Set mobjArticulos = New TextList
'    mobjArticulos.Load "vNombreArticulos", , mudtprops.TemporadaID
    
    mflgNew = False

End Sub

'Devuelve el identificador del articulo que queremos comprobar si existe.
Private Function CodigoArticuloID(Codigo As String, _
                                TemporadaID As Long) As Long
    Dim objMyArticulo As Articulo
    
    Set objMyArticulo = New Articulo
    objMyArticulo.LoadCodigo Codigo, TemporadaID
    
    CodigoArticuloID = objMyArticulo.ArticuloID
    Set objMyArticulo = Nothing
    
End Function

'Private Function ExisteCodigo(Codigo As String, TemporadaID As Long) As Boolean
Public Property Get ExisteCodigo(Codigo As String, TemporadaID As Long) As Boolean
    Dim objPersist As ArticuloColorPersist
    
    On Error GoTo ErrorManager

    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.ArticuloColorPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(0, Codigo, TemporadaID)
    Set objPersist = Nothing
    
    ExisteCodigo = True
    
    Exit Property

ErrorManager:
    ExisteCodigo = False
End Property

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
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

Public Property Get ArticuloColorID() As Long

    ArticuloColorID = mudtprops.ArticuloColorID
  
End Property

Public Property Get Nombre() As String

    Nombre = Trim(mudtprops.Nombre)

End Property

Public Property Get Codigo() As String
' OJO, aquí habría que asegurarse de que 7 y 2 no van a ser modificadas
' si varía la manera de codificar articulos.
    Codigo = Mid(Trim(mudtprops.Nombre), 7, 2)

End Property

Public Property Let Codigo(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.NombreColor) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    ' exijo que haya dos caracteres para que pueda componer el nombre
    mobjValid.RuleBroken "Codigo", (Len(Trim(Value)) <> 2)

    If NombreColor = Codigo Then NombreColor = Trim(UCase$(Value))
    mudtprops.Nombre = Trim(objArticulo.Nombre) & Trim(UCase$(Value))
    
    mflgDirty = True

End Property

Public Property Let NombreColor(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.NombreColor) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    ' exijo que no sea nulo
    mobjValid.RuleBroken "NombreColor", (Len(Trim(Value)) = 0)

    mudtprops.NombreColor = Trim(Value)
    mflgDirty = True

End Property

Public Property Get NombreColor() As String

    NombreColor = Trim(mudtprops.NombreColor)

End Property

Public Property Get NombreCompleto() As String

    NombreCompleto = Me.Nombre & "-" & objArticulo.NombreCompleto & " " & Me.NombreColor

End Property

Public Property Let ArticuloID(Value As Long)

    If Not mflgEditing Then Err.Raise 383
    If Not mflgNew Then Err.Raise 383
  
    mudtprops.ArticuloID = Value
  
    mobjValid.RuleBroken "ArticuloID", (mudtprops.ArticuloID = 0)
  
    mflgDirty = True

End Property

Public Property Let Articulo(Value As String)

    If Not mflgEditing Then Err.Raise 383
    If Not mflgNew Then Err.Raise 383
  
    If mobjArticulos Is Nothing Then CargarListaArticulos
    mudtprops.ArticuloID = mobjArticulos.Key(Value)
  
    mobjValid.RuleBroken "ArticuloID", (mudtprops.ArticuloID = 0)
  
    mflgDirty = True

End Property

Public Property Get Articulo() As String

    If mobjArticulos Is Nothing Then CargarListaArticulos
    Articulo = mobjArticulos.Item(mudtprops.ArticuloID)

End Property

Public Property Get Articulos() As TextList

    If mobjArticulos Is Nothing Then CargarListaArticulos
    Set Articulos = mobjArticulos

End Property

Public Property Get StockActual() As Double

    With mudtprops
        StockActual = .StockActualT36 + .StockActualT38 + .StockActualT40 + _
            .StockActualT42 + .StockActualT44 + .StockActualT46 + _
            .StockActualT48 + .StockActualT50 + .StockActualT52 + _
            .StockActualT54 + .StockActualT56
    End With
  
End Property

Public Property Get StockActualT36() As Double

    StockActualT36 = mudtprops.StockActualT36
    
End Property

Public Property Get StockActualT38() As Double

    StockActualT38 = mudtprops.StockActualT38
    
End Property

Public Property Get StockActualT40() As Double

    StockActualT40 = mudtprops.StockActualT40
    
End Property

Public Property Get StockActualT42() As Double

    StockActualT42 = mudtprops.StockActualT42
    
End Property

Public Property Get StockActualT44() As Double
    
    StockActualT44 = mudtprops.StockActualT44
    
End Property

Public Property Get StockActualT46() As Double

    StockActualT46 = mudtprops.StockActualT46
    
End Property

Public Property Get StockActualT48() As Double

    StockActualT48 = mudtprops.StockActualT48
    
End Property

Public Property Get StockActualT50() As Double

    StockActualT50 = mudtprops.StockActualT50
    
End Property

Public Property Get StockActualT52() As Double

    StockActualT52 = mudtprops.StockActualT52
    
End Property

Public Property Get StockActualT54() As Double
    
    StockActualT54 = mudtprops.StockActualT54
    
End Property

Public Property Get StockActualT56() As Double

    StockActualT56 = mudtprops.StockActualT56

End Property

Public Property Get StockPendiente() As Double

    With mudtprops
        StockPendiente = .StockPendienteT36 + .StockPendienteT38 + .StockPendienteT40 + _
            .StockPendienteT42 + .StockPendienteT44 + .StockPendienteT46 + _
            .StockPendienteT48 + .StockPendienteT50 + .StockPendienteT52 + _
            .StockPendienteT54 + .StockPendienteT56
    End With
  
End Property

Public Property Get StockPendienteT36() As Double

    StockPendienteT36 = mudtprops.StockPendienteT36
    
End Property

Public Property Get StockPendienteT38() As Double

    StockPendienteT38 = mudtprops.StockPendienteT38

End Property

Public Property Get StockPendienteT40() As Double

    StockPendienteT40 = mudtprops.StockPendienteT40

End Property

Public Property Get StockPendienteT42() As Double

    StockPendienteT42 = mudtprops.StockPendienteT42

End Property

Public Property Get StockPendienteT44() As Double

    StockPendienteT44 = mudtprops.StockPendienteT44
    
End Property

Public Property Get StockPendienteT46() As Double

    StockPendienteT46 = mudtprops.StockPendienteT46

End Property

Public Property Get StockPendienteT48() As Double

    StockPendienteT48 = mudtprops.StockPendienteT48

End Property

Public Property Get StockPendienteT50() As Double
    
    StockPendienteT50 = mudtprops.StockPendienteT50

End Property

Public Property Get StockPendienteT52() As Double

    StockPendienteT52 = mudtprops.StockPendienteT52

End Property

Public Property Get StockPendienteT54() As Double
    
    StockPendienteT54 = mudtprops.StockPendienteT54
    
End Property

Public Property Get StockPendienteT56() As Double

    StockPendienteT56 = mudtprops.StockPendienteT56

End Property

Public Property Let TemporadaID(Value As Long)

    If Not mflgNew Then Err.Raise 383
  
    mobjValid.RuleBroken "TemporadaID", (Value = 0)

    mudtprops.TemporadaID = Value
  
    Set mobjArticulos = Nothing
'    Set mobjArticulos = New TextList
'    mobjArticulos.Load "vNombreArticulos", , mudtprops.TemporadaID
'    mudtprops.ArticuloID = mobjArticulos.Key(mobjArticulos.IndexItem(1))
    mudtprops.ArticuloID = 0
  
    mflgDirty = True

End Property

Public Property Get TemporadaID() As Long

    TemporadaID = mudtprops.TemporadaID
  
End Property

Public Property Get objArticulo() As Articulo
    Dim auxArticulo As Articulo
  
    Set auxArticulo = New Articulo
    auxArticulo.Load mudtprops.ArticuloID
    Set objArticulo = auxArticulo
    Set auxArticulo = Nothing

End Property

'Esta propiedad devuelve la referencia de proveedor para el artículo /color.
Public Property Get SuReferencia() As String
    Dim strSuReferenciaArticulo As String
    
    strSuReferenciaArticulo = objArticulo.SuReferencia
    
    SuReferencia = strSuReferenciaArticulo & "." & Codigo

End Property

Private Function GetState() As String

    Dim udtData As ArticuloColorData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As ArticuloColorData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub


' Registra un movimiento de inventario
' No tiene que haber movimientos realizados.
Public Sub StockInventario(CantidadT36 As Double, CantidadT38 As Double, _
                           CantidadT40 As Double, CantidadT42 As Double, _
                           CantidadT44 As Double, CantidadT46 As Double, _
                           CantidadT48 As Double, CantidadT50 As Double, _
                           CantidadT52 As Double, CantidadT54 As Double, _
                           CantidadT56 As Double, Concepto As String, Precio As Double)
    
    If mflgHayMovimiento Then Err.Raise 383
    If mflgNew Then Err.Raise 383
    
    Set mobjMoviArticulo = New MoviArticulo
    
    With mobjMoviArticulo
        .BeginEdit
        .CantidadT36 = CantidadT36
        .CantidadT38 = CantidadT38
        .CantidadT40 = CantidadT40
        .CantidadT42 = CantidadT42
        .CantidadT44 = CantidadT44
        .CantidadT46 = CantidadT46
        .CantidadT48 = CantidadT48
        .CantidadT50 = CantidadT50
        .CantidadT52 = CantidadT52
        .CantidadT54 = CantidadT54
        .CantidadT56 = CantidadT56
        .Tipo = TMMInventarioTexto
        .Concepto = Concepto
        mudtprops.StockActualT36 = CantidadT36
        mudtprops.StockActualT38 = CantidadT38
        mudtprops.StockActualT40 = CantidadT40
        mudtprops.StockActualT42 = CantidadT42
        mudtprops.StockActualT44 = CantidadT44
        mudtprops.StockActualT46 = CantidadT46
        mudtprops.StockActualT48 = CantidadT48
        mudtprops.StockActualT50 = CantidadT50
        mudtprops.StockActualT52 = CantidadT52
        mudtprops.StockActualT54 = CantidadT54
        mudtprops.StockActualT56 = CantidadT56
        .StockFinal = StockActual
        .Precio = Precio
    End With
    
    mflgDirty = True

    mflgHayMovimiento = True
    
End Sub

' Registra un movimiento de Entrada
' No tiene que haber movimientos realizados.
Public Sub StockEntrada(CantidadT36 As Double, CantidadT38 As Double, _
                        CantidadT40 As Double, CantidadT42 As Double, _
                        CantidadT44 As Double, CantidadT46 As Double, _
                        CantidadT48 As Double, CantidadT50 As Double, _
                        CantidadT52 As Double, CantidadT54 As Double, _
                        CantidadT56 As Double, Concepto As String, Precio As Double)
    
    If mflgHayMovimiento Then Err.Raise 383
    If mflgNew Then Err.Raise 383
    
    Set mobjMoviArticulo = New MoviArticulo
    
    With mobjMoviArticulo
        .BeginEdit
        .CantidadT36 = CantidadT36
        .CantidadT38 = CantidadT38
        .CantidadT40 = CantidadT40
        .CantidadT42 = CantidadT42
        .CantidadT44 = CantidadT44
        .CantidadT46 = CantidadT46
        .CantidadT48 = CantidadT48
        .CantidadT50 = CantidadT50
        .CantidadT52 = CantidadT52
        .CantidadT54 = CantidadT54
        .CantidadT56 = CantidadT56
        .Tipo = TMMEntradaTexto
        .Concepto = Concepto
        mudtprops.StockActualT36 = mudtprops.StockActualT36 + CantidadT36
        mudtprops.StockActualT38 = mudtprops.StockActualT38 + CantidadT38
        mudtprops.StockActualT40 = mudtprops.StockActualT40 + CantidadT40
        mudtprops.StockActualT42 = mudtprops.StockActualT42 + CantidadT42
        mudtprops.StockActualT44 = mudtprops.StockActualT44 + CantidadT44
        mudtprops.StockActualT46 = mudtprops.StockActualT46 + CantidadT46
        mudtprops.StockActualT48 = mudtprops.StockActualT48 + CantidadT48
        mudtprops.StockActualT50 = mudtprops.StockActualT50 + CantidadT50
        mudtprops.StockActualT52 = mudtprops.StockActualT52 + CantidadT52
        mudtprops.StockActualT54 = mudtprops.StockActualT54 + CantidadT54
        mudtprops.StockActualT56 = mudtprops.StockActualT56 + CantidadT56
        
        .StockFinal = StockActual
        .Precio = Precio
    End With
    
    mflgDirty = True

    mflgHayMovimiento = True
    
End Sub

' Registra un movimiento de Salida
' No tiene que haber movimientos realizados.
Public Sub StockSalida(CantidadT36 As Double, CantidadT38 As Double, _
                       CantidadT40 As Double, CantidadT42 As Double, _
                       CantidadT44 As Double, CantidadT46 As Double, _
                       CantidadT48 As Double, CantidadT50 As Double, _
                       CantidadT52 As Double, CantidadT54 As Double, _
                       CantidadT56 As Double, Concepto As String, Precio As Double)
    
    If mflgHayMovimiento Then Err.Raise 383
    If mflgNew Then Err.Raise 383
    
    Set mobjMoviArticulo = New MoviArticulo
    
    With mobjMoviArticulo
        .BeginEdit
        .CantidadT36 = CantidadT36
        .CantidadT38 = CantidadT38
        .CantidadT40 = CantidadT40
        .CantidadT42 = CantidadT42
        .CantidadT44 = CantidadT44
        .CantidadT46 = CantidadT46
        .CantidadT48 = CantidadT48
        .CantidadT50 = CantidadT50
        .CantidadT52 = CantidadT52
        .CantidadT54 = CantidadT54
        .CantidadT56 = CantidadT56
        .Tipo = TMMSalidaTexto
        .Concepto = Concepto
        mudtprops.StockActualT36 = mudtprops.StockActualT36 - CantidadT36
        mudtprops.StockActualT38 = mudtprops.StockActualT38 - CantidadT38
        mudtprops.StockActualT40 = mudtprops.StockActualT40 - CantidadT40
        mudtprops.StockActualT42 = mudtprops.StockActualT42 - CantidadT42
        mudtprops.StockActualT44 = mudtprops.StockActualT44 - CantidadT44
        mudtprops.StockActualT46 = mudtprops.StockActualT46 - CantidadT46
        mudtprops.StockActualT48 = mudtprops.StockActualT48 - CantidadT48
        mudtprops.StockActualT50 = mudtprops.StockActualT50 - CantidadT50
        mudtprops.StockActualT52 = mudtprops.StockActualT52 - CantidadT52
        mudtprops.StockActualT54 = mudtprops.StockActualT54 - CantidadT54
        mudtprops.StockActualT56 = mudtprops.StockActualT56 - CantidadT56
        .StockFinal = StockActual
        .Precio = Precio
    End With
    
    mflgDirty = True

    mflgHayMovimiento = True
    
End Sub

' Registra un movimiento de Entrega de Articulo (recepcion)
' No tiene que haber movimientos realizados.
Public Sub StockEntrega(CantidadT36 As Double, CantidadT38 As Double, _
                        CantidadT40 As Double, CantidadT42 As Double, _
                        CantidadT44 As Double, CantidadT46 As Double, _
                        CantidadT48 As Double, CantidadT50 As Double, _
                        CantidadT52 As Double, CantidadT54 As Double, _
                        CantidadT56 As Double, Concepto As String, Precio As Double)
    
    If mflgHayMovimiento Then Err.Raise 383
    If mflgNew Then Err.Raise 383
    
    Set mobjMoviArticulo = New MoviArticulo
    
    With mobjMoviArticulo
        .BeginEdit
        .CantidadT36 = CantidadT36
        .CantidadT38 = CantidadT38
        .CantidadT40 = CantidadT40
        .CantidadT42 = CantidadT42
        .CantidadT44 = CantidadT44
        .CantidadT46 = CantidadT46
        .CantidadT48 = CantidadT48
        .CantidadT50 = CantidadT50
        .CantidadT52 = CantidadT52
        .CantidadT54 = CantidadT54
        .CantidadT56 = CantidadT56
        .Tipo = TMMEntregaTexto
        .Concepto = Concepto
        mudtprops.StockPendienteT36 = mudtprops.StockPendienteT36 - CantidadT36
        mudtprops.StockPendienteT38 = mudtprops.StockPendienteT38 - CantidadT38
        mudtprops.StockPendienteT40 = mudtprops.StockPendienteT40 - CantidadT40
        mudtprops.StockPendienteT42 = mudtprops.StockPendienteT42 - CantidadT42
        mudtprops.StockPendienteT44 = mudtprops.StockPendienteT44 - CantidadT44
        mudtprops.StockPendienteT46 = mudtprops.StockPendienteT46 - CantidadT46
        mudtprops.StockPendienteT48 = mudtprops.StockPendienteT48 - CantidadT48
        mudtprops.StockPendienteT50 = mudtprops.StockPendienteT50 - CantidadT50
        mudtprops.StockPendienteT52 = mudtprops.StockPendienteT52 - CantidadT52
        mudtprops.StockPendienteT54 = mudtprops.StockPendienteT54 - CantidadT54
        mudtprops.StockPendienteT56 = mudtprops.StockPendienteT56 - CantidadT56
        .StockFinal = StockPendiente
        .Precio = Precio
    End With
    
    mflgDirty = True

    mflgHayMovimiento = True
    
End Sub

' Registra un movimiento de Reserva de Articulo (entrega pendiente)
' No tiene que haber movimientos realizados.
Public Sub StockReserva(CantidadT36 As Double, CantidadT38 As Double, _
                        CantidadT40 As Double, CantidadT42 As Double, _
                        CantidadT44 As Double, CantidadT46 As Double, _
                        CantidadT48 As Double, CantidadT50 As Double, _
                        CantidadT52 As Double, CantidadT54 As Double, _
                        CantidadT56 As Double, Concepto As String, Precio As Double)
    
    If mflgHayMovimiento Then Err.Raise 383
    If mflgNew Then Err.Raise 383
    
    Set mobjMoviArticulo = New MoviArticulo
    
    With mobjMoviArticulo
        .BeginEdit
        .CantidadT36 = CantidadT36
        .CantidadT38 = CantidadT38
        .CantidadT40 = CantidadT40
        .CantidadT42 = CantidadT42
        .CantidadT44 = CantidadT44
        .CantidadT46 = CantidadT46
        .CantidadT48 = CantidadT48
        .CantidadT50 = CantidadT50
        .CantidadT52 = CantidadT52
        .CantidadT54 = CantidadT54
        .CantidadT56 = CantidadT56
        .Tipo = TMMReservaTexto
        .Concepto = Concepto
        mudtprops.StockPendienteT36 = mudtprops.StockPendienteT36 + CantidadT36
        mudtprops.StockPendienteT38 = mudtprops.StockPendienteT38 + CantidadT38
        mudtprops.StockPendienteT40 = mudtprops.StockPendienteT40 + CantidadT40
        mudtprops.StockPendienteT42 = mudtprops.StockPendienteT42 + CantidadT42
        mudtprops.StockPendienteT44 = mudtprops.StockPendienteT44 + CantidadT44
        mudtprops.StockPendienteT46 = mudtprops.StockPendienteT46 + CantidadT46
        mudtprops.StockPendienteT48 = mudtprops.StockPendienteT48 + CantidadT48
        mudtprops.StockPendienteT50 = mudtprops.StockPendienteT50 + CantidadT50
        mudtprops.StockPendienteT52 = mudtprops.StockPendienteT52 + CantidadT52
        mudtprops.StockPendienteT54 = mudtprops.StockPendienteT54 + CantidadT54
        mudtprops.StockPendienteT56 = mudtprops.StockPendienteT56 + CantidadT56
        .StockFinal = StockPendiente
        .Precio = Precio
    End With
    
    mflgDirty = True
    mflgHayMovimiento = True
    
End Sub

Private Sub MovimientoArticulo()
  
    Dim auxArticulo As Articulo
  
    Set auxArticulo = New Articulo
    auxArticulo.Load mudtprops.ArticuloID

    With mobjMoviArticulo
        .PrecioVenta = auxArticulo.PrecioVenta
        .PrecioCoste = auxArticulo.PrecioCoste
        .ArticuloColorID = mudtprops.ArticuloColorID
        .ApplyEdit
    End With
  
    With auxArticulo
        .BeginEdit
        Select Case mobjMoviArticulo.Tipo
            Case TMMInventarioTexto
                .StockActual = mobjMoviArticulo.Cantidad
            Case TMMEntradaTexto
                .StockActual = .StockActual + mobjMoviArticulo.Cantidad
            Case TMMSalidaTexto
                .StockActual = .StockActual - mobjMoviArticulo.Cantidad
            Case TMMReservaTexto
                .StockPendiente = .StockPendiente + mobjMoviArticulo.Cantidad
            Case TMMEntregaTexto
                .StockPendiente = .StockPendiente - mobjMoviArticulo.Cantidad
        End Select
        .ApplyEdit
    End With
     
    Set mobjMoviArticulo = Nothing
    mflgHayMovimiento = False
    Set auxArticulo = Nothing

End Sub

Private Sub CargarListaArticulos()
    
    Set mobjArticulos = New TextList
    mobjArticulos.Load "vNombreArticulos", , mudtprops.TemporadaID
    
End Sub
  

