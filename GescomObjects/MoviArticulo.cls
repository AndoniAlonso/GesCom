VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MoviArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As MoviArticuloProps
Private mudtSave As MoviArticuloProps

' Aqui declaro el ComboBox de tipos de movimiento
Private mobjTipoMoviArticulo As TextList

' Aqui declaro el ComboBox de Articulos
'Private mobjArticulo As TextList

Event Valid(IsValid As Boolean)

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean

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
  
End Sub

Public Sub ApplyEdit()

    Dim objPersist As MoviArticuloPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.MoviArticuloPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.MoviArticuloID
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

End Sub

Private Sub Class_Initialize()
    
    mflgNew = True
    Set mobjValid = New BrokenRules
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    '
    mobjValid.RuleBroken "Concepto", True
    mobjValid.RuleBroken "ArticuloColorID", True
    mobjValid.RuleBroken "Cantidad", True
    mobjValid.RuleBroken "Tipo", True
      
    Set mobjTipoMoviArticulo = New TextList
    mobjTipoMoviArticulo.Load "MoviMateriales"
    mudtprops.Tipo = mobjTipoMoviArticulo.Key(mobjTipoMoviArticulo.IndexItem(1))
  
    mudtprops.StockFinal = 0
    mudtprops.CantidadT36 = 0
    mudtprops.CantidadT38 = 0
    mudtprops.CantidadT40 = 0
    mudtprops.CantidadT42 = 0
    mudtprops.CantidadT44 = 0
    mudtprops.CantidadT46 = 0
    mudtprops.CantidadT48 = 0
    mudtprops.CantidadT50 = 0
    mudtprops.CantidadT52 = 0
    mudtprops.CantidadT54 = 0
    mudtprops.CantidadT56 = 0
    mudtprops.PrecioCosteEUR = 0
    mudtprops.PrecioVentaEUR = 0
    mudtprops.PrecioEUR = 0
      
End Sub

Private Sub Class_Terminate()
    
    Set mobjValid = Nothing
    Set mobjTipoMoviArticulo = Nothing

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

Public Sub Load(MoviArticuloID As Long)

    Dim objPersist As MoviArticuloPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.MoviArticuloPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(MoviArticuloID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Concepto", False
    mobjValid.RuleBroken "ArticuloColorID", False
    mobjValid.RuleBroken "Cantidad", False
    mobjValid.RuleBroken "StockFinal", False
    mobjValid.RuleBroken "Tipo", False
      
    mflgNew = False

End Sub

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

Public Property Get MoviArticuloID() As Long

    MoviArticuloID = mudtprops.MoviArticuloID
  
End Property

Public Property Get Fecha() As Date

    Fecha = mudtprops.Fecha

End Property

Public Property Let ArticuloColorID(Value As Long)

    If Not mflgEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "ArticuloColorID", (Value <= 0)

    mudtprops.ArticuloColorID = Value
    mflgDirty = True

End Property

Public Property Get ArticuloColorID() As Long

    ArticuloColorID = mudtprops.ArticuloColorID
  
End Property

Public Property Let Concepto(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Concepto) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    mobjValid.RuleBroken "Concepto", (Len(Trim(Value)) = 0)

    mudtprops.Concepto = Value
    mflgDirty = True

End Property

Public Property Get Concepto() As String

    Concepto = Trim(mudtprops.Concepto)

End Property

Public Property Let Tipo(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.Tipo = vbNullString
  
    Select Case Value
        Case TMMInventarioTexto
           mudtprops.Tipo = TMMInventario
        Case TMMEntradaTexto
            mudtprops.Tipo = TMMEntrada
        Case TMMSalidaTexto
            mudtprops.Tipo = TMMSalida
        Case TMMReservaTexto
            mudtprops.Tipo = TMMReserva
        Case TMMEntregaTexto
            mudtprops.Tipo = TMMEntrega
        Case Else
            Err.Raise vbObjectError + 1001, "Valor de tipo de movimiento no permitido"
    End Select
  
    mobjValid.RuleBroken "Tipo", (mudtprops.Tipo = vbNullString)
  
    mflgDirty = True

End Property

Public Property Get Tipo() As String

    Select Case mudtprops.Tipo
        Case TMMInventario
            Tipo = TMMInventarioTexto
        Case TMMEntrada
            Tipo = TMMEntradaTexto
        Case TMMSalida
            Tipo = TMMSalidaTexto
        Case TMMReserva
            Tipo = TMMReservaTexto
        Case TMMEntrega
            Tipo = TMMEntregaTexto
        Case Else
            Tipo = "Valor desconocido"
    End Select
      
End Property

Public Property Get Tipos() As TextList

    Set Tipos = mobjTipoMoviArticulo

End Property

Public Property Let PrecioVenta(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "PrecioVenta", (Value < 0)
    mudtprops.PrecioVentaEUR = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get PrecioVenta() As Double

    PrecioVenta = mudtprops.PrecioVentaEUR
  
End Property

Public Property Let PrecioCoste(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
        
    mobjValid.RuleBroken "PrecioCoste", (Value < 0)
    
    mudtprops.PrecioCosteEUR = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get PrecioCoste() As Double

    PrecioCoste = mudtprops.PrecioCosteEUR
  
End Property

Public Property Let Precio(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "Precio", (Value < 0)
  
    mudtprops.PrecioEUR = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get Precio() As Double

    Precio = mudtprops.PrecioEUR
  
End Property

Public Property Let StockFinal(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
     
    'If Value < 0 Then _
    '   Err.Raise vbObjectError + 1001, "Stock no puede ser negativo"
  
    ' Se valida que no sean valores negativos.
    'mobjValid.RuleBroken "StockFinal", (Value < 0)
  
    mudtprops.StockFinal = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get StockFinal() As Double

    StockFinal = mudtprops.StockFinal
  
End Property

Public Property Let CantidadT36(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Cantidad no puede ser negativa"
  
    mudtprops.CantidadT36 = Round(Value, 2)
    
    ' Se valida que no sean valores negativos y que haya alguno informado.
    mobjValid.RuleBroken "Cantidad", (Value < 0) Or (Cantidad = 0)
  
    mflgDirty = True

End Property

Public Property Get CantidadT36() As Double

    CantidadT36 = mudtprops.CantidadT36
  
End Property

Public Property Let CantidadT38(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Cantidad no puede ser negativa"
  
    mudtprops.CantidadT38 = Round(Value, 2)
    
    ' Se valida que no sean valores negativos y que haya alguno informado.
    mobjValid.RuleBroken "Cantidad", (Value < 0) Or (Cantidad = 0)
  
    mflgDirty = True
    
End Property

Public Property Get CantidadT38() As Double

    CantidadT38 = mudtprops.CantidadT38
  
End Property

Public Property Let CantidadT40(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Cantidad no puede ser negativa"
  
    mudtprops.CantidadT40 = Round(Value, 2)
    
    ' Se valida que no sean valores negativos y que haya alguno informado.
    mobjValid.RuleBroken "Cantidad", (Value < 0) Or (Cantidad = 0)
  
    mflgDirty = True

End Property

Public Property Get CantidadT40() As Double

    CantidadT40 = mudtprops.CantidadT40
  
End Property

Public Property Let CantidadT42(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Cantidad no puede ser negativa"
  
    mudtprops.CantidadT42 = Round(Value, 2)
        
    ' Se valida que no sean valores negativos y que haya alguno informado.
    mobjValid.RuleBroken "Cantidad", (Value < 0) Or (Cantidad = 0)
  
    mflgDirty = True

End Property

Public Property Get CantidadT42() As Double

    CantidadT42 = mudtprops.CantidadT42
  
End Property

Public Property Let CantidadT44(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Cantidad no puede ser negativa"
  
    mudtprops.CantidadT44 = Round(Value, 2)
    
    ' Se valida que no sean valores negativos y que haya alguno informado.
    mobjValid.RuleBroken "Cantidad", (Value < 0) Or (Cantidad = 0)
  
    mflgDirty = True

End Property

Public Property Get CantidadT44() As Double

    CantidadT44 = mudtprops.CantidadT44
  
End Property

Public Property Let CantidadT46(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Cantidad no puede ser negativa"
  
    mudtprops.CantidadT46 = Round(Value, 2)
    
    ' Se valida que no sean valores negativos y que haya alguno informado.
    mobjValid.RuleBroken "Cantidad", (Value < 0) Or (Cantidad = 0)
  
    mflgDirty = True

End Property

Public Property Get CantidadT46() As Double

    CantidadT46 = mudtprops.CantidadT46
  
End Property

Public Property Let CantidadT48(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Cantidad no puede ser negativa"
  
    mudtprops.CantidadT48 = Round(Value, 2)
    
    ' Se valida que no sean valores negativos y que haya alguno informado.
    mobjValid.RuleBroken "Cantidad", (Value < 0) Or (Cantidad = 0)
      
    mflgDirty = True

End Property

Public Property Get CantidadT48() As Double

    CantidadT48 = mudtprops.CantidadT48
  
End Property

Public Property Let CantidadT50(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Cantidad no puede ser negativa"
  
    mudtprops.CantidadT50 = Round(Value, 2)
    
    ' Se valida que no sean valores negativos y que haya alguno informado.
    mobjValid.RuleBroken "Cantidad", (Value < 0) Or (Cantidad = 0)
  
    mflgDirty = True

End Property

Public Property Get CantidadT50() As Double

    CantidadT50 = mudtprops.CantidadT50
  
End Property

Public Property Let CantidadT52(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Cantidad no puede ser negativa"
  
    mudtprops.CantidadT52 = Round(Value, 2)
    
    ' Se valida que no sean valores negativos y que haya alguno informado.
    mobjValid.RuleBroken "Cantidad", (Value < 0) Or (Cantidad = 0)
  
    mflgDirty = True

End Property

Public Property Get CantidadT52() As Double

    CantidadT52 = mudtprops.CantidadT52
  
End Property

Public Property Let CantidadT54(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Cantidad no puede ser negativa"
  
    mudtprops.CantidadT54 = Round(Value, 2)
    
    ' Se valida que no sean valores negativos y que haya alguno informado.
    mobjValid.RuleBroken "Cantidad", (Value < 0) Or (Cantidad = 0)
  
    mflgDirty = True

End Property

Public Property Get CantidadT54() As Double

    CantidadT54 = mudtprops.CantidadT54
  
End Property

Public Property Let CantidadT56(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Cantidad no puede ser negativa"
  
    mudtprops.CantidadT56 = Round(Value, 2)
    
    ' Se valida que no sean valores negativos y que haya alguno informado.
    mobjValid.RuleBroken "Cantidad", (Value < 0) Or (Cantidad = 0)
  
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

Friend Function GetState() As String

    Dim udtData As MoviArticuloData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Friend Sub SetState(Buffer As String)

    Dim udtData As MoviArticuloData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
