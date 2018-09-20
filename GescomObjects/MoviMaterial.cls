VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MoviMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As MoviMaterialProps
Private mudtSave As MoviMaterialProps

' Aqui declaro el ComboBox de tipos de movimiento
Private mobjTipoMoviMaterial As TextList

' Aqui declaro el ComboBox de materiales
'Private mobjMaterial As TextList

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

    Dim objPersist As MoviMaterialPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.MoviMaterialPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.MoviMaterialID
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
    mobjValid.RuleBroken "MaterialID", True
    mobjValid.RuleBroken "Cantidad", True
    mobjValid.RuleBroken "Tipo", True
      
    Set mobjTipoMoviMaterial = New TextList
    mobjTipoMoviMaterial.Load "MoviMateriales"
    mudtprops.Tipo = mobjTipoMoviMaterial.Key(mobjTipoMoviMaterial.IndexItem(1))
  
    mudtprops.StockFinal = 0
    mudtprops.Cantidad = 0
    mudtprops.PrecioCostePTA = 0
    mudtprops.PrecioCosteEUR = 0
    mudtprops.PrecioPTA = 0
    mudtprops.PrecioEUR = 0
    mudtprops.DocumentoID = 0
    mudtprops.TipoDocumento = vbNullString
      
End Sub

Private Sub Class_Terminate()

  Set mobjValid = Nothing
  
  Set mobjTipoMoviMaterial = Nothing
  
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

Public Sub Load(MoviMaterialID As Long)

    Dim objPersist As MoviMaterialPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.MoviMaterialPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(MoviMaterialID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Concepto", False
    mobjValid.RuleBroken "MaterialID", False
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

Public Property Get MoviMaterialID() As Long

    MoviMaterialID = mudtprops.MoviMaterialID
  
End Property

Public Property Get Fecha() As Date

    Fecha = mudtprops.Fecha

End Property

Public Property Let MaterialID(Value As Long)

    If Not mflgEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "MaterialID", (Value <= 0)

    mudtprops.MaterialID = Value
    mflgDirty = True

End Property

Public Property Get MaterialID() As Long

    MaterialID = mudtprops.MaterialID
  
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

    Set Tipos = mobjTipoMoviMaterial

End Property

Public Property Let PrecioCoste(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "PrecioCoste", (Value < 0)
    mudtprops.PrecioCosteEUR = Round(Value, 3)
    mudtprops.PrecioCostePTA = EUR2PTA(Value)
        
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
    mudtprops.PrecioPTA = EUR2PTA(Value)
    
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

Public Property Let Cantidad(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Cantidad no puede ser negativa"
  
    ' Se valida que no sean valores negativos.
    mobjValid.RuleBroken "Cantidad", (Value < 0)
  
    mudtprops.Cantidad = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get Cantidad() As Double

    Cantidad = mudtprops.Cantidad
  
End Property

Public Property Let DocumentoID(Value As Long)
  
    If Not mflgEditing Then Err.Raise 383
    
    mudtprops.DocumentoID = Value
    
    mflgDirty = True

End Property

Public Property Get DocumentoID() As Long

    DocumentoID = mudtprops.DocumentoID
  
End Property

Public Property Let TipoDocumento(Value As String)
  
    If Not mflgEditing Then Err.Raise 383
    
    mudtprops.TipoDocumento = Value
    
    mflgDirty = True

End Property

Public Property Get TipoDocumento() As String

    TipoDocumento = mudtprops.TipoDocumento
  
End Property

Private Function GetState() As String

    Dim udtData As MoviMaterialData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As MoviMaterialData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
