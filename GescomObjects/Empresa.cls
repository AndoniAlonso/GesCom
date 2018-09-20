VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As EmpresaProps
Private mudtSave As EmpresaProps

' Aqui declaro el objeto 'child' Direccion
Private mobjDireccion As Direccion

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
  
  mobjDireccion.BeginEdit
  mflgEditing = True

End Sub

Public Sub CancelEdit()

  If Not mflgEditing Then Err.Raise 445
  
  mflgEditing = False
  mflgDeleted = False
  mflgDirty = False
  ' restore object state
  LSet mudtprops = mudtSave
  
  mobjDireccion.CancelEdit
  
End Sub

Public Sub ApplyEdit()

  Dim objPersist As EmpresaPersist
  
  If Not mflgEditing Then Err.Raise 445

  Set objPersist = CreateObject("GCServerMTS.EmpresaPersist", PERSIST_SERVER)

  If mflgDeleted And Not mflgNew Then
    ' code to delete the object's data goes here
    objPersist.DeleteObject mudtprops.EmpresaID, DireccionID
        
    ' tratar con los datos de direcciones
    mobjDireccion.Delete
    mobjDireccion.ApplyEdit
    
    mflgNew = True
    mflgDeleted = False

  ElseIf IsDirty Or mflgNew Then
    If Not IsValid Then Err.Raise 445
        mudtprops.Direccion.Buffer = mobjDireccion.GetState
    
    ' save object to database if appropriate
    SetState objPersist.Save(GetState)
    
    ' salvar el objeto hijo primero (en este caso)
    mobjDireccion.SetState (mudtprops.Direccion.Buffer)
    mobjDireccion.ApplyEdit
    mudtprops.DireccionID = mobjDireccion.DireccionID
    
    ' save object state
    LSet mudtSave = mudtprops
    mflgNew = False
  Else
    mobjDireccion.ApplyEdit

  End If
  Set objPersist = Nothing
  mflgDirty = False
  mflgEditing = False

End Sub

Private Sub Class_Initialize()

  mflgNew = True
  Set mobjValid = New BrokenRules
 
  Set mobjDireccion = New Direccion
  
  ' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  '
  
  mobjValid.RuleBroken "Nombre", True
  mobjValid.RuleBroken "Codigo", True
  mobjValid.RuleBroken "Titular", True
  mobjValid.RuleBroken "CodigoContawin", True
  mobjValid.RuleBroken "EmpresaContawin", True
  'mobjValid.RuleBroken "DireccionID", True
  
  mudtprops.PedidoCompras = 0
  mudtprops.PedidoVentas = 0
  mudtprops.AlbaranCompras = 0
  mudtprops.AlbaranVentas = 0
  mudtprops.FacturaCompras = 0
  mudtprops.FacturaVentas = 0
  ' Por defecto un tratamiento normal del IVA.
  mudtprops.TratamientoIVA = "S"
  mudtprops.CodigoContawin = 0
  mudtprops.EmpresaContawin = vbNullString
End Sub

Private Sub Class_Terminate()

  Set mobjValid = Nothing
 
  Set mobjDireccion = Nothing

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

Public Sub Load(EmpresaID As Long)

  Dim objPersist As EmpresaPersist
  
  If mflgEditing Then Err.Raise 445
  If Not mflgNew Then Err.Raise 445
  
 ' code to load the object goes here
  Set objPersist = CreateObject("GCServerMTS.EmpresaPersist", PERSIST_SERVER)
  SetState objPersist.Fetch(EmpresaID)
  Set objPersist = Nothing

  mobjValid.RuleBroken "Nombre", False
  mobjValid.RuleBroken "Codigo", False
  mobjValid.RuleBroken "Titular", False
  mobjValid.RuleBroken "CodigoContawin", False
  mobjValid.RuleBroken "EmpresaContawin", False
  
  mobjDireccion.SetState (mudtprops.Direccion.Buffer)
  
  mflgNew = False

End Sub

Public Sub Delete()

  If Not mflgEditing Then Err.Raise 445
  
  mflgDeleted = True
  
  mobjDireccion.Delete

End Sub

Public Property Get IsDeleted() As Boolean

  IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

  IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

  IsDirty = mflgDirty Or mobjDireccion.IsDirty

End Property

Public Property Get EmpresaID() As Long

  EmpresaID = mudtprops.EmpresaID
  
End Property

Public Property Let Codigo(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtprops.Codigo) Then _
    Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
  ' Regla de longitud de Codigo si es distinto de su longitud máxima--> mal
  mobjValid.RuleBroken "Codigo", (Len(Trim(Value)) <> Len(mudtprops.Codigo))

  mudtprops.Codigo = UCase$(Trim(Value))
  mflgDirty = True

End Property

Public Property Get Codigo() As String

  Codigo = Trim(mudtprops.Codigo)

End Property

Public Property Let Nombre(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtprops.Nombre) Then _
    Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
  ' Regla de longitud de Nombre si es distinto de cero-->bien
  mobjValid.RuleBroken "Nombre", (Len(Trim(Value)) = 0)

  mudtprops.Nombre = Trim(Value)
  mflgDirty = True

End Property

Public Property Get Nombre() As String

  Nombre = Trim(mudtprops.Nombre)

End Property
Public Property Let Titular(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtprops.Titular) Then _
    Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
  ' Regla de longitud de Titular si es distinto de 0 --> bien
  mobjValid.RuleBroken "Titular", (Len(Trim(Value)) = 0)

  mudtprops.Titular = Trim(Value)
  mflgDirty = True

End Property

Public Property Get Titular() As String

  Titular = Trim(mudtprops.Titular)

End Property
Public Property Let DNINIF(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtprops.DNINIF) Then _
    Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"

  mudtprops.DNINIF = UCase$(Trim(Value))
  mflgDirty = True

End Property

Public Property Get DNINIF() As String

  DNINIF = Trim(mudtprops.DNINIF)

End Property
Public Property Let Actividad(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtprops.Actividad) Then _
    Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"

  mudtprops.Actividad = Trim(Value)
  mflgDirty = True

End Property

Public Property Get Actividad() As String

  Actividad = Trim(mudtprops.Actividad)

End Property

Public Property Let DireccionID(Value As Long)
  
  If Not mflgEditing Then Err.Raise 383
  
  'mobjValid.RuleBroken "DireccionID", (Value=0)
  mudtprops.DireccionID = Value
  mflgDirty = True

End Property

Public Property Get DireccionID() As Long

  DireccionID = mudtprops.DireccionID
  
End Property

Public Property Get Direccion() As Direccion

  Set Direccion = mobjDireccion
  
End Property

Public Property Let PedidoCompras(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Nº de pedido no puede ser negativo"

  mobjValid.RuleBroken "PedidoCompras", (Value < 0)
  mudtprops.PedidoCompras = Value
  mflgDirty = True

End Property
Public Property Get PedidoCompras() As Long

  PedidoCompras = mudtprops.PedidoCompras
  
End Property
Public Property Let PedidoVentas(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Nº de pedido no puede ser negativo"

  mobjValid.RuleBroken "PedidoVentas", (Value < 0)
  mudtprops.PedidoVentas = Value
  mflgDirty = True

End Property
Public Property Get PedidoVentas() As Long

  PedidoVentas = mudtprops.PedidoVentas
  
End Property
Public Property Let AlbaranCompras(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Nº de Albaran no puede ser negativo"

  mobjValid.RuleBroken "AlbaranCompras", (Value < 0)
  mudtprops.AlbaranCompras = Value
  mflgDirty = True

End Property

Public Property Get AlbaranCompras() As Long

  AlbaranCompras = mudtprops.AlbaranCompras
  
End Property

Public Property Let AlbaranVentas(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Nº de Albaran no puede ser negativo"

  mobjValid.RuleBroken "AlbaranVentas", (Value < 0)
  mudtprops.AlbaranVentas = Value
  mflgDirty = True

End Property

Public Property Get AlbaranVentas() As Long

  AlbaranVentas = mudtprops.AlbaranVentas
  
End Property

Public Property Let FacturaCompras(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Nº de Factura no puede ser negativo"

  mobjValid.RuleBroken "FacturaCompras", (Value < 0)
  mudtprops.FacturaCompras = Value
  mflgDirty = True

End Property

Public Property Get FacturaCompras() As Long

  FacturaCompras = mudtprops.FacturaCompras
  
End Property

Public Property Let FacturaVentas(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Nº de Factura no puede ser negativo"

  mobjValid.RuleBroken "FacturaVentas", (Value < 0)
  mudtprops.FacturaVentas = Value
  mflgDirty = True

End Property

Public Property Get FacturaVentas() As Long

  FacturaVentas = mudtprops.FacturaVentas
  
End Property

Public Property Let OrdenCorte(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Nº de Orden de corte no puede ser negativo"

  mobjValid.RuleBroken "OrdenCorte", (Value < 0)
  mudtprops.OrdenCorte = Value
  mflgDirty = True

End Property

Public Property Get OrdenCorte() As Long

  OrdenCorte = mudtprops.OrdenCorte
  
End Property

Public Property Let AnularIVA(Value As Boolean)

  If Not mflgEditing Then Err.Raise 383
  
  mudtprops.TratamientoIVA = IIf(Value, "N", "S")
  mflgDirty = True

End Property

Public Property Get AnularIVA() As Boolean

  AnularIVA = (mudtprops.TratamientoIVA = "N")
  
End Property

Private Function GetState() As String

  Dim udtdata As EmpresaData

  With mudtprops
    .IsDeleted = mflgDeleted
    .IsNew = mflgNew
    .IsDirty = mflgDirty
  End With

  LSet udtdata = mudtprops
  GetState = udtdata.Buffer

End Function

Private Sub SetState(Buffer As String)

  Dim udtdata As EmpresaData
  
  udtdata.Buffer = Buffer
  LSet mudtprops = udtdata

End Sub

Public Function IncrementaPedidoVentas() As Long
  Me.BeginEdit
  Me.PedidoVentas = Me.PedidoVentas + 1
  Me.ApplyEdit
  IncrementaPedidoVentas = Me.PedidoVentas
End Function

Public Function DecrementaPedidoVentas() As Long
  Me.BeginEdit
  Me.PedidoVentas = Me.PedidoVentas - 1
  Me.ApplyEdit
  DecrementaPedidoVentas = Me.PedidoVentas
End Function

Public Sub EstablecePedidoVentas(PedidoVentas As Long)
  Me.BeginEdit
  Me.PedidoVentas = PedidoVentas
  Me.ApplyEdit
End Sub

Public Function IncrementaPedidoCompras() As Long
  
    Me.BeginEdit
    Me.PedidoCompras = Me.PedidoCompras + 1
    Me.ApplyEdit
    IncrementaPedidoCompras = Me.PedidoCompras
    
End Function

Public Function DecrementaPedidoCompras() As Long
  
    Me.BeginEdit
    Me.PedidoCompras = Me.PedidoCompras - 1
    Me.ApplyEdit
    DecrementaPedidoCompras = Me.PedidoCompras
    
End Function

Public Sub EstablecePedidoCompras(PedidoCompras As Long)
    
    Me.BeginEdit
    Me.PedidoCompras = PedidoCompras
    Me.ApplyEdit
    
End Sub

Public Function IncrementaAlbaranVentas() As Long
  Me.BeginEdit
  Me.AlbaranVentas = Me.AlbaranVentas + 1
  Me.ApplyEdit
  IncrementaAlbaranVentas = Me.AlbaranVentas
End Function

Public Function DecrementaAlbaranVentas() As Long
  Me.BeginEdit
  Me.AlbaranVentas = Me.AlbaranVentas - 1
  Me.ApplyEdit
  DecrementaAlbaranVentas = Me.AlbaranVentas
End Function

Public Sub EstableceAlbaranVentas(AlbaranVentas As Long)
  Me.BeginEdit
  Me.AlbaranVentas = AlbaranVentas
  Me.ApplyEdit
End Sub

Public Function IncrementaFacturaVentas() As Long
  Me.BeginEdit
  Me.FacturaVentas = Me.FacturaVentas + 1
  Me.ApplyEdit
  IncrementaFacturaVentas = Me.FacturaVentas
End Function

Public Function DecrementaFacturaVentas() As Long
  Me.BeginEdit
  Me.FacturaVentas = Me.FacturaVentas - 1
  Me.ApplyEdit
  DecrementaFacturaVentas = Me.FacturaVentas
End Function

Public Sub EstableceFacturaVentas(FacturaVentas As Long)
  Me.BeginEdit
  Me.FacturaVentas = FacturaVentas
  Me.ApplyEdit
End Sub

Public Function IncrementaAlbaranCompras() As Long
  
    Me.BeginEdit
    Me.AlbaranCompras = Me.AlbaranCompras + 1
    Me.ApplyEdit
    IncrementaAlbaranCompras = Me.AlbaranCompras
    
End Function

Public Function DecrementaAlbaranCompras() As Long
  
    Me.BeginEdit
    Me.AlbaranCompras = Me.AlbaranCompras - 1
    Me.ApplyEdit
    DecrementaAlbaranCompras = Me.AlbaranCompras
    
End Function

Public Sub EstableceAlbaranCompras(AlbaranCompras As Long)
  
    Me.BeginEdit
    Me.AlbaranCompras = AlbaranCompras
    Me.ApplyEdit

End Sub

Public Function IncrementaFacturaCompras() As Long
  
    Me.BeginEdit
    Me.FacturaCompras = Me.FacturaCompras + 1
    Me.ApplyEdit
    IncrementaFacturaCompras = Me.FacturaCompras
    
End Function

Public Function DecrementaFacturaCompras() As Long
  
    Me.BeginEdit
    Me.FacturaCompras = Me.FacturaCompras - 1
    Me.ApplyEdit
    DecrementaFacturaCompras = Me.FacturaCompras
    
End Function

Public Sub EstableceFacturaCompras(FacturaCompras As Long)
  
    Me.BeginEdit
    Me.FacturaCompras = FacturaCompras
    Me.ApplyEdit
    
End Sub

Public Function IncrementaOrdenCorte() As Long
  
    Me.BeginEdit
    Me.OrdenCorte = Me.OrdenCorte + 1
    Me.ApplyEdit
    IncrementaOrdenCorte = Me.OrdenCorte
    
End Function

Public Function DecrementaOrdenCorte() As Long
  
    Me.BeginEdit
    Me.OrdenCorte = Me.OrdenCorte - 1
    Me.ApplyEdit
    DecrementaOrdenCorte = Me.OrdenCorte
    
End Function

Public Sub EstableceOrdenCorte(OrdenCorte As Long)
  
    Me.BeginEdit
    Me.OrdenCorte = OrdenCorte
    Me.ApplyEdit

End Sub

Public Property Let EmpresaContawin(Value As String)
  
    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.EmpresaContawin) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mobjValid.RuleBroken "EmpresaContawin", False

    mudtprops.EmpresaContawin = Trim(Value)
    mflgDirty = True

End Property

Public Property Get EmpresaContawin() As String

    EmpresaContawin = Trim(mudtprops.EmpresaContawin)

End Property

Public Property Let CodigoContawin(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value <= 0 Then _
    Err.Raise vbObjectError + 1001, "Código empresa contawin no puede ser negativo ni cero"

  mobjValid.RuleBroken "CodigoContawin", (Value <= 0)
  mudtprops.CodigoContawin = Value
  mflgDirty = True

End Property

Public Property Get CodigoContawin() As Long

  CodigoContawin = mudtprops.CodigoContawin
  
End Property




