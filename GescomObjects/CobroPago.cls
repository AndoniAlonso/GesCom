VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CobroPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private mMoneda As String

' Aqui declaro los ComboBox de la lista de personas (clientes o proveedores),
' formas de pago, y medios de pago.
Private mobjClientes As TextList
Private mobjProveedores As TextList
Private mobjFormasPago As TextList
Private mobjBancos As TextList
Private mobjMediosPago As TextList

Private mudtprops As CobroPagoProps
Private mudtSave As CobroPagoProps
Private mudtChild As CobroPagoProps
Private mobjParent As CobrosPagos
Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChild As Boolean
Private mflgChildEditing As Boolean
Private mflgDeletedRemesa As Boolean
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
  mflgDeletedRemesa = False
  ' restore object state
  LSet mudtprops = mudtSave
  
End Sub

Public Sub ApplyEdit()

  Dim objPersist As CobroPagoPersist
  
  If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
  'If Not mflgChild And mflgNew Then Err.Raise 445
  If Not mflgEditing Then Err.Raise 445

  If Not mflgChild Then
    
    Set objPersist = CreateObject("GCServerMTS.CobroPagoPersist", PERSIST_SERVER)
  
    If mflgDeleted And Not mflgNew Then
      ' code to delete the object's data goes here
      objPersist.DeleteObject mudtprops.CobroPagoID
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
    
    mflgDirty = False
    Set objPersist = Nothing
  
  ElseIf mflgNew Then
    If Not mobjParent Is Nothing Then
      mobjParent.AddCobroPago Me
      Set mobjParent = Nothing
    End If
  
  End If
    
  mflgEditing = False
  
End Sub

Private Sub Class_Initialize()

  mflgNew = True
  Set mobjValid = New BrokenRules
  
  ' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  mobjValid.RuleBroken "Tipo", True
  mobjValid.RuleBroken "NumeroGiro", True
  mobjValid.RuleBroken "Importe", True
  mobjValid.RuleBroken "PersonaID", True
  mobjValid.RuleBroken "FormaPagoID", True
  mobjValid.RuleBroken "EmpresaID", True
  mobjValid.RuleBroken "TemporadaID", True
  
  With mudtprops
    .Tipo = Trim(vbNullString)
    .FechaEmision = Date
    .Vencimiento = Date
    .FacturaID = 0
    .NumeroGiro = 0
    .SituacionContable = "A"    ' Marco la situacion del cobro como "A"--> alta,
                        ' pendiente de codificar las posibles situaciones.
    .SituacionComercial = "A"    ' Marco la situacion del cobro como "A"--> alta,
                        ' pendiente de codificar las posibles situaciones.
    .ImportePTA = 0
    .ImporteEUR = 0
    '.FechaDomiciliacion = vbNullString
    '.FechaContable = vbNullString
    
    Set mobjClientes = New TextList
    mobjClientes.Load "Clientes"
    '.ClienteID = mobjClientes.Key(mobjClientes.IndexItem(1))
        
    Set mobjProveedores = New TextList
    mobjProveedores.Load "Proveedores"
    '.ProveedorID = mobjProveedores.Key(mobjProveedores.IndexItem(1))
    
    .PersonaID = 0
        
    Set mobjFormasPago = New TextList
    mobjFormasPago.Load "FormasDePago"
    .FormaPagoID = mobjFormasPago.Key(mobjFormasPago.IndexItem(1))
        
    Set mobjMediosPago = New TextList
    mobjMediosPago.Load "MediosPago"
    .MedioPagoID = mobjMediosPago.Key(mobjMediosPago.IndexItem(1))
        
    ' Hay que aclarar si se puede domiciliar en el banco de otra empresa
    Set mobjBancos = New TextList
    mobjBancos.Load "vBancosEmpresas"
    '.BancoID = mobjBancos.Key(mobjBancos.IndexItem(1))
    
    mudtprops.TemporadaID = 0
    mudtprops.EmpresaID = 0
        
  End With
  
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjClientes = Nothing
    Set mobjProveedores = Nothing
    Set mobjFormasPago = Nothing
    Set mobjMediosPago = Nothing
    Set mobjBancos = Nothing
  
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

Public Sub Load(CobroPagoID As Long, Moneda As String)

  Dim objPersist As CobroPagoPersist
  
  If mflgChild Then Err.Raise 445
  If mflgEditing Then Err.Raise 445
  If Not mflgNew Then Err.Raise 445
  
  mMoneda = Moneda
  
 ' code to load the object goes here
  Set objPersist = CreateObject("GCServerMTS.CobroPagoPersist", PERSIST_SERVER)
  SetState objPersist.Fetch(CobroPagoID)
  Set objPersist = Nothing
  
  mobjValid.RuleBroken "Tipo", False
  mobjValid.RuleBroken "NumeroGiro", False
  mobjValid.RuleBroken "Importe", False
  mobjValid.RuleBroken "PersonaID", False
  mobjValid.RuleBroken "FormaPagoID", False
  mobjValid.RuleBroken "EmpresaID", False
  mobjValid.RuleBroken "TemporadaID", False

  mflgNew = False

End Sub

Public Sub Delete()

  If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
  If Not mflgEditing Then Err.Raise 445
  
  mflgDeleted = True
  mflgDirty = True
  
End Sub

Public Sub DeleteRemesa()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    If Contabilizado Then _
        Err.Raise vbObjectError + 1001, "No se puede eliminar el cobro de la remesa, ya está contabilizado."
    
    If Remesado Then _
        Err.Raise vbObjectError + 1001, "No se puede eliminar el cobro de la remesa, ya está remesado."
    
    If Not Seleccionado Then _
        Err.Raise vbObjectError + 1001, "No se puede eliminar de la remesa, no esta seleccionado."
        
    mudtprops.BancoID = 0
    mudtprops.FechaDomiciliacion = 0
    mudtprops.SituacionComercial = "A"
    
    mflgDeletedRemesa = True
    mflgDirty = True
    
End Sub

Public Property Get IsDeleted() As Boolean

  IsDeleted = mflgDeleted

End Property

Public Property Get IsDeletedRemesa() As Boolean

  IsDeletedRemesa = mflgDeletedRemesa

End Property

Public Property Get IsNew() As Boolean

  IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

  IsDirty = mflgDirty Or mflgNew

End Property

Friend Sub SetAsChild(objCobrosPagos As CobrosPagos, Moneda As String)

  mMoneda = Moneda
  If mflgNew Then Set mobjParent = objCobrosPagos
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
  mflgDeletedRemesa = False
  ' restore object state
  LSet mudtprops = mudtChild

End Sub

Public Property Get CobroPagoID() As Long

  CobroPagoID = mudtprops.CobroPagoID

End Property

Public Function GetClienteID() As Long

  If mudtprops.Tipo <> "C" Then _
     Err.Raise vbObjectError + 1001, "No es un cobro, no hay cliente."
     
  GetClienteID = mudtprops.PersonaID
  
End Function

' Esta funcion lee el cliente para asignar sus propiedades por defecto
Private Sub LetCliente(Value As String)
Dim objCliente As Cliente

  If Not mflgEditing Then Err.Raise 383
  
  If mudtprops.Tipo <> "C" Then _
     Err.Raise vbObjectError + 1001, "No es un cobro, no hay cliente."
     
  mobjValid.RuleBroken "PersonaID", (mobjClientes.Key(Value) = 0)
    
  'Salimos si queremos asignar algo invalido
  If mobjClientes.Key(Value) = 0 Then Exit Sub
     
  ' Devolvemos error si ya se ha contabilizado el cobro y queremos cambiar de cliente.
  If Contabilizado Then _
     Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, ya se ha contabilizado el cobro"
     
  ' Devolvemos error si ya se ha remesado el cobro y queremos cambiar de cliente.
  If Remesado Then _
     Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, ya se ha remesado el cobro"
     
  ' Devolvemos error si se han asignado a partir de Albaranes (por lo tanto pertenecen a un cliente).
  If HayFactura Then _
     Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, el cobro está relacionado con una factura"
     
  ' Cambiamos de cliente
  Set objCliente = New Cliente
  
  objCliente.Load mobjClientes.Key(Value)
  
  mudtprops.FormaPagoID = objCliente.FormaPagoID
  
  mudtprops.PersonaID = objCliente.ClienteID
  
  Set objCliente = Nothing
  
  mflgDirty = True

End Sub

Private Function GetCliente() As String

  If mudtprops.Tipo <> "C" Then _
     Err.Raise vbObjectError + 1001, "No es un cobro, no hay cliente."
     
  GetCliente = mobjClientes.Item(mudtprops.PersonaID)

End Function

Private Function GetClientes() As TextList

  If mudtprops.Tipo <> "C" Then _
     Err.Raise vbObjectError + 1001, "No es un cobro, no hay cliente."
     
  Set GetClientes = mobjClientes

End Function

'Private Function GetProveedorID() As Long
'
'  If mudtprops.Tipo <> "P" Then _
'     Err.Raise vbObjectError + 1001, "No es un pago, no hay Proveedor."
'
'  GetProveedorID = mudtprops.PersonaID
'
'End Function
' Esta funcion lee el Proveedor para asignar sus propiedades por defecto
Private Sub LetProveedor(Value As String)
Dim objProveedor As Proveedor

  If Not mflgEditing Then Err.Raise 383
  
  If mudtprops.Tipo <> "P" Then _
     Err.Raise vbObjectError + 1001, "No es un pago, no hay Proveedor."
     
  mobjValid.RuleBroken "PersonaID", (mobjProveedores.Key(Value) = 0)
    
  'Salimos si queremos asignar algo invalido
  If mobjProveedores.Key(Value) = 0 Then Exit Sub
     
  ' Devolvemos error si ya se ha contabilizado el cobro y queremos cambiar de Proveedor.
  If Contabilizado Then _
     Err.Raise vbObjectError + 1001, "No se puede cambiar el Proveedor, ya se ha contabilizado el cobro"
     
  ' Devolvemos error si ya se ha remesado el cobro y queremos cambiar de Proveedor.
  If Remesado Then _
     Err.Raise vbObjectError + 1001, "No se puede cambiar el Proveedor, ya se ha remesado el cobro"
     
  ' Devolvemos error si está relacionado con una factura (por lo tanto pertenece a un Proveedor).
  If HayFactura Then _
     Err.Raise vbObjectError + 1001, "No se puede cambiar el Proveedor, el pago está relacionado con una factura"
     
  ' Cambiamos de Proveedor
  Set objProveedor = New Proveedor
  
  objProveedor.Load mobjProveedores.Key(Value)
  
  mudtprops.FormaPagoID = objProveedor.FormaPagoID
  
  mudtprops.PersonaID = objProveedor.ProveedorID
  
  Set objProveedor = Nothing
  
  mflgDirty = True

End Sub

Private Function GetProveedor() As String

  If mudtprops.Tipo <> "P" Then _
     Err.Raise vbObjectError + 1001, "No es un pago, no hay Proveedor."
     
  GetProveedor = mobjProveedores.Item(mudtprops.PersonaID)

End Function

Private Function GetProveedores() As TextList

  If mudtprops.Tipo <> "P" Then _
     Err.Raise vbObjectError + 1001, "No es un pago, no hay Proveedor."
     
  Set GetProveedores = mobjProveedores

End Function
Public Property Get FormaPagoID() As Long

  FormaPagoID = mudtprops.FormaPagoID
  
End Property

Public Property Let FormaPago(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  ' Devolvemos error si está relacionado con una factura (por lo tanto pertenece a un Proveedor).
  If HayFactura Then _
     Err.Raise vbObjectError + 1001, "No se puede cambiar la forma de pago, el pago está relacionado con una factura"
     
  mudtprops.FormaPagoID = mobjFormasPago.Key(Value)
  
  mobjValid.RuleBroken "FormaPagoID", (mudtprops.FormaPagoID = 0)
  
  mflgDirty = True

End Property

Public Property Get FormaPago() As String

  FormaPago = mobjFormasPago.Item(mudtprops.FormaPagoID)

End Property

Public Property Get FormasPago() As TextList

  Set FormasPago = mobjFormasPago

End Property

Public Property Get MedioPagoID() As Long

  MedioPagoID = mudtprops.MedioPagoID
  
End Property

Public Property Let MedioPago(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  mudtprops.MedioPagoID = mobjMediosPago.Key(Value)
  
  mflgDirty = True

End Property

Public Property Get MedioPago() As String

  MedioPago = mobjMediosPago.Item(mudtprops.MedioPagoID)

End Property

Public Property Get MediosPago() As TextList

  Set MediosPago = mobjMediosPago

End Property

Public Property Get BancoID() As Long

  BancoID = mudtprops.BancoID
  
End Property
Public Property Let Banco(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  ' En el caso de cobros solo asignar banco si remesado.
  If Tipo = "C" And Not Remesado Then Err.Raise vbObjectError + 1001, "No se puede asignar banco, cobro no remesado"
  
  mudtprops.BancoID = mobjBancos.Key(Value)
  
  ' El banco puede ser nulo excepto si esta remesado el cobro/pago.
  mobjValid.RuleBroken "BancoID", (mudtprops.BancoID = 0) And Remesado
  
  mflgDirty = True

End Property

Public Property Get Banco() As String

  Banco = mobjBancos.Item(mudtprops.BancoID)

End Property

Public Property Get Bancos() As TextList

  Set Bancos = mobjBancos

End Property

Public Property Let TemporadaID(Value As Long)

    If Not mflgNew Then Err.Raise 383
  
    mobjValid.RuleBroken "TemporadaID", (Value = 0)

    mudtprops.TemporadaID = Value
    mflgDirty = True

End Property

Public Property Get TemporadaID() As Long

    TemporadaID = mudtprops.TemporadaID
  
End Property

Public Property Let EmpresaID(Value As Long)

    If Not mflgNew Then Err.Raise 383
  
    mobjValid.RuleBroken "EmpresaID", (Value = 0)

    mudtprops.EmpresaID = Value
    mflgDirty = True

End Property

Public Property Get EmpresaID() As Long

    EmpresaID = mudtprops.EmpresaID
  
End Property

Public Property Get SituacionContable() As String
  Select Case mudtprops.SituacionContable
  Case "A"
     SituacionContable = "Generada"
  Case "C"
     SituacionContable = "Contabilizada"
  Case Else
    Err.Raise vbObjectError + 1001, "Situacion contable del cobro/pago erronea"
  End Select

End Property

Public Property Get SituacionComercial() As String
  Select Case mudtprops.SituacionComercial
  Case "A"
     SituacionComercial = "Pendiente"
  Case "S"
     SituacionComercial = "Seleccionado"
  Case "R"
     SituacionComercial = "Remesado"
  Case "C"
     SituacionComercial = "Cobrado"
  Case Else
    Err.Raise vbObjectError + 1001, "Situacion Comercial del cobro/pago erronea"
  End Select

End Property

Public Property Let Tipo(Value As String)

  ' Doy error no solo cuando no estoy editando, sino cuando ademas no es nuevo
  If Not mflgEditing Then _
     If Not mflgNew Then Err.Raise 383
  
  If Len(Value) > Len(mudtprops.Tipo) Then _
    Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
  ' El tipo puede ser UNICAMENTE cobro o pago
  mobjValid.RuleBroken "Tipo", (Trim(Value) <> "C") And (Trim(Value) <> "P")
  
  mudtprops.Tipo = Trim(Value)
  mflgDirty = True

End Property

Public Property Get Tipo() As String

  Tipo = mudtprops.Tipo

End Property

Public Property Let Vencimiento(Value As Variant)

    If Not mflgEditing Then Err.Raise 383
    
    If Len(Trim(Value)) = 0 Then
        mudtprops.Vencimiento = vbNullString
        mobjValid.RuleBroken "Vencimiento", True
    
    ElseIf IsDate(Value) Then
        mudtprops.Vencimiento = CDate(Value)
        mobjValid.RuleBroken "Vencimiento", False
    
    Else
        mobjValid.RuleBroken "Vencimiento", True
    
    End If
    
    ' Comprobar que el vencimiento no es anterior a la fecha de emision
    'He quitado este control, ver que implicaciones tiene
'    If IsDate(mudtprops.Vencimiento) And IsDate(mudtprops.FechaEmision) Then
'        mobjValid.RuleBroken "EmisionVencimiento", mudtprops.Vencimiento < mudtprops.FechaEmision
'    End If
    
    mflgDirty = True
    
End Property

Public Property Get Vencimiento() As Variant

  With mudtprops
    Vencimiento = IIf(.Vencimiento = 0, vbNullString, .Vencimiento)
  End With

End Property

Public Property Let FechaEmision(Value As Variant)
' hay que tratar que la fecha de emision no sea posterior a la fecha de vencimiento.

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Trim(Value)) = 0 Then
    mudtprops.FechaEmision = vbNullString
    mobjValid.RuleBroken "FechaEmision", True

  ElseIf IsDate(Value) Then
    mudtprops.FechaEmision = CDate(Value)
    mobjValid.RuleBroken "FechaEmision", False

  Else
    mobjValid.RuleBroken "FechaEmision", True

  End If
  
  ' Comprobar que el vencimiento no es anterior a la fecha de emision
  ' quito este control.
'  If IsDate(mudtprops.Vencimiento) And IsDate(mudtprops.FechaEmision) Then
'     mobjValid.RuleBroken "EmisionVencimiento", mudtprops.Vencimiento < mudtprops.FechaEmision
'  End If

  mflgDirty = True

End Property

Public Property Get FechaEmision() As Variant

  With mudtprops
    FechaEmision = IIf(.FechaEmision = 0, vbNullString, .FechaEmision)
  End With

End Property

Public Property Let FacturaID(Value As Long)
  If Not mflgEditing Then Err.Raise 383
  
  If FacturaID < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"

  mudtprops.FacturaID = Value
  
  mobjValid.RuleBroken "FacturaID", (FacturaID <= 0)
  mflgDirty = True
  
End Property
  
Public Property Get FacturaID() As Long

  FacturaID = mudtprops.FacturaID

End Property

Public Property Let NumeroGiro(Value As Long)
  If Not mflgEditing Then Err.Raise 383
  
  If NumeroGiro < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"

  mudtprops.NumeroGiro = Value
  
  mobjValid.RuleBroken "NumeroGiro", (NumeroGiro <= 0)
  mflgDirty = True
  
End Property
  
Public Property Get NumeroGiro() As Long

    NumeroGiro = mudtprops.NumeroGiro

End Property

Public Property Get Contabilizado() As Boolean
  
    Contabilizado = (mudtprops.SituacionContable = "C")
  
End Property

Public Property Get Remesado() As Boolean
  
    Remesado = (mudtprops.SituacionComercial = "R")
  
End Property

Public Property Get Seleccionado() As Boolean
  
    Seleccionado = (mudtprops.SituacionComercial = "S")
  
End Property

Public Property Get Cobrado() As Boolean
  
    Cobrado = (mudtprops.SituacionComercial = "C")
  
End Property

Public Property Let Importe(Value As Double)
  
  If Not mflgEditing Then Err.Raise 383
    
'  If Value < 0 Then _
'     Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
'
  mobjValid.RuleBroken "Importe", Value = 0

  If EsEUR(mMoneda) Then
     mudtprops.ImporteEUR = Round(Value, 2)
     mudtprops.ImportePTA = EUR2PTA(Value)
  ElseIf EsPTA(mMoneda) Then
     mudtprops.ImportePTA = Round(Value, 0)
     mudtprops.ImporteEUR = PTA2EUR(Value)
  Else
     mudtprops.ImportePTA = 0
     mudtprops.ImporteEUR = 0
  End If
  
  mflgDirty = True

End Property
Public Property Get Importe() As Double

  If EsEUR(mMoneda) Then
     Importe = mudtprops.ImporteEUR
  ElseIf EsPTA(mMoneda) Then
     Importe = mudtprops.ImportePTA
  Else
     Importe = 0
  End If
  
End Property
Friend Function GetState() As String

  Dim udtData As CobroPagoData
  
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
    Dim udtData As CobroPagoData
    
    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    mobjValid.RuleBroken "Tipo", False
    mobjValid.RuleBroken "NumeroGiro", False
    mobjValid.RuleBroken "Importe", False
    mobjValid.RuleBroken "PersonaID", False
    mobjValid.RuleBroken "FormaPagoID", False
    mobjValid.RuleBroken "EmpresaID", False
    mobjValid.RuleBroken "TemporadaID", False
  
End Sub


Public Sub MarcarContabilizado()
  
    If Not mflgEditing Then Err.Raise 383
      
    If mudtprops.SituacionContable = "C" Then _
       Err.Raise vbObjectError + 1010, "El " & TipoText & " ya esta marcado como contabilizado"
       
    mudtprops.SituacionContable = "C"
    
    mflgDirty = True

End Sub

Private Sub MarcarSeleccionado()

    If Not mflgEditing Then Err.Raise 383
    
    If mudtprops.Tipo <> "C" Then _
       Err.Raise vbObjectError + 1010, "Solamente se pueden seleccionar para remesas los cobros pendientes"
    
    If mudtprops.SituacionComercial <> "A" And mudtprops.SituacionComercial <> "S" Then _
       Err.Raise vbObjectError + 1010, "Solamente se pueden seleccionar para remesas los cobros pendientes"
    
    mudtprops.SituacionComercial = "S"
      
    mflgDirty = True

End Sub

Public Sub MarcarCobrado()
  
    If Not mflgEditing Then Err.Raise 383
    
    If mudtprops.Tipo = "C" Then
      ' Cobros
        If mudtprops.SituacionComercial = "C" Then _
            Err.Raise vbObjectError + 1010, "El cobro ya esta marcado como cobrado"
           
        ' Ojo: referencia a un valor de la base de datos (forma de pago contado):
        ' El cobro puede marcarse como cobrado si era un pago a contado, o
        ' si es un pago remesado
        If mudtprops.SituacionComercial = "A" And mudtprops.FormaPagoID <> 1 Then _
            Err.Raise vbObjectError + 1010, "El cobro debe ser remesado antes de cobrarlo"
        
        If mudtprops.SituacionComercial = "S" Then _
           Err.Raise vbObjectError + 1010, "El cobro debe ser comunicado al banco antes de cobrarlo"
        
        If mudtprops.SituacionComercial = "R" Then _
           Err.Raise vbObjectError + 1010, "El cobro está remesado, se supone cobrardo"
    End If
       
    mudtprops.SituacionComercial = "C"
    
    mflgDirty = True
    
End Sub


Public Property Get FechaDomiciliacion() As Variant

  With mudtprops
    FechaDomiciliacion = IIf(.FechaDomiciliacion = 0, vbNullString, .FechaDomiciliacion)
  End With

End Property

'Public Property Let FechaContable(Value As Variant)
'
'  If Not mflgEditing Then Err.Raise 383
'
'  If Not Contabilizado Then _
'     Err.Raise vbObjectError + 1001, "No se puede asignar fecha contable, cobro no remesado"
'
'  If Len(Trim(Value)) = 0 Then
'    mudtProps.FechaContable = vbNullString
'    mobjValid.RuleBroken "FechaContable", True
'
'  ElseIf IsDate(Value) Then
'    mudtProps.FechaContable = cDate(Value)
'    mobjValid.RuleBroken "FechaContable", False
'
'  Else
'    mobjValid.RuleBroken "FechaContable", True
'
'  End If
'
'  mflgDirty = True
'
'End Property

Public Property Get FechaContable() As Variant

  With mudtprops
    FechaContable = IIf(.FechaContable = 0, vbNullString, .FechaContable)
  End With

End Property

Public Property Get HayFactura() As Boolean
  
  HayFactura = (mudtprops.FacturaID <> 0)
  
End Property

Public Property Get Persona() As String

    If mudtprops.Tipo = "C" Then
       Persona = GetCliente
    ElseIf mudtprops.Tipo = "P" Then
       Persona = GetProveedor
    Else
       Err.Raise vbObjectError + 1010, "Tipo de cobro/pago incorrecto"
    End If
       
End Property

Public Property Get TipoText() As String

    If mudtprops.Tipo = "C" Then
       TipoText = "cobro"
    ElseIf mudtprops.Tipo = "P" Then
       TipoText = "pago"
    Else
       Err.Raise vbObjectError + 1010, "Tipo de cobro/pago incorrecto"
    End If
       
End Property

Public Property Get Personas() As TextList

    If mudtprops.Tipo = "C" Then
       Set Personas = GetClientes
    ElseIf mudtprops.Tipo = "P" Then
       Set Personas = GetProveedores
    Else
       Err.Raise vbObjectError + 1010, "Tipo de cobro/pago incorrecto"
    End If
       
End Property

Public Property Let Persona(Value As String)

    If mudtprops.Tipo = "C" Then
       LetCliente (Value)
    ElseIf mudtprops.Tipo = "P" Then
       LetProveedor (Value)
    Else
       Err.Raise vbObjectError + 1010, "Tipo de cobro/pago incorrecto"
    End If
       
End Property

'Establece los datos de una remesa:
' - Banco
' - Fecha de domiciliacion
Friend Sub SetDatosRemesa(BancoID As Long, FechaDomiciliacion As Variant)
    
    If Not mflgEditing Then Err.Raise 383
    
    If Remesado Then _
       Err.Raise vbObjectError + 1001, "No se puede asignar fecha y banco de domiciliación, cobro ya remesado"
    
    If Contabilizado Then _
       Err.Raise vbObjectError + 1001, "No se puede asignar fecha y banco de domiciliación, cobro ya contabilizado"
    
    If Len(Trim(FechaDomiciliacion)) = 0 Then
        Err.Raise vbObjectError + 1001, "No se puede asignar fecha de domiciliación, fecha no válida"
        
    ElseIf IsDate(FechaDomiciliacion) Then
        mudtprops.FechaDomiciliacion = CDate(FechaDomiciliacion)
      
    Else
        Err.Raise vbObjectError + 1001, "No se puede asignar fecha de domiciliación, fecha no válida"
        
    End If
    
    If BancoID = 0 Then Err.Raise vbObjectError + 1001, "No se puede asignar banco de domiciliación, banco no válido"
    
    mudtprops.BancoID = BancoID
    
    ' Al asignar los datos de la remesa establecemos la situacion en seleccionada
    
    MarcarSeleccionado
    
    mflgDirty = True
    
End Sub

'Marca un cobro que estaba seleccionado como remesado.
Friend Sub MarcarRemesado()
    
    If Not mflgEditing Then Err.Raise 383
    
    If Not (Remesado Or Seleccionado) Then _
       Err.Raise vbObjectError + 1001, "No se puede asignar fecha y banco de domiciliación, cobro ya remesado"
    
    mudtprops.SituacionComercial = "R"
      
    mflgDirty = True
    
End Sub

