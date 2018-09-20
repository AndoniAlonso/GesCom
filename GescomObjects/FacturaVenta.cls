VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturaVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As FacturaVentaProps
Private mudtSave As FacturaVentaProps

' Aqui declaro la comision (depende del representante) a aplicar a las lineas de pedido
' nuevas
Private mComision As Double

Private mobjFacturaVentaItems As FacturaVentaItems

Private mobjCobrosPagos As CobrosPagos

' Aqui declaro el objeto 'child' dato comercial
Private mobjDatoComercial As DatoComercial

' Aqui declaro los ComboBox de la lista de clientes, representantes, transportistas y formas de pago.
Private mobjClientes As TextList
Private mobjRepresentantes As TextList
Private mobjTransportistas As TextList
Private mobjFormasPago As TextList
Private mobjMediosPago As TextList

' Aquí declaro el objeto empresa para ejecutar las reglas de forma de pago y descuento para empresas con anulación de IVA.
Private mobjEmpresa As Empresa

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
    
    mobjFacturaVentaItems.BeginEdit
    mobjCobrosPagos.BeginEdit "EUR"
    mobjDatoComercial.BeginEdit
    
    mflgEditing = True
    
End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    ' restore object state
    LSet mudtprops = mudtSave
    mobjFacturaVentaItems.CancelEdit
    mobjCobrosPagos.CancelEdit
    mobjDatoComercial.CancelEdit
    
End Sub

Public Sub ApplyEdit()
    Dim objPersist As FacturaVentaPersist
    Dim strItems As String
    Dim strCobrosPagos As String
    
    If Not mflgEditing Then Err.Raise 445
    
    Set objPersist = CreateObject("GCServerMTS.FacturaVentaPersist", PERSIST_SERVER)
    
    mobjFacturaVentaItems.ApplyEdit mComision
    
    If mflgDeleted And Not mflgNew Then
        mobjFacturaVentaItems.Actualizar
        
        If Me.FacturaVentaIDAB = 0 Then
            ' code to delete the object's data goes here
            objPersist.DeleteObject mudtprops.FacturaVentaID, mudtprops.DatoComercialID, mobjFacturaVentaItems.GetState, mobjCobrosPagos.GetState
        Else
            Dim objFacturaComplementaria As FacturaVenta
            Set objFacturaComplementaria = New FacturaVenta
            objFacturaComplementaria.Load mudtprops.FacturaVentaIDAB
            objFacturaComplementaria.BeginEdit
            objFacturaComplementaria.Delete
            objPersist.DeleteFacturaAB mudtprops.FacturaVentaID, mudtprops.DatoComercialID, mobjFacturaVentaItems.GetState, mobjCobrosPagos.GetState, _
                                       objFacturaComplementaria.FacturaVentaID, objFacturaComplementaria.DatoComercialID, objFacturaComplementaria.FacturaVentaItems.GetState, objFacturaComplementaria.CobrosPagos.GetState
        End If
        
        mobjDatoComercial.Delete
        mobjDatoComercial.ApplyEdit
        
        mflgNew = True
        mflgDeleted = False
    
    ElseIf IsDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
            mudtprops.DatoComercial.Buffer = mobjDatoComercial.GetState
            mobjFacturaVentaItems.Actualizar
            
            strItems = mobjFacturaVentaItems.GetState
            strCobrosPagos = mobjCobrosPagos.GetState
            SetState objPersist.Save(GetState, strItems, strCobrosPagos)
            mobjFacturaVentaItems.SetState (strItems)
            mobjCobrosPagos.SetState (strCobrosPagos)
            
            mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
            mobjDatoComercial.ApplyEdit
            mudtprops.DatoComercialID = mobjDatoComercial.DatoComercialID
            
            ' save object state
            LSet mudtSave = mudtprops
            mflgNew = False
    Else
        mobjDatoComercial.ApplyEdit
    
    End If
    
    mobjCobrosPagos.ApplyEdit
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
    
    Set mobjDatoComercial = New DatoComercial
    
    mudtprops.TemporadaID = 0
    mudtprops.EmpresaID = 0
    mudtprops.Fecha = Date
    mudtprops.Numero = 0
    mudtprops.NuestraReferencia = vbNullString
    mudtprops.SuReferencia = vbNullString
    mudtprops.Observaciones = vbNullString
    mudtprops.SituacionContable = "A"    ' Marco la situacion contable de la Factura como "A"--> alta,
                                         ' pendiente de codificar las posibles situaciones.
    mudtprops.Bultos = 0
    mudtprops.PesoNeto = 0
    mudtprops.PesoBruto = 0
    mudtprops.BrutoEUR = 0
    mudtprops.BrutoPTA = 0
    mudtprops.DescuentoEUR = 0
    mudtprops.DescuentoPTA = 0
    mudtprops.PortesEUR = 0
    mudtprops.PortesPTA = 0
    mudtprops.EmbalajesEUR = 0
    mudtprops.EmbalajesPTA = 0
    mudtprops.BaseImponibleEUR = 0
    mudtprops.BaseImponiblePTA = 0
    mudtprops.IVAEUR = 0
    mudtprops.IVAPTA = 0
    mudtprops.RecargoEUR = 0
    mudtprops.RecargoPTA = 0
    mudtprops.NetoEUR = 0
    mudtprops.NetoPTA = 0
    mudtprops.FacturaVentaIDAB = 0
      
    Set mobjClientes = New TextList
    mobjClientes.Load "Clientes"
    mudtprops.ClienteID = mobjClientes.Key(mobjClientes.IndexItem(1))
    
    Set mobjRepresentantes = New TextList
    mobjRepresentantes.Load "Representantes"
    mudtprops.RepresentanteID = mobjRepresentantes.Key(mobjRepresentantes.IndexItem(1))
    
    Set mobjTransportistas = New TextList
    mobjTransportistas.Load "Transportistas"
    mudtprops.TransportistaID = mobjTransportistas.Key(mobjTransportistas.IndexItem(1))
    
    Set mobjFormasPago = New TextList
    mobjFormasPago.Load "FormasDePago"
    mudtprops.FormaPagoID = mobjFormasPago.Key(mobjFormasPago.IndexItem(1))
    
    Set mobjMediosPago = New TextList
    mobjMediosPago.Load "MediosPago"
    mudtprops.MedioPagoID = mobjMediosPago.Key(mobjMediosPago.IndexItem(1))
  
    mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "EmpresaID", True
    mobjValid.RuleBroken "ClienteID", True
    mobjValid.RuleBroken "RepresentanteID", True
    mobjValid.RuleBroken "TransportistaID", True
    mobjValid.RuleBroken "FormaPagoID", True
    mobjValid.RuleBroken "MedioPagoID", True
    
    Set mobjFacturaVentaItems = New FacturaVentaItems
    Set mobjCobrosPagos = New CobrosPagos
    
    mComision = 0

End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjDatoComercial = Nothing
    Set mobjClientes = Nothing
    Set mobjRepresentantes = Nothing
    Set mobjTransportistas = Nothing
    Set mobjFormasPago = Nothing
    Set mobjFacturaVentaItems = Nothing
    Set mobjCobrosPagos = Nothing
    Set mobjEmpresa = Nothing
    Set mobjMediosPago = Nothing
    
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

Public Sub Load(FacturaVentaID As Long)
    Dim objPersist As FacturaVentaPersist
    
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.FacturaVentaPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(FacturaVentaID)
    Set objPersist = Nothing
    
    mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
    
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "EmpresaID", False
    mobjValid.RuleBroken "ClienteID", False
    mobjValid.RuleBroken "RepresentanteID", False
    mobjValid.RuleBroken "TransportistaID", False
    mobjValid.RuleBroken "FormaPagoID", False
    mobjValid.RuleBroken "MedioPagoID", False
    
    mobjFacturaVentaItems.Load FacturaVentaID
    mobjCobrosPagos.Load FacturaVentaID, "C", "EUR"
    
    LeerComision
    
    mflgNew = False
    
End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
    
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjCobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    If Contabilizado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha contabilizado la factura"
       
'    If Me.FacturaVentaIDAB <> 0 Then _
'        Err.Raise vbObjectError + 1001, "Factura contabilizada también en la otra empresa, se debe eliminar la factura complementaria"
    
    mflgDeleted = True
    mflgDirty = True
    
    mobjFacturaVentaItems.Delete
    mobjCobrosPagos.Delete
    mobjDatoComercial.Delete
    
End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty Or mobjDatoComercial.IsDirty Or mobjFacturaVentaItems.IsDirty Or mobjCobrosPagos.IsDirty
    
End Property

Public Property Get FacturaVentaID() As Long

    FacturaVentaID = mudtprops.FacturaVentaID
  
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

Public Property Get TerminalID() As Long

    TerminalID = mudtprops.TerminalID
  
End Property

Public Property Let TerminalID(Value As Long)

    If Not mflgNew Then Err.Raise 383
    
    mudtprops.TerminalID = Value
    mflgDirty = True
    
End Property

Public Property Get CentroGestionID() As Long

    CentroGestionID = mudtprops.CentroGestionID
  
End Property

Public Property Let CentroGestionID(Value As Long)

    If Not mflgNew Then Err.Raise 383
    
    mudtprops.CentroGestionID = Value
    mflgDirty = True
    
End Property

Public Property Get AlmacenID() As Long

    AlmacenID = mudtprops.AlmacenID
  
End Property

Public Property Let AlmacenID(Value As Long)

    If Not mflgNew Then Err.Raise 383
    
    mudtprops.AlmacenID = Value
    mflgDirty = True
    
End Property

Public Property Get FacturaVentaIDAB() As Long

    FacturaVentaIDAB = mudtprops.FacturaVentaIDAB
  
End Property

Public Property Get HayFacturaComplementaria() As Boolean

    HayFacturaComplementaria = (mudtprops.FacturaVentaIDAB <> 0)
    
End Property

Public Property Let Bruto(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjCobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    mudtprops.BrutoEUR = Round(Value, 2)
    mudtprops.BrutoPTA = EUR2PTA(Value)

    CalcularDescuento
      
    mflgDirty = True
    
End Property

Public Property Get Bruto() As Double

    Bruto = mudtprops.BrutoEUR

End Property
Public Property Let Descuento(Value As Double)
    
    If Not mflgEditing Then Err.Raise 383
      
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjCobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    mudtprops.DescuentoEUR = Round(Value, 2)
    mudtprops.DescuentoPTA = EUR2PTA(Value)
      
    CalcularBaseImponible
    
    mflgDirty = True

End Property

Public Property Get Descuento() As Double

    Descuento = mudtprops.DescuentoEUR

End Property

Public Property Let Portes(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
    If Value < 0 Then _
       Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
    
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjCobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    mobjValid.RuleBroken "Portes", (Value < 0)
    
    mudtprops.PortesEUR = Round(Value, 2)
    mudtprops.PortesPTA = EUR2PTA(Value)
    
    CalcularBaseImponible
    
    mflgDirty = True
    
End Property

Public Property Get Portes() As Double

    Portes = mudtprops.PortesEUR

End Property

Public Property Let Embalajes(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
    If Value < 0 Then _
       Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
    
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjCobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    mobjValid.RuleBroken "Embalajes", (Value < 0)
    
    mudtprops.EmbalajesEUR = Round(Value, 2)
    mudtprops.EmbalajesPTA = EUR2PTA(Value)
      
    CalcularBaseImponible
    
    mflgDirty = True
    
End Property

Public Property Get Embalajes() As Double

    Embalajes = mudtprops.EmbalajesEUR
 
End Property

Public Property Let BaseImponible(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjCobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    mudtprops.BaseImponibleEUR = Round(Value, 2)
    mudtprops.BaseImponiblePTA = EUR2PTA(Value)
    
    CalcularIVA
    CalcularRecargo
    
    mflgDirty = True
    
End Property

Public Property Get BaseImponible() As Double

    BaseImponible = mudtprops.BaseImponibleEUR
 
End Property

Public Property Let IVA(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjCobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    mudtprops.IVAEUR = Round(Value, 2)
    mudtprops.IVAPTA = EUR2PTA(Value)
    
    CalcularNeto
    
    mflgDirty = True

End Property

Public Property Get IVA() As Double

    IVA = mudtprops.IVAEUR

End Property

Public Property Let Recargo(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjCobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    mudtprops.RecargoEUR = Round(Value, 2)
    mudtprops.RecargoPTA = EUR2PTA(Value)
    
    CalcularNeto
    
    mflgDirty = True

End Property

Public Property Get Recargo() As Double

    Recargo = mudtprops.RecargoEUR
 
End Property

Public Property Let Neto(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjCobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    mudtprops.NetoEUR = Round(Value, 2)
    mudtprops.NetoPTA = EUR2PTA(Value)
    
    mflgDirty = True

End Property

Public Property Get Neto() As Double

    Neto = mudtprops.NetoEUR
 
End Property

Public Property Get Numero() As Long

    Numero = mudtprops.Numero
  
End Property

Public Property Let Numero(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"
  
  mobjValid.RuleBroken "Numero", (Value < 0)
  
  mudtprops.Numero = Value
  
  mflgDirty = True

End Property

Public Property Get ClienteID() As Long

  ClienteID = mudtprops.ClienteID
  
End Property

' Esta funcion lee el cliente para asignar sus propiedades por defecto
Public Property Let Cliente(Value As String)
    Dim objCliente As Cliente
    Dim objEmpresa As Empresa
    
    If Not mflgEditing Then Err.Raise 383
    
    mobjValid.RuleBroken "ClienteID", (mobjClientes.Key(Value) = 0)
      
    'Salimos si queremos asignar algo invalido
    If mobjClientes.Key(Value) = 0 Then Exit Property
       
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjCobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    ' Devolvemos error si ya se ha contabilizado y queremos cambiar de cliente.
    If Me.Contabilizado Then _
       Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, ya se ha contabilizado"
       
    ' Devolvemos error si ya se han contabilizado los cobros y queremos cambiar de cliente.
    If mobjCobrosPagos.Contabilizado Then _
       Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, ya se han contabilizado los cobros"
       
    ' Devolvemos error si ya se han remesado los cobros y queremos cambiar de cliente.
    If mobjCobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, ya se han remesado los cobros"
       
    ' Devolvemos error si se han asignado a partir de Albaranes (por lo tanto pertenecen a un cliente).
    If mobjFacturaVentaItems.HayAlbaran Then _
       Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, ya se han asignado Albaranes pendientes"
       
    ' Cambiamos de cliente
    Set objCliente = New Cliente
    
    objCliente.Load mobjClientes.Key(Value)
    
    mudtprops.RepresentanteID = objCliente.RepresentanteID
    CambiarComision
    
    mudtprops.TransportistaID = objCliente.TransportistaID
    
    'OJOOJO El medio de pago se puede tomar de un parametro o del cliente!!!
    mudtprops.MedioPagoID = 1
    mobjValid.RuleBroken "MedioPagoID", False
    
    ' Si la empresa no ha sido cargada, lo hacemos.
    If mobjEmpresa Is Nothing Then _
        CargarDatosEmpresa
        
    ' Si la empresa tiene la condicion de anular IVA, se pone la forma de pago al contado.
    If mobjEmpresa.AnularIVA Then
        ' Ojo: referencia a un valor de la base de datos (forma de pago contado):
        mudtprops.FormaPagoID = 1
    Else
        mudtprops.FormaPagoID = objCliente.FormaPagoID
    End If
    
    ' Asignamos unos u otros datos comerciales en funcion de si en la empresa se anula el IVA.
    With mobjDatoComercial
        .ChildBeginEdit
        If mobjEmpresa.AnularIVA Then
            .Descuento = objCliente.DatoComercialB.Descuento
            .RecargoEquivalencia = objCliente.DatoComercialB.RecargoEquivalencia
            .IVA = objCliente.DatoComercialB.IVA
        Else
            .Descuento = objCliente.DatoComercial.Descuento
            .RecargoEquivalencia = objCliente.DatoComercial.RecargoEquivalencia
            .IVA = objCliente.DatoComercial.IVA
        End If
        mobjDatoComercial.ChildApplyEdit
    End With
    
    mudtprops.ClienteID = objCliente.ClienteID
    
    CambiarComision
    
    Set objEmpresa = Nothing
    Set objCliente = Nothing
    
    mflgDirty = True

End Property

Public Property Get Cliente() As String

    Cliente = mobjClientes.Item(mudtprops.ClienteID)

End Property

Public Property Get Clientes() As TextList

  Set Clientes = mobjClientes

End Property

Public Property Let Fecha(Value As Variant)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Trim(Value)) = 0 Then
    mudtprops.Fecha = vbNullString
    mobjValid.RuleBroken "Fecha", True

  ElseIf IsDate(Value) Then
    mudtprops.Fecha = CDate(Value)
    mobjValid.RuleBroken "Fecha", False

  Else
    mobjValid.RuleBroken "Fecha", True

  End If

  mflgDirty = True

End Property

Public Property Get Fecha() As Variant

  With mudtprops
    Fecha = IIf(.Fecha = 0, vbNullString, .Fecha)
  End With

End Property

Public Property Let NuestraReferencia(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtprops.NuestraReferencia) Then _
    Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
  mudtprops.NuestraReferencia = Trim(Value)
  mflgDirty = True

End Property

Public Property Get NuestraReferencia() As String

  NuestraReferencia = Trim(mudtprops.NuestraReferencia)

End Property

Public Property Let SuReferencia(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  If Len(Value) > Len(mudtprops.SuReferencia) Then _
    Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
  mudtprops.SuReferencia = Trim(Value)
  mflgDirty = True

End Property

Public Property Get SuReferencia() As String

  SuReferencia = Trim(mudtprops.SuReferencia)

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

Public Property Get SituacionContable() As String
  
    Select Case mudtprops.SituacionContable
    Case "A"
       SituacionContable = "Generada"
    Case "C"
       SituacionContable = "Contabilizada"
    Case Else
      Err.Raise vbObjectError + 1001, "Situacion contable de Factura erronea"
    End Select

End Property

Public Property Get Contabilizado() As Boolean

    Contabilizado = (mudtprops.SituacionContable = "C")
  
End Property

Public Sub Contabilizar(Optional ForzarContabilidad As Boolean = False)
    Dim objAsiento As Asiento
    Dim objCliente As Cliente
    Dim strDescripcion As String
    Dim objPersist As FacturaVentaPersist
    Dim strCuentaBase As String
    Dim strTitular As String
    Dim strCuentaIVA As String
    Dim strCuentaRE As String

    
    'If Not mflgEditing Then Err.Raise 383
    If mflgEditing Then Err.Raise 383
    
    ' Realizamos la comprobación de si ya está contabilizada, y damos error si no se ha
    ' marcada la contabilidad como forzada.
    If Contabilizado And _
       Not ForzarContabilidad Then _
        Err.Raise vbObjectError + 1010, "La Factura ya esta marcada como contabilizada"
     
    mudtprops.SituacionContable = "C"
    
    ' Aquí por fin contabilizo
    Set objAsiento = New Asiento
    Set objCliente = New Cliente
    
    If Me.Neto = 0 Then Exit Sub
    
    objCliente.Load Me.ClienteID
    
    
    With objAsiento
        If Me.Neto > 0 Then
            strDescripcion = "VENTAS S/FRA:" & Me.Numero
        Else
            strDescripcion = "DEVOLUCION VENTAS FRA:" & Me.Numero
        End If
        
        objAsiento.NuevoAsiento Me.Numero, CStr(Year(Me.Fecha)), "Factura Venta nº " & Me.Numero, _
                                Me.TemporadaID, Me.EmpresaID
                                
        .NuevoApunte objCliente.CuentaContable, "Debe", _
                               Me.Neto, strDescripcion, _
                               Me.Fecha, Me.Numero
        .NuevoApunte "6650000001", "Debe", _
                               Me.Descuento, strDescripcion, _
                               Me.Fecha, Me.Numero
        ' Las cuentas de IVA se derivan según el porcentaje de IVA
        ' - 477.16 --> el 16%,  etc.
        strCuentaIVA = "47700000" & Format(Me.DatoComercial.IVA, "00")
        strCuentaRE = "47700000" & Format(Me.DatoComercial.RecargoEquivalencia, "00")
        .NuevoApunte strCuentaIVA, "Haber", _
                               Me.IVA, strDescripcion, _
                               Me.Fecha, Me.Numero
        .NuevoApunte strCuentaRE, "Haber", _
                               Me.Recargo, strDescripcion, _
                               Me.Fecha, Me.Numero
                               
        If Me.Neto > 0 Then
            strCuentaBase = "7000000001"
        Else
            strCuentaBase = "7080000001"
        End If
        
        .NuevoApunte strCuentaBase, "Haber", _
                            Me.Bruto, strDescripcion, _
                            Me.Fecha, Me.Numero
        ' Apuntes de IVA.
        If Me.IVA <> 0 Or Me.Recargo <> 0 Then
            strTitular = IIf(objCliente.Titular = vbNullString, objCliente.Nombre, objCliente.Titular)
            .NuevoIVAApunte "R", Me.Numero, strCuentaIVA, strCuentaBase, objCliente.CuentaContable, _
                          strTitular, objCliente.DNINIF, Me.BaseImponible, _
                        Me.Neto, Me.DatoComercial.IVA, Me.IVA, Me.DatoComercial.RecargoEquivalencia, Me.Recargo
        End If
        
        .NuevoApunte "7000000002", "Haber", _
                               Me.Portes, strDescripcion, _
                               Me.Fecha, Me.Numero
        .NuevoApunte "7000000003", "Haber", _
                               Me.Embalajes, strDescripcion, _
                               Me.Fecha, Me.Numero
        .CuadrarAsiento objCliente.CuentaContable, Me.Fecha, Me.Numero
        
        
    End With
    
    Set objPersist = CreateObject("GCServerMTS.FacturaVentaPersist", PERSIST_SERVER)
    objPersist.Contabilizar mudtprops.FacturaVentaID, _
                            mudtprops.SituacionContable, _
                            objAsiento.GetSuperState
    Set objPersist = Nothing

    Set objAsiento = Nothing
    Set objCliente = Nothing
    
End Sub

Public Sub DesContabilizar()
    
    If Not mflgEditing Then Err.Raise 383
    
    ' Realizamos la comprobación de si no está contabilizada, y damos error si no se ha
    ' marcada la contabilidad como forzada.
    If Not Contabilizado Then _
        Err.Raise vbObjectError + 1010, "La Factura no esta marcada como contabilizada"
     
    mudtprops.SituacionContable = "A"
    mflgDirty = True
    
End Sub

Public Property Get Bultos() As Long

  Bultos = mudtprops.Bultos
  
End Property

Public Property Let Bultos(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"
  
  mobjValid.RuleBroken "Bultos", (Value < 0)
  
  mudtprops.Bultos = Value
  
  mflgDirty = True

End Property

Public Property Get PesoNeto() As Long

  PesoNeto = mudtprops.PesoNeto
  
End Property

Public Property Let PesoNeto(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"
  
  mobjValid.RuleBroken "PesoNeto", (Value < 0)
  
  mudtprops.PesoNeto = Value
  
  mflgDirty = True

End Property

Public Property Get PesoBruto() As Long

  PesoBruto = mudtprops.PesoBruto
  
End Property

Public Property Let PesoBruto(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"
  
  mobjValid.RuleBroken "PesoBruto", (Value < 0)
  
  mudtprops.PesoBruto = Value
  
  mflgDirty = True

End Property

Public Property Get RepresentanteID() As Long

  RepresentanteID = mudtprops.RepresentanteID
  
End Property

Public Property Let Representante(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  mudtprops.RepresentanteID = mobjRepresentantes.Key(Value)
  
  mobjValid.RuleBroken "RepresentanteID", (mudtprops.RepresentanteID = 0)
  
  CambiarComision
  
  mflgDirty = True

End Property

Public Property Get Representante() As String

  Representante = mobjRepresentantes.Item(mudtprops.RepresentanteID)

End Property

Public Property Get Representantes() As TextList

  Set Representantes = mobjRepresentantes

End Property

' Obtiene la comision del representante
Private Sub LeerComision()
    Dim objRepresentante As Representante
    
    If mudtprops.RepresentanteID = 0 Then Exit Sub
    
    ' Cargamos el representante para asignar su comision.
    Set objRepresentante = New Representante
    objRepresentante.Load mudtprops.RepresentanteID
    
    mComision = objRepresentante.Comision
    
    Set objRepresentante = Nothing

End Sub

Public Sub CambiarComision()

    LeerComision
    mobjFacturaVentaItems.CambiarComision (mComision)
    
End Sub

Public Property Get ImporteComision() As Double

  ImporteComision = mobjFacturaVentaItems.ImporteComision
  
End Property

Public Property Get TransportistaID() As Long

  TransportistaID = mudtprops.TransportistaID
  
End Property

Public Property Let Transportista(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
  mudtprops.TransportistaID = mobjTransportistas.Key(Value)
  
  mobjValid.RuleBroken "TransportistaID", (mudtprops.TransportistaID = 0)
  
  mflgDirty = True

End Property

Public Property Get Transportista() As String

  Transportista = mobjTransportistas.Item(mudtprops.TransportistaID)

End Property

Public Property Get Transportistas() As TextList

  Set Transportistas = mobjTransportistas

End Property

Public Property Get FormaPagoID() As Long

  FormaPagoID = mudtprops.FormaPagoID
  
End Property

Public Property Let FormaPago(Value As String)

  If Not mflgEditing Then Err.Raise 383
  
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
  
    mobjValid.RuleBroken "MedioPagoID", (mudtprops.MedioPagoID = 0)
  
    mflgDirty = True

End Property

Public Property Get MedioPago() As String

    MedioPago = mobjMediosPago.Item(mudtprops.MedioPagoID)

End Property

Public Property Get MediosPago() As TextList

    Set MediosPago = mobjMediosPago

End Property

Public Property Get FacturaVentaItems() As FacturaVentaItems

  Set FacturaVentaItems = mobjFacturaVentaItems

End Property

Public Property Get CobrosPagos() As CobrosPagos

  Set CobrosPagos = mobjCobrosPagos

End Property

Public Property Let DatoComercialID(Value As Long)
  
  If Not mflgEditing Then Err.Raise 383
  
  mudtprops.DatoComercialID = Value
  mflgDirty = True

End Property

Public Property Get DatoComercialID() As Long

  DatoComercialID = mudtprops.DatoComercialID
  
End Property

Public Property Get DatoComercial() As DatoComercial
  Set DatoComercial = mobjDatoComercial
  
End Property

Public Sub CalcularBruto()
   Bruto = mobjFacturaVentaItems.Bruto
End Sub

Public Sub CalcularDescuento()
   Descuento = Bruto * mobjDatoComercial.Descuento / 100
End Sub

Public Sub CalcularBaseImponible()
   BaseImponible = (Bruto - Descuento + Portes + Embalajes)
End Sub

Public Sub CalcularRecargo()
   Recargo = BaseImponible * mobjDatoComercial.RecargoEquivalencia / 100
End Sub

Public Sub CalcularIVA()
   IVA = BaseImponible * mobjDatoComercial.IVA / 100
End Sub

Public Sub CalcularNeto()
   Neto = BaseImponible + Recargo + IVA
End Sub

Public Sub CrearCobros()
    Dim objFormaPago As FormaDePago
    Dim i As Integer
    Dim FechaGiro As Date
    Dim ImporteGiro As Double
    Dim ImporteAcumulado As Double
    Dim objCobroPago As CobroPago
    Dim objCliente As Cliente
    Dim DiaPago1 As Integer
    Dim DiaPago2 As Integer
    Dim DiaPago3 As Integer
         
    ' Devolvemos error si ya se han remesado los cobros y queremos recalcularlos.
    If mobjCobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se pueden recalcular, ya se han remesado los cobros"
       
    Set objFormaPago = New FormaDePago
    objFormaPago.Load mudtprops.FormaPagoID
  
    ' Eliminamos todos los giros que pudiera haber anteriormente.
    mobjCobrosPagos.Delete
    
    ' Pasamos de calcular la fecha en meses a calcularla en días.
    'FechaGiro = DateAdd("m", objFormaPago.MesesPrimerGiro, mudtprops.Fecha)
    FechaGiro = DateAdd("d", objFormaPago.MesesPrimerGiro * 30, mudtprops.Fecha)
    
    ' Asignar la fecha en función de los días de pago del cliente.
    Set objCliente = New Cliente
    objCliente.Load mudtprops.ClienteID
    DiaPago1 = objCliente.DiaPago1
    DiaPago2 = objCliente.DiaPago2
    DiaPago3 = objCliente.DiaPago3
    Set objCliente = Nothing
    
    ImporteGiro = Neto / objFormaPago.Giros
    ImporteAcumulado = 0
    
    For i = 1 To objFormaPago.Giros
        Set objCobroPago = mobjCobrosPagos.Add
        
        With objCobroPago
            .BeginEdit "EUR"
            .Tipo = "C"
            .Persona = Cliente
            .FormaPago = FormaPago
            .MedioPago = MedioPago
            .NumeroGiro = i
            If i = objFormaPago.Giros Then
                ImporteGiro = Neto - ImporteAcumulado
            End If
            
            ' Se asigna el importe y se devuelve por si las reglas de negocio han
            ' cambiado el valor.
            .Importe = ImporteGiro
            ImporteGiro = .Importe
            
            ImporteAcumulado = ImporteAcumulado + ImporteGiro
            
            ' Se asigna el vencimiento y se devuelve por si las reglas de negocio han
            ' cambiado el valor.
            .Vencimiento = CalcularDiasPago(FechaGiro, DiaPago1, DiaPago2, DiaPago3)
            'FechaGiro = .Vencimiento
            ' Pasamos de calcular la fecha en meses a calcularla en días.
            'FechaGiro = DateAdd("m", objFormaPago.MesesEntreGiros, FechaGiro)
            FechaGiro = DateAdd("d", objFormaPago.MesesEntreGiros * 30, FechaGiro)
            
            .FechaEmision = Fecha
            ' TODO: Ver si asignar el banco o hacerlo al domiciliarlo.
            
            .TemporadaID = TemporadaID
            .EmpresaID = EmpresaID
            
            .ApplyEdit
            
        End With
        
        Set objCobroPago = Nothing
    Next i
    
    Set objFormaPago = Nothing
  
End Sub

Private Sub CargarDatosEmpresa()
    ' Cargamos los datos de la empresa para comprobar sus condiciones de IVA.

    Set mobjEmpresa = New Empresa
    mobjEmpresa.Load mudtprops.EmpresaID

End Sub

Friend Function GetState() As String
    Dim udtData As FacturaVentaData
    
    mudtprops.DatoComercial.Buffer = mobjDatoComercial.GetState
    
    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With
    
    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As FacturaVentaData
    
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub

