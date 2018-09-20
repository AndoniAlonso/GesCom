VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlbaranVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As AlbaranVentaProps
Private mudtSave As AlbaranVentaProps

' Aqui declaro la comision (depende del representante) a aplicar a las lineas de pedido
' nuevas
Private mComision As Double

Private mobjAlbaranVentaItems As AlbaranVentaItems

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

Private mbolDevolucion As Boolean

Event Valid(IsValid As Boolean)

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Private Const cnRayaSeparador = "--------------------------------"
Private Const cnEspacioAreaTexto = 23
'Private Const cnESCP2Escape = &H1B
'Private Const cnESCP2InicializePrinter = &H40
'Private Const cnESCP2ColorSelect = &H72
'Private Const cnESCP2PrintMode = &H21
'Private Const cnESCP2DoubleHeight = &H17
'Private Const cnESCP2DefaultFont = 0
'Private Const cnESCP2FontB = 1
'Private Const cnESCP2BlackColor = &H30
'Private Const cnESCP2RedColor = &H31


Public Sub BeginEdit()

    If mflgEditing Then Err.Raise 445
    
    ' save object state
    LSet mudtSave = mudtprops
    
    mobjAlbaranVentaItems.BeginEdit
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
    mobjAlbaranVentaItems.CancelEdit
    mobjDatoComercial.CancelEdit
    
End Sub

Public Sub ApplyEdit()
    Dim objPersist As AlbaranVentaPersist
    Dim strItems As String
  
    If Not mflgEditing Then Err.Raise 445
    
    Set objPersist = CreateObject("GCServerMTS.AlbaranVentaPersist", PERSIST_SERVER)
    
    TotalBruto = TotalBruto
    
    If mflgDeleted And Not mflgNew Then
        mobjAlbaranVentaItems.Actualizar
        
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.AlbaranVentaID, mudtprops.DatoComercialID, mobjAlbaranVentaItems.GetState
        
        mobjDatoComercial.Delete
        mobjDatoComercial.ApplyEdit
        
        mflgNew = True
        mflgDeleted = False

    ElseIf mflgDirty Or mobjAlbaranVentaItems.IsDirty Or mflgNew Then
            If Not IsValid Then Err.Raise 445
            mudtprops.DatoComercial.Buffer = mobjDatoComercial.GetState
            
            mobjAlbaranVentaItems.Actualizar
            
            strItems = mobjAlbaranVentaItems.GetState
            SetState objPersist.Save(GetState, strItems)
            mobjAlbaranVentaItems.SetState (strItems)
             
            mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
            mobjDatoComercial.ApplyEdit
            mudtprops.DatoComercialID = mobjDatoComercial.DatoComercialID
            
            ' save object state
            LSet mudtSave = mudtprops
            mflgNew = False
    Else
        mobjDatoComercial.ApplyEdit
    
    End If
    
    mobjAlbaranVentaItems.ApplyEdit mComision 'mudtprops.AlbaranVentaID,
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
    mudtprops.Bultos = 0
    mudtprops.PesoNeto = 0
    mudtprops.PesoBruto = 0
    mudtprops.PortesEUR = 0
    mudtprops.PortesPTA = 0
    mudtprops.EmbalajesEUR = 0
    mudtprops.EmbalajesPTA = 0
    mudtprops.TotalBrutoEUR = 0
    mudtprops.TotalBrutoPTA = 0
    mudtprops.AlmacenID = 0
    
    mbolDevolucion = False
    
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
      
    Set mobjAlbaranVentaItems = New AlbaranVentaItems
    
    mComision = 0
        
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjDatoComercial = Nothing
    Set mobjClientes = Nothing
    Set mobjRepresentantes = Nothing
    Set mobjTransportistas = Nothing
    Set mobjFormasPago = Nothing
    Set mobjAlbaranVentaItems = Nothing
    Set mobjMediosPago = Nothing
    
    Set mobjEmpresa = Nothing
    
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

Public Sub Load(AlbaranVentaID As Long)
    Dim objPersist As AlbaranVentaPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.AlbaranVentaPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(AlbaranVentaID)
    Set objPersist = Nothing
    
    mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
    
    UnbrokeRules
    
    mobjAlbaranVentaItems.Load AlbaranVentaID
    
    LeerComision
      
    mflgNew = False
    
End Sub

Private Sub UnbrokeRules()
    
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "EmpresaID", False
    mobjValid.RuleBroken "ClienteID", False
    mobjValid.RuleBroken "RepresentanteID", False
    mobjValid.RuleBroken "TransportistaID", False
    mobjValid.RuleBroken "FormaPagoID", False
    mobjValid.RuleBroken "MedioPagoID", False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
    
    mflgDeleted = True
    mflgDirty = True
    
    mobjAlbaranVentaItems.Delete
    mobjDatoComercial.Delete
    
End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty Or mobjDatoComercial.IsDirty

End Property

Public Property Let EsDevolucion(Value As Boolean)

    mbolDevolucion = Value
    
End Property

Public Property Get EsDevolucion() As Boolean

    EsDevolucion = mbolDevolucion

End Property

Public Property Get AlbaranVentaID() As Long

    AlbaranVentaID = mudtprops.AlbaranVentaID
  
End Property

Public Property Get FacturaVentaIDA() As Long

    FacturaVentaIDA = mudtprops.FacturaVentaIDA
  
End Property

Public Property Get FacturaVentaIDB() As Long

    FacturaVentaIDB = mudtprops.FacturaVentaIDB
  
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

Public Property Let EmpresaID(Value As Long)

    If Not mflgNew Then Err.Raise 383
    
    mobjValid.RuleBroken "EmpresaID", (Value = 0)
    
    mudtprops.EmpresaID = Value
    mflgDirty = True

End Property

Public Property Get EmpresaID() As Long

    EmpresaID = mudtprops.EmpresaID
  
End Property


Public Property Let Portes(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
    
    mobjValid.RuleBroken "Portes", (Value < 0)
    
    mudtprops.PortesEUR = Round(Value, 2)
    mudtprops.PortesPTA = EUR2PTA(Value)
      
    mflgDirty = True
    
End Property

Public Property Get Portes() As Double

    Portes = mudtprops.PortesEUR
    
End Property
Public Property Let Embalajes(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
    
    mobjValid.RuleBroken "Embalajes", (Value < 0)
    
    mudtprops.EmbalajesEUR = Round(Value, 2)
    mudtprops.EmbalajesPTA = EUR2PTA(Value)
      
    mflgDirty = True
    
End Property

Public Property Get Embalajes() As Double

    Embalajes = mudtprops.EmbalajesEUR
 
End Property

Public Property Get Numero() As Long

    Numero = mudtprops.Numero
  
End Property

Public Property Let Numero(Value As Long)

    If Not mflgEditing Then Err.Raise 383
    
    If Value <= 0 Then _
        Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"
    
    mobjValid.RuleBroken "Numero", (Value <= 0)
    
    mudtprops.Numero = Value
    
    mflgDirty = True
    
End Property

Public Property Get ClienteID() As Long

    ClienteID = mudtprops.ClienteID
  
End Property

' Esta funcion lee el cliente para asignar sus propiedades por defecto
Public Property Let Cliente(Value As String)
    Dim objCliente As Cliente

    If Not mflgEditing Then Err.Raise 383
    
    mobjValid.RuleBroken "ClienteID", (mobjClientes.Key(Value) = 0)
      
    'Salimos si queremos asignar algo invalido
    If mobjClientes.Key(Value) = 0 Then Exit Property
       
    ' Devolvemos error si ya ha habido entregas y queremos cambiar de cliente.
    If mobjAlbaranVentaItems.Facturado Then _
        Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, ya se han realizado entregas"
       
    ' Devolvemos error si se han asignado a partir de pedidos (por lo tanto pertenecen a un cliente).
    If mobjAlbaranVentaItems.HayPedido Then _
        Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, ya se han asignado pedidos pendientes"
       
    ' Cambiamos de cliente
    Set objCliente = New Cliente
    
    objCliente.Load mobjClientes.Key(Value)
    
    mudtprops.RepresentanteID = objCliente.RepresentanteID
    mobjValid.RuleBroken "RepresentanteID", (mudtprops.RepresentanteID = 0)
    mudtprops.TransportistaID = objCliente.TransportistaID
    mobjValid.RuleBroken "TransportistaID", (mudtprops.TransportistaID = 0)
    
    'OJOOJO el medio de pago se puede tomar del cliente o de un parametro!!!
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
    mobjValid.RuleBroken "FormaPagoID", (mudtprops.FormaPagoID = 0)
    
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

Public Property Get Bultos() As Long

    Bultos = mudtprops.Bultos
  
End Property

Public Property Let Bultos(Value As Long)

    If Not mflgEditing Then Err.Raise 383
    
    If Value <= 0 Then _
        Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"
    
    mobjValid.RuleBroken "Bultos", (Value <= 0)
    
    mudtprops.Bultos = Value
    
    mflgDirty = True
    
End Property

Public Property Get PesoNeto() As Long

    PesoNeto = mudtprops.PesoNeto
  
End Property

Public Property Let PesoNeto(Value As Long)

    If Not mflgEditing Then Err.Raise 383
    
    If Value <= 0 Then _
        Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"
    
    mobjValid.RuleBroken "PesoNeto", (Value <= 0)
    
    mudtprops.PesoNeto = Value
    
    mflgDirty = True
    
End Property

Public Property Get PesoBruto() As Long

  PesoBruto = mudtprops.PesoBruto
  
End Property

Public Property Let PesoBruto(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  If Value <= 0 Then _
    Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"
  
  mobjValid.RuleBroken "PesoBruto", (Value <= 0)
  
  mudtprops.PesoBruto = Value
  
  mflgDirty = True

End Property

Public Property Let TotalBruto(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
'    If Value < 0 Then _
'       Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
    
'    mobjValid.RuleBroken "TotalBruto", (Value < 0)
    
    mudtprops.TotalBrutoEUR = Round(Value, 2)
    mudtprops.TotalBrutoPTA = EUR2PTA(Value)
      
    mflgDirty = True
    
End Property

Public Property Get TotalBruto() As Double

    TotalBruto = mobjAlbaranVentaItems.Bruto
 
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

Private Sub CambiarComision()

    LeerComision
    mobjAlbaranVentaItems.CambiarComision (mComision)
    
End Sub


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

Public Property Get AlbaranVentaItems() As AlbaranVentaItems

  Set AlbaranVentaItems = mobjAlbaranVentaItems

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

Private Sub CargarDatosEmpresa()
    ' Cargamos los datos de la empresa para comprobar sus condiciones de IVA.

    Set mobjEmpresa = New Empresa
    mobjEmpresa.Load mudtprops.EmpresaID

End Sub


Friend Function GetState() As String

  Dim udtData As AlbaranVentaData

  With mudtprops
    .IsDeleted = mflgDeleted
    .IsNew = mflgNew
    .IsDirty = mflgDirty
  End With

  LSet udtData = mudtprops
  GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

  Dim udtData As AlbaranVentaData
  
  udtData.Buffer = Buffer
  LSet mudtprops = udtData

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
        .WriteProperty "Devolucion", mbolDevolucion
        .WriteProperty "Children", mobjAlbaranVentaItems.GetSuperState
        .WriteProperty "DatoComercial", mobjDatoComercial.GetSuperState
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
        mbolDevolucion = .ReadProperty("Devolucion")
        mobjAlbaranVentaItems.SetSuperState .ReadProperty("Children")
        mobjDatoComercial.SetSuperState .ReadProperty("DatoComercial")
    End With
   
    UnbrokeRules
    
    Set objPB = Nothing
End Function

' A partir de los items de albaran, un articulo a ubicar y la lista de pedidos pendientes,
' ubicar el albaran en una linea:
'1.- Intentar ubicar el artículo en una linea de albaran existente relacionada con un pedido.
'2.- Intentar ubicar el artículo en una nueva linea de albaran, relacionandola con un pedido pendiente.
'3.- Intentar ubicar el artículo en una linea de albarán existente no relacionada con un pedido.
'4.- Ubicar el artículo en una nueva linea no relacionada con pedidos.
' La función devuelve el string informativo del error en caso de que haya alguna incidencia.
Public Function AlbaranItemCodigoBarras(ByVal strArticuloColor As String, ByVal intTalla As Integer, ByRef objPickerItems As PickerItems, _
                                        ByVal ArticuloColorID As Long, ByVal bolPermitirSinPedido As Boolean, ByVal strNombreArticuloColor As String) As String
Dim i As Integer
Dim j As Integer
Dim flgPedidoExistente As Boolean
Dim flgPedidoNuevo As Boolean
Dim flgAlbaranExistente As Boolean
Dim flgAlbaranNuevo As Boolean
Dim flgPedidoYaUtilizado As Boolean
Dim objAlbaranVentaItem As AlbaranVentaItem
Dim objPedidoVentaItem As PedidoVentaItem
Dim dblCantidadActual As Double
Dim dblCantidadFinal As Double
Dim lngPedidoVentaItemID As Long
Dim strResultadoFuncion As String


    strResultadoFuncion = vbNullString
    flgPedidoExistente = False
    For i = 1 To mobjAlbaranVentaItems.Count
        If mobjAlbaranVentaItems(i).CodigoArticuloColor = strArticuloColor And _
           mobjAlbaranVentaItems(i).HayPedido And (Not mobjAlbaranVentaItems(i).IsDeleted) Then
           ' Esta linea es una candidata...
           ' Cargamos el pedido para ver si tiene cantidades pendientes...
            Set objPedidoVentaItem = New PedidoVentaItem
            objPedidoVentaItem.Load mobjAlbaranVentaItems(i).PedidoVentaItemID, "EUR"
            
            ' Llamamos a la propiedad de pendiente, segun la talla, para comprobar si la cantidad pendiente es mayor que la cantidad del albaran
            If CallByName(objPedidoVentaItem, "PendienteT" & CStr(intTalla), VbGet) > _
               CallByName(mobjAlbaranVentaItems(i), "CantidadT" & CStr(intTalla), VbGet) Then
                'Bingo!
                Set objAlbaranVentaItem = mobjAlbaranVentaItems(i)
                flgPedidoExistente = True
                Set objPedidoVentaItem = Nothing
                Exit For
            End If
            Set objPedidoVentaItem = Nothing
        End If
    Next
    
    flgPedidoNuevo = False
    If Not flgPedidoExistente Then
        For i = 1 To objPickerItems.Count
            If objPickerItems(i).Nombre = strArticuloColor Then
                ' Cargamos el pedido para ver si tiene cantidades pendientes...
                Set objPedidoVentaItem = New PedidoVentaItem
                lngPedidoVentaItemID = objPickerItems(i).DocumentoID
                objPedidoVentaItem.Load lngPedidoVentaItemID, "EUR"
                If CallByName(objPedidoVentaItem, "PendienteT" & CStr(intTalla), VbGet) > 0 Then
                    'Cuidado, hay que buscar en los albaranes existentes para asegurarnos que este pedido no se referencie en ninguno de ellos.
                    flgPedidoYaUtilizado = False
                    For j = 1 To mobjAlbaranVentaItems.Count
                        If mobjAlbaranVentaItems(j).PedidoVentaItemID = lngPedidoVentaItemID Then
                            flgPedidoYaUtilizado = True
                            Exit For
                        End If
                    Next
                    'Bingo!
                    If Not flgPedidoYaUtilizado Then
                        Set objAlbaranVentaItem = Me.AlbaranVentaItems.Add
                        objAlbaranVentaItem.BeginEdit
                        objAlbaranVentaItem.TemporadaID = mudtprops.TemporadaID
                        objAlbaranVentaItem.CodigoArticuloColor = strArticuloColor
                        objAlbaranVentaItem.PedidoVentaItemID = lngPedidoVentaItemID
                        objAlbaranVentaItem.Descripcion = objPedidoVentaItem.ArticuloColor
                        objAlbaranVentaItem.PrecioVenta = objPedidoVentaItem.PrecioVenta
                        objAlbaranVentaItem.Descuento = objPedidoVentaItem.Descuento
                        objAlbaranVentaItem.Bruto = objPedidoVentaItem.Bruto
                        objAlbaranVentaItem.Comision = objPedidoVentaItem.Comision
                        objAlbaranVentaItem.ApplyEdit
                        flgPedidoNuevo = True
                        Set objPedidoVentaItem = Nothing
                        Exit For
                    End If
                End If
                
                Set objPedidoVentaItem = Nothing
            End If
        Next
    End If
    
    flgAlbaranExistente = False
    If Not flgPedidoExistente And _
       Not flgPedidoNuevo Then
        For i = 1 To mobjAlbaranVentaItems.Count
            If mobjAlbaranVentaItems(i).CodigoArticuloColor = strArticuloColor And _
               (Not mobjAlbaranVentaItems(i).HayPedido) And (Not mobjAlbaranVentaItems(i).IsDeleted) Then
                Set objAlbaranVentaItem = mobjAlbaranVentaItems(i)
                flgAlbaranExistente = True
                Exit For
            End If
        Next
    End If
    
    flgAlbaranNuevo = False
    If Not flgPedidoExistente And _
       Not flgPedidoNuevo And _
       Not flgAlbaranExistente Then
       flgAlbaranNuevo = True
       Set objAlbaranVentaItem = Me.AlbaranVentaItems.Add
    End If
    
    With objAlbaranVentaItem
        .BeginEdit
        If flgAlbaranNuevo Then
        ' Si es nuevo establecemos las propiedades de ser nuevo
            .TemporadaID = mudtprops.TemporadaID
            .ArticuloColorID = ArticuloColorID
            '.CodigoArticuloColor = strArticuloColor
            .Descripcion = strNombreArticuloColor
            .AlmacenID = mudtprops.AlmacenID
        End If
       
        ' Incrementamos la cantidad actual por talla
        dblCantidadActual = CallByName(objAlbaranVentaItem, "CantidadT" & CStr(intTalla), VbGet)
        'Miramos si es devolución para restar en vez de sumar
        If EsDevolucion Then
            dblCantidadFinal = dblCantidadActual - 1
        Else
            dblCantidadFinal = dblCantidadActual + 1
        End If
        CallByName objAlbaranVentaItem, "CantidadT" & CStr(intTalla), VbLet, dblCantidadFinal

        .ApplyEdit
        If .PedidoVentaItemID = 0 And Not bolPermitirSinPedido Then
            ' Emitimos una incidencia porque el albarán no está respaldado por un pedido.
            strResultadoFuncion = "Aviso, el artículo " & strArticuloColor & ", talla " & intTalla & " no se corresponde con ningún pedido pendiente."
        End If

    End With
    
    AlbaranItemCodigoBarras = strResultadoFuncion
    Set objAlbaranVentaItem = Nothing
End Function

' A partir de las capturas del TPV crear los items del albaran.
' Versión más simplificada de AlbaranItemCodigoBarras
Public Function AlbaranItemTPV(ByVal strCodigo As String) As String
    Dim intTalla As Integer
    Dim lngArticuloColorID As Long
    Dim intOrdenTalla As Integer
    Dim objArticuloColor As ArticuloColor
    Dim strArticuloColor As String
    Dim strNombreArticuloColor As String
    Dim objPedidosPendientes As PickerItems
    
    ' Validar que la información de talla sea correcta
    intTalla = CInt(Left(strCodigo, 2))
    If intTalla Mod 2 <> 0 Then
        Err.Raise vbObjectError + 1001, "AlbaranItemTPV", "Talla errónea! (" & intTalla & "), en el código " & strCodigo
        Exit Function
    End If
    If intTalla > 56 Or intTalla < 36 Then
        Err.Raise vbObjectError + 1001, "AlbaranItemTPV", "Talla errónea! (" & intTalla & "), en el código " & strCodigo
        Exit Function
    End If
        
    intOrdenTalla = (intTalla - 36) / 2
        
    ' Cargar el artículo, etc
    Set objArticuloColor = New ArticuloColor
    lngArticuloColorID = CLng(Right(strCodigo, Len(strCodigo) - 2))
    objArticuloColor.Load lngArticuloColorID, "EUR"
    strArticuloColor = objArticuloColor.NombreCompleto
    strNombreArticuloColor = strArticuloColor '& "-" & objArticuloColor.ObjArticulo.NombreCompleto &   '.Articulo & " " & objArticuloColor.NombreColor
    Set objArticuloColor = Nothing
    
    
    Set objPedidosPendientes = New PickerItems
    
    AlbaranItemTPV = AlbaranItemCodigoBarras(strArticuloColor, intTalla, objPedidosPendientes, lngArticuloColorID, True, strNombreArticuloColor)
    
    Set objPedidosPendientes = Nothing
    
End Function

' A partir de un ticket dado de alta por TPV se genera la factura correspondiente.
Public Sub AltaTicketTPV()
    Dim colAlbaranes As Collection
    Dim lngFacturaVentaIDA As Long
    Dim lngFacturaVentaIDB As Long
    Dim objAlbaranVentaFacturar As AlbaranVenta

    Me.ApplyEdit
    
    Set objAlbaranVentaFacturar = New AlbaranVenta
    Set colAlbaranes = New Collection
    colAlbaranes.Add Me.AlbaranVentaID
    objAlbaranVentaFacturar.FacturarAlbaranAB colAlbaranes, lngFacturaVentaIDA, lngFacturaVentaIDB, 100
    Set objAlbaranVentaFacturar = Nothing

End Sub

Public Sub FacturarAlbaranAB(ByRef colListaAlbaranes As Collection, ByRef lngFacturaVentaIDA As Long, ByRef lngFacturaVentaIDB As Long, ByVal PorcFacturacionAB As Integer)
Dim objFacturaVentaA As FacturaVenta
Dim objFacturaVentaB As FacturaVenta
Dim objEmpresaA As Empresa
Dim objEmpresaB As Empresa
Dim flgHayLineasFacturaA As Boolean     ' Si hay algo metido en la factura A
Dim flgHayLineasFacturaB As Boolean     ' Si hay algo metido en la factura B
Dim objAlbaranVentaItem As AlbaranVentaItem
Dim dblImporteLimiteA As Double
Dim dblImporteFacturaA As Double
Dim dblImporteFacturaB As Double
Dim intPorcFacturacionAB As Integer
'Dim bolGenerarFacturaA As Boolean
'Dim bolGenerarFacturaB As Boolean
Dim i As Integer
Dim j As Integer
Dim lngClienteID As Long
Dim lngAlbaranID As Long
Dim objAlbaranVenta As AlbaranVenta
Dim colAlbaranes As Collection
Dim dblTotalAFacturar As Double
Dim objPrimerAlbaranVenta As AlbaranVenta
Dim colAlbaranesPersist As Collection
Dim strObservaciones As String
Dim dblTotalImportesNegativos As Double
Dim dblTotalImportesPositivos As Double
Dim dblImporteMargen As Double


    
    ' Dar error si el albaran se está editando, no es nuevo o se ha borrado.
    If mflgEditing Then Err.Raise 383
    If Not mflgNew Then Err.Raise 383
    If mflgDeleted Then Err.Raise 383
    
    Set colAlbaranes = New Collection
    
    lngClienteID = 0
    dblTotalAFacturar = 0
    dblTotalImportesNegativos = 0
    dblTotalImportesPositivos = 0
    
    If colListaAlbaranes.Count = 0 Then Exit Sub
    
    For i = 1 To colListaAlbaranes.Count
        lngAlbaranID = CLng(colListaAlbaranes(i))
        Set objAlbaranVenta = New AlbaranVenta
        objAlbaranVenta.Load lngAlbaranID
        If lngClienteID = 0 Then lngClienteID = objAlbaranVenta.ClienteID
        If lngClienteID <> objAlbaranVenta.ClienteID Then _
            Err.Raise vbObjectError + 1001, "Sólo se pueden facturar juntos albaranes de un mismo cliente:" & objAlbaranVenta.Cliente
            
        dblTotalAFacturar = dblTotalAFacturar + objAlbaranVenta.AlbaranVentaItems.ImportePendienteFacturar
        For j = 1 To objAlbaranVenta.AlbaranVentaItems.Count
            If objAlbaranVenta.AlbaranVentaItems(j).Bruto >= 0 Then
                dblTotalImportesPositivos = dblTotalImportesPositivos + objAlbaranVenta.AlbaranVentaItems(j).Bruto
            Else
                dblTotalImportesNegativos = dblTotalImportesNegativos + Abs(objAlbaranVenta.AlbaranVentaItems(j).Bruto)
            End If
        Next
        colAlbaranes.Add objAlbaranVenta
        Set objAlbaranVenta = Nothing
    Next
    
    Set objPrimerAlbaranVenta = colAlbaranes.Item(1)
    
            
    '!!!!!!! OJO, LEER LA EMPRESA A Y LA B A PARTIR DE PARAMETROS, NO "A PELO".
    Set objEmpresaA = New Empresa
    Set objEmpresaB = New Empresa
    
    objEmpresaA.Load 2
    objEmpresaB.Load 4
        
    flgHayLineasFacturaA = False
    flgHayLineasFacturaB = False
    
    Set objFacturaVentaA = New FacturaVenta
    Set objFacturaVentaB = New FacturaVenta
    
    ' Establecemos los datos de la facturaA
    With objFacturaVentaA
        .BeginEdit
        .EmpresaID = objEmpresaA.EmpresaID
        .TemporadaID = objPrimerAlbaranVenta.TemporadaID
        .Portes = 0
        .Embalajes = 0
        .Numero = objEmpresaA.IncrementaFacturaVentas
        ' 3/6/2005 Por petición de Iñigo ponemos la fecha predeterminada como la del albaran
        .Fecha = objPrimerAlbaranVenta.Fecha  'Date
        .NuestraReferencia = vbNullString
        .SuReferencia = vbNullString
        '.Observaciones = "Facturación albarán " & .Numero & " de fecha " & Me.Fecha
        .Bultos = 0
        .PesoNeto = 0
        .PesoBruto = 0
        .Cliente = objPrimerAlbaranVenta.Cliente
        .Representante = objPrimerAlbaranVenta.Representante
        .Transportista = objPrimerAlbaranVenta.Transportista
        .FormaPago = objPrimerAlbaranVenta.FormaPago
        .DatoComercial.ChildBeginEdit
        ' El descuento lo debe coger de las condiciones del cliente, no del albarán (Isabel: 20/8/04)
                    '.DatoComercial.Descuento = objPrimerAlbaranVenta.DatoComercial.Descuento
        .DatoComercial.IVA = objPrimerAlbaranVenta.DatoComercial.IVA
        .DatoComercial.RecargoEquivalencia = objPrimerAlbaranVenta.DatoComercial.RecargoEquivalencia
        .DatoComercial.ChildApplyEdit
        'DAtos de TPV
        .CentroGestionID = objPrimerAlbaranVenta.CentroGestionID
        .AlmacenID = objPrimerAlbaranVenta.AlmacenID
        .TerminalID = objPrimerAlbaranVenta.TerminalID
    End With
    ' Establecemos los datos de la facturaB
    With objFacturaVentaB
        .BeginEdit
        .EmpresaID = objEmpresaB.EmpresaID
        .TemporadaID = objPrimerAlbaranVenta.TemporadaID
        .Portes = 0
        .Embalajes = 0
        .Numero = objEmpresaB.IncrementaFacturaVentas
        ' 3/6/2005 Por petición de Iñigo ponemos la fecha predeterminada como la del albaran
        .Fecha = objPrimerAlbaranVenta.Fecha  'Date
        .NuestraReferencia = vbNullString
        .SuReferencia = vbNullString
        '.Observaciones = "Facturación albarán " & .Numero & " de fecha " & Me.Fecha
        .Bultos = 0
        .PesoNeto = 0
        .PesoBruto = 0
        .Cliente = objPrimerAlbaranVenta.Cliente
        .Representante = objPrimerAlbaranVenta.Representante
        .Transportista = objPrimerAlbaranVenta.Transportista
        ' No alimento los datos de forma de pago ni de descuento ni IVA ni RE del dato comercial porque se están
        ' estableciendo cuando se informa del cliente (sabe la empresa que es y lo pone bien).
        '        .DatoComercial.ChildBeginEdit
        '        .DatoComercial.Descuento = objPrimerAlbaranVenta.DatoComercial.Descuento
        '        .DatoComercial.ChildApplyEdit
        'DAtos de TPV
        .CentroGestionID = objPrimerAlbaranVenta.CentroGestionID
        .AlmacenID = objPrimerAlbaranVenta.AlmacenID
        .TerminalID = objPrimerAlbaranVenta.TerminalID
    End With
        
    
    intPorcFacturacionAB = PorcFacturacionAB
    
    ' Calculo el importe limite
    Select Case intPorcFacturacionAB
    Case 100
        dblImporteLimiteA = dblTotalAFacturar
    Case 0
        dblImporteLimiteA = 0
    Case Else
        dblImporteLimiteA = IIf(intPorcFacturacionAB <> 0, dblTotalAFacturar * intPorcFacturacionAB / 100, 0)
    End Select
    
    If dblImporteLimiteA >= 0 Then
        dblImporteMargen = dblTotalImportesNegativos
    Else
        dblImporteMargen = dblTotalImportesPositivos
    End If
    
    strObservaciones = "Facturación albarán: "
    'Set colAlbaranesPersist = New Collection
    Dim objFacturaVentaItem As FacturaVentaItem
    For Each objAlbaranVenta In colAlbaranes
        If Not objAlbaranVenta.AlbaranVentaItems.Facturado Then
            strObservaciones = strObservaciones & objAlbaranVenta.Numero & " "
        End If
        ' Ahora recorro todas las lineas de albarán y las asigno a cada factura.
        For Each objAlbaranVentaItem In objAlbaranVenta.AlbaranVentaItems
            ' compruebo que la linea no esté facturada!!
            If Not objAlbaranVentaItem.Facturado Then
                ' Si no supera el limite, se asigna a la factura A.
                ' Si no va a superar el límite B, se asigna a la B.
                ' En caso contrario se reasigna segun la cantidad de prendas.
                If Abs(dblImporteFacturaA + objAlbaranVentaItem.Bruto) <= Abs(dblImporteLimiteA) + dblImporteMargen Then
                    Set objFacturaVentaItem = objFacturaVentaA.FacturaVentaItems.Add
                    objFacturaVentaItem.BeginEdit
                    objFacturaVentaItem.FacturaDesdeAlbaran objAlbaranVentaItem.AlbaranVentaItemID
                    objFacturaVentaItem.ApplyEdit
                    Set objFacturaVentaItem = Nothing
                    dblImporteFacturaA = dblImporteFacturaA + objAlbaranVentaItem.Bruto
                    flgHayLineasFacturaA = True
                ElseIf Abs(dblImporteFacturaB + objAlbaranVentaItem.Bruto) <= Abs(dblTotalAFacturar - dblImporteLimiteA) Then
                    Set objFacturaVentaItem = objFacturaVentaB.FacturaVentaItems.Add
                    objFacturaVentaItem.BeginEdit
                    objFacturaVentaItem.FacturaDesdeAlbaran objAlbaranVentaItem.AlbaranVentaItemID
                    objFacturaVentaItem.ApplyEdit
                    Set objFacturaVentaItem = Nothing
                    dblImporteFacturaB = dblImporteFacturaB + objAlbaranVentaItem.Bruto
                    flgHayLineasFacturaB = True
                Else
                    Dim intCantidadA As Integer
                    intCantidadA = RepartirA(dblImporteLimiteA, objAlbaranVentaItem.PrecioVenta, dblImporteFacturaA)
                    If intCantidadA <> 0 Then
                        ' creamos la parte de la linea de la factura A
                        Set objFacturaVentaItem = objFacturaVentaA.FacturaVentaItems.Add
                        objFacturaVentaItem.BeginEdit
                        objFacturaVentaItem.TemporadaID = objAlbaranVentaItem.TemporadaID
                        If objAlbaranVentaItem.ArticuloColorID <> 0 Then
                            objFacturaVentaItem.CodigoArticuloColor = objAlbaranVentaItem.CodigoArticuloColor
                        Else
                            objFacturaVentaItem.ArticuloColorID = 0
                        End If
                        objFacturaVentaItem.Descripcion = objAlbaranVentaItem.Descripcion
                        objFacturaVentaItem.Cantidad = intCantidadA
                        objFacturaVentaItem.AlbaranVentaItemID = objAlbaranVentaItem.AlbaranVentaItemID
                        objFacturaVentaItem.PrecioVenta = objAlbaranVentaItem.PrecioVenta
                        objFacturaVentaItem.Descuento = objAlbaranVentaItem.Descuento
                        objFacturaVentaItem.ApplyEdit
                        dblImporteFacturaA = dblImporteFacturaA + objFacturaVentaItem.Bruto
                        Set objFacturaVentaItem = Nothing
                        flgHayLineasFacturaA = True
                    End If
                    
                    If objAlbaranVentaItem.Cantidad - intCantidadA <> 0 Then
                        ' creamos la parte de la linea de la factura B
                        Set objFacturaVentaItem = objFacturaVentaB.FacturaVentaItems.Add
                        objFacturaVentaItem.BeginEdit
                        objFacturaVentaItem.TemporadaID = objAlbaranVentaItem.TemporadaID
                        If objAlbaranVentaItem.ArticuloColorID <> 0 Then
                            objFacturaVentaItem.CodigoArticuloColor = objAlbaranVentaItem.CodigoArticuloColor
                        Else
                            objFacturaVentaItem.ArticuloColorID = 0
                        End If
                        objFacturaVentaItem.Descripcion = objAlbaranVentaItem.Descripcion
                        objFacturaVentaItem.Cantidad = objAlbaranVentaItem.Cantidad - intCantidadA
                        objFacturaVentaItem.AlbaranVentaItemID = objAlbaranVentaItem.AlbaranVentaItemID
                        objFacturaVentaItem.PrecioVenta = objAlbaranVentaItem.PrecioVenta
                        objFacturaVentaItem.Descuento = objAlbaranVentaItem.Descuento
                        objFacturaVentaItem.ApplyEdit
                        dblImporteFacturaB = dblImporteFacturaB + objFacturaVentaItem.Bruto
                        Set objFacturaVentaItem = Nothing
                        flgHayLineasFacturaB = True
                    End If
                End If
                objAlbaranVentaItem.FacturarAB
            End If
        Next
    Next
    Set colAlbaranes = Nothing
    
    objFacturaVentaA.Observaciones = strObservaciones
    objFacturaVentaB.Observaciones = strObservaciones
    
    objFacturaVentaA.CalcularBruto
    objFacturaVentaA.CrearCobros
    objFacturaVentaA.CambiarComision
    
    objFacturaVentaB.CalcularBruto
    objFacturaVentaB.CrearCobros
    objFacturaVentaB.CambiarComision
    
    Dim objPersist As AlbaranVentaPersist
    Dim strResul As String
    Set objPersist = CreateObject("GCServerMTS.AlbaranVentaPersist", PERSIST_SERVER)
    strResul = objPersist.SaveAlbaranesAB(objFacturaVentaA.GetState, objFacturaVentaA.FacturaVentaItems.GetState, objFacturaVentaA.CobrosPagos.GetState, flgHayLineasFacturaA, _
                               objFacturaVentaB.GetState, objFacturaVentaB.FacturaVentaItems.GetState, objFacturaVentaB.CobrosPagos.GetState, flgHayLineasFacturaB, _
                               colAlbaranesPersist)
    Set objPersist = Nothing
    
    If flgHayLineasFacturaA Then
        objEmpresaA.EstableceFacturaVentas objFacturaVentaA.Numero
    Else
        objEmpresaA.DecrementaFacturaVentas
    End If
    If flgHayLineasFacturaB Then
        objEmpresaB.EstableceFacturaVentas objFacturaVentaB.Numero
    Else
        objEmpresaB.DecrementaFacturaVentas
    End If
    
    Set objEmpresaA = Nothing
    Set objEmpresaB = Nothing
    
    Set objFacturaVentaA = Nothing
    Set objFacturaVentaB = Nothing
    
    Dim strResulArr() As String
    strResulArr = Split(strResul, "|", 2)
    lngFacturaVentaIDA = CLng(strResulArr(0))
    lngFacturaVentaIDB = CLng(strResulArr(1))
End Sub

Private Function RepartirA(ImporteLimite As Double, Precio As Double, ImporteA As Double) As Integer
Dim intRepartirA As Integer

    If Precio <> 0 Then
        intRepartirA = (ImporteLimite - ImporteA) \ Precio
        If ImporteLimite - ImporteA Mod Precio = 0 Then
            RepartirA = intRepartirA
        Else
            RepartirA = intRepartirA + 1
        End If
    Else
        RepartirA = 0
    End If
End Function

Public Sub AsignarClientePredeterminado()
    Dim objParametroAplicacion As ParametroAplicacion
    Dim lngClienteID As Long
    
    Set objParametroAplicacion = New ParametroAplicacion
    lngClienteID = objParametroAplicacion.ClientePredeterminado
    Me.Cliente = mobjClientes.Item(lngClienteID)
    Set objParametroAplicacion = Nothing

End Sub

Public Function ImpresionTicket() As String
    Dim strResultado As String
    Dim LineaCuerpo As Integer
    Dim objAlbaranVentaItem As AlbaranVentaItem
    Dim txtTalla As String
    Dim objCentroGestion As CentroGestion
    Dim objDireccion As Direccion
    Dim strNombreCentroGestion As String
    Dim objEmpresa As Empresa
    Dim strDNINIF As String
    
    
    Set objCentroGestion = New CentroGestion
    objCentroGestion.Load mudtprops.CentroGestionID
    
    Set objDireccion = objCentroGestion.Direccion
    strNombreCentroGestion = objCentroGestion.Nombre
    
    'Cargamos la empresa asociada
    Set objEmpresa = New Empresa
    objEmpresa.Load objCentroGestion.EmpresaID
    strDNINIF = objEmpresa.DNINIF
    Set objEmpresa = Nothing
    
    Set objCentroGestion = Nothing
    
    ' Cabecera
    'Datos de Empresa
    strResultado = vbNullString
'    strResultado = strResultado & Chr$(cnESCP2Escape) & Chr$(cnESCP2InicializePrinter)   'Inicializar la impresora
'    strResultado = strResultado & Chr$(cnESCP2Escape) & Chr$(cnESCP2PrintMode) & Chr$(cnESCP2DoubleHeight)
    strResultado = strResultado & strNombreCentroGestion
    strResultado = strResultado & vbCrLf
'    strResultado = strResultado & Chr$(cnESCP2Escape) & Chr$(cnESCP2InicializePrinter)   'Inicializar la impresora
    strResultado = strResultado & "Telf.: " & objDireccion.Telefono1
    strResultado = strResultado & vbCrLf
    strResultado = strResultado & objDireccion.Calle
    strResultado = strResultado & vbCrLf
    strResultado = strResultado & objDireccion.CodigoPostal & "-" & objDireccion.Poblacion & "  (" & objDireccion.Provincia & ")"
    strResultado = strResultado & vbCrLf
    strResultado = strResultado & "C.I.F.: " & strDNINIF
    strResultado = strResultado & vbCrLf
    
    ' Fecha y numero
    strResultado = strResultado & Format(Me.Fecha, "dd/mm/yyyy HH:MM") & " NUM.:" & Format(Me.Numero, "###,###,###")
    strResultado = strResultado & vbCrLf
    strResultado = strResultado & cnRayaSeparador
    strResultado = strResultado & vbCrLf
    
    
    LineaCuerpo = 1
    For Each objAlbaranVentaItem In mobjAlbaranVentaItems
        
        Select Case True
        Case objAlbaranVentaItem.CantidadT36 > 0
            txtTalla = "36"
        Case objAlbaranVentaItem.CantidadT38 > 0
            txtTalla = "38"
        Case objAlbaranVentaItem.CantidadT40 > 0
            txtTalla = "40"
        Case objAlbaranVentaItem.CantidadT42 > 0
            txtTalla = "42"
        Case objAlbaranVentaItem.CantidadT44 > 0
            txtTalla = "44"
        Case objAlbaranVentaItem.CantidadT46 > 0
            txtTalla = "46"
        Case objAlbaranVentaItem.CantidadT48 > 0
            txtTalla = "48"
        Case objAlbaranVentaItem.CantidadT50 > 0
            txtTalla = "50"
        Case objAlbaranVentaItem.CantidadT52 > 0
            txtTalla = "52"
        Case objAlbaranVentaItem.CantidadT54 > 0
            txtTalla = "54"
        Case objAlbaranVentaItem.CantidadT56 > 0
            txtTalla = "56"
        Case Else
            'ojoojo: devolver error
        End Select
        
        strResultado = strResultado & txtTalla & objAlbaranVentaItem.ArticuloColorID & " " & Left(objAlbaranVentaItem.Descripcion, cnEspacioAreaTexto) & Right(FormatoMoneda(objAlbaranVentaItem.PrecioVenta * objAlbaranVentaItem.Cantidad, "EUR", False), 9)
        strResultado = strResultado & vbCrLf
        LineaCuerpo = LineaCuerpo + 1
    
        If objAlbaranVentaItem.Descuento > 0 Then
            strResultado = strResultado & "(con el " & objAlbaranVentaItem.Descuento & "% de dcto.)          " & FormatoMoneda(objAlbaranVentaItem.Bruto, "EUR", False)
            strResultado = strResultado & vbCrLf
            LineaCuerpo = LineaCuerpo + 1
        End If
    
    Next
    
    strResultado = strResultado & cnRayaSeparador
    strResultado = strResultado & vbCrLf
 '   strResultado = strResultado & Chr$(cnESCP2Escape) & Chr$(cnESCP2PrintMode) & Chr$(cnESCP2DoubleHeight)
    strResultado = strResultado & "    T O T A L...:      " & FormatoMoneda(Me.TotalBruto, "EUR", False)
    strResultado = strResultado & vbCrLf
'    strResultado = strResultado & Chr$(cnESCP2Escape) & Chr$(cnESCP2InicializePrinter)   'Inicializar la impresora
    strResultado = strResultado & " ENTREGADO......:      "  ' OJOOJO PONER LA CANTIDAD ENTREGADA & FormatoMoneda(mobjAlbaranVenta.TotalBruto, "EUR")
    strResultado = strResultado & vbCrLf
    strResultado = strResultado & " CAMBIO.........:      "  ' OJOOJO PONER EL CAMBIO & FormatoMoneda(mobjAlbaranVenta.TotalBruto, "EUR")
    strResultado = strResultado & vbCrLf
    strResultado = strResultado & "  P.V.P    I.V.A.  INCLUIDO"
    strResultado = strResultado & " " & vbCrLf
    strResultado = strResultado & " " & vbCrLf
'    strResultado = strResultado & Chr$(cnESCP2Escape) & Chr$(cnESCP2ColorSelect) & Chr$(cnESCP2RedColor)  'Poner el ticket en rojo
    strResultado = strResultado & "  GRACIAS POR SU VISITA"
    strResultado = strResultado & " " & vbCrLf
'    strResultado = strResultado & Chr$(cnESCP2Escape) & Chr$(cnESCP2InicializePrinter)   'Inicializar la impresora
    strResultado = strResultado & " " & vbCrLf
    strResultado = strResultado & " " & vbCrLf
    strResultado = strResultado & " " & vbCrLf
    strResultado = strResultado & " " & vbCrLf
    strResultado = strResultado & " " & vbCrLf
    strResultado = strResultado & " " & vbCrLf
    strResultado = strResultado & " " & vbCrLf
    strResultado = strResultado & " " & vbCrLf
    strResultado = strResultado & " " & vbCrLf
    
    ImpresionTicket = strResultado
    

End Function

