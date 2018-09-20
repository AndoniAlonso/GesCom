VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PedidoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As PedidoVentaProps
Private mudtSave As PedidoVentaProps

Private mMoneda As String

' Aqui declaro la comision (depende del representante) a aplicar a las lineas de pedido
' nuevas
Private mComision As Double

Private mobjPedidoVentaItems As PedidoVentaItems

' Aqui declaro el objeto 'child' dato comercial
Private mobjDatoComercial As DatoComercial

' Aqui declaro los ComboBox de la lista de clientes, representantes, transportistas y formas de pago.
Private mobjClientes As TextList
Private mobjRepresentantes As TextList
Private mobjTransportistas As TextList
Private mobjFormasPago As TextList

' Aquí declaro el objeto empresa para ejecutar las reglas de forma de pago y descuento para empresas con anulación de IVA.
Private mobjEmpresa As Empresa

Event Valid(IsValid As Boolean)

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit(Moneda As String)

    If mflgEditing Then Err.Raise 445
  
    ' save object state
    LSet mudtSave = mudtprops
  
    mMoneda = Moneda
    mobjPedidoVentaItems.BeginEdit Moneda
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
    mobjPedidoVentaItems.CancelEdit
    mobjDatoComercial.CancelEdit
  
End Sub

Public Sub ApplyEdit()
    Dim objPersist As PedidoVentaPersist
    Dim strItems As String
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.PedidoVentaPersist", PERSIST_SERVER)

    TotalBruto = TotalBruto
    
    If mflgDeleted And Not mflgNew Then
        mobjPedidoVentaItems.Actualizar
        
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.PedidoVentaID, mudtprops.DatoComercialID, mobjPedidoVentaItems.GetState
    
        mobjDatoComercial.Delete
        mobjDatoComercial.ApplyEdit
    
        mflgNew = True
        mflgDeleted = False
    
    ElseIf IsDirty Or mobjPedidoVentaItems.IsDirty Or mflgNew Then
            If Not IsValid Then Err.Raise 445
            mudtprops.DatoComercial.Buffer = mobjDatoComercial.GetState
            
            mobjPedidoVentaItems.Actualizar
            
            strItems = mobjPedidoVentaItems.GetState
            SetState objPersist.Save(GetState, strItems)
            mobjPedidoVentaItems.SetState (strItems)
            
            mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
            mobjDatoComercial.ApplyEdit
            mudtprops.DatoComercialID = mobjDatoComercial.DatoComercialID
            
            ' save object state
            LSet mudtSave = mudtprops
            mflgNew = False
    Else
        mobjDatoComercial.ApplyEdit
    
    End If

    mobjPedidoVentaItems.ApplyEdit mComision
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
    mudtprops.FechaTopeServicio = Date
    mudtprops.Numero = 0
    mudtprops.Observaciones = vbNullString
    mudtprops.TotalBrutoEUR = 0
    mudtprops.TotalBrutoPTA = 0
      
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
      
    mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "EmpresaID", True
    mobjValid.RuleBroken "ClienteID", True
    mobjValid.RuleBroken "RepresentanteID", True
    mobjValid.RuleBroken "TransportistaID", True
    mobjValid.RuleBroken "FormaPagoID", True
    
    Set mobjPedidoVentaItems = New PedidoVentaItems
    
    mComision = 0
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjDatoComercial = Nothing
    Set mobjClientes = Nothing
    Set mobjRepresentantes = Nothing
    Set mobjTransportistas = Nothing
    Set mobjFormasPago = Nothing
    Set mobjPedidoVentaItems = Nothing
    
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

Public Sub Load(PedidoVentaID As Long, Moneda As String, Optional Numero As Long, Optional EmpresaID As Long)
    Dim objPersist As PedidoVentaPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.PedidoVentaPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(PedidoVentaID, Numero, EmpresaID)
    Set objPersist = Nothing

    mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
  
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "EmpresaID", False
    mobjValid.RuleBroken "ClienteID", False
    mobjValid.RuleBroken "RepresentanteID", False
    mobjValid.RuleBroken "TransportistaID", False
    mobjValid.RuleBroken "FormaPagoID", False
  
    mMoneda = Moneda
    mobjPedidoVentaItems.Load PedidoVentaID, Moneda
    
    LeerComision
    
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
      
    If mobjPedidoVentaItems.Servido Then _
       Err.Raise vbObjectError + 1001, "El pedido tiene entregas realizadas, no se puede eliminar."

    If mobjPedidoVentaItems.OrdenCorte Then _
       Err.Raise vbObjectError + 1001, "El pedido tiene ordenes de corte, no se puede eliminar."

    mflgDeleted = True
    mflgDirty = True

    mobjPedidoVentaItems.Delete
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

Public Property Get PedidoVentaID() As Long

    PedidoVentaID = mudtprops.PedidoVentaID
  
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

Public Property Get Moneda() As String

    Moneda = Trim(mMoneda)

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
    If mobjPedidoVentaItems.Servido Then _
        Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, ya se han realizado entregas"
     
    ' Devolvemos error si ya ha habido ordenes de corte y queremos cambiar de cliente.
    If mobjPedidoVentaItems.OrdenCorte Then _
        Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, ya hay ordenes de corte"
     
    ' Cambiamos de cliente
    Set objCliente = New Cliente
      
    objCliente.Load mobjClientes.Key(Value)
  
    mudtprops.RepresentanteID = objCliente.RepresentanteID
    mudtprops.TransportistaID = objCliente.TransportistaID
    
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

Public Property Let FechaTopeServicio(Value As Variant)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) = 0 Then
        mudtprops.FechaTopeServicio = vbNullString
        mobjValid.RuleBroken "FechaTopeServicio", True

    ElseIf IsDate(Value) Then
        mudtprops.FechaTopeServicio = CDate(Value)
        mobjValid.RuleBroken "FechaTopeServicio", False

    ElseIf IsNull(Value) Then
        mudtprops.FechaTopeServicio = CDate("00:00:00")
        mobjValid.RuleBroken "FechaTopeServicio", False

    Else
        mobjValid.RuleBroken "FechaTopeServicio", True

    End If

    mflgDirty = True

End Property

Public Property Get FechaTopeServicio() As Variant

    With mudtprops
        FechaTopeServicio = IIf(.FechaTopeServicio = 0, vbNullString, .FechaTopeServicio)
    End With

End Property

Public Property Let FechaEntrega(Value As Variant)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) = 0 Then
        mudtprops.FechaEntrega = vbNullString
        mobjValid.RuleBroken "FechaEntrega", True
    
    ElseIf IsDate(Value) Then
        mudtprops.FechaEntrega = CDate(Value)
        mobjValid.RuleBroken "FechaEntrega", False
    
    ElseIf IsNull(Value) Then
        mudtprops.FechaTopeServicio = CDate("00:00:00")
        mobjValid.RuleBroken "FechaTopeServicio", False

    Else
        mobjValid.RuleBroken "FechaTopeServicio", True

    End If

    mflgDirty = True

End Property

Public Property Get FechaEntrega() As Variant

    With mudtprops
        FechaEntrega = IIf(.FechaEntrega = 0, vbNullString, .FechaEntrega)
    End With

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

Private Property Let TotalBruto(Value As Double)

    If Not mflgEditing Then Err.Raise 383

    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"

    mobjValid.RuleBroken "TotalBruto", (Value < 0)
    If EsEUR(mMoneda) Then
        mudtprops.TotalBrutoEUR = Round(Value, 2)
        mudtprops.TotalBrutoPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.TotalBrutoPTA = Round(Value, 0)
        mudtprops.TotalBrutoEUR = PTA2EUR(Value)
    Else
        mudtprops.TotalBrutoEUR = 0
        mudtprops.TotalBrutoPTA = 0
    End If

    mflgDirty = True

End Property

Public Property Get TotalBruto() As Double

    TotalBruto = mobjPedidoVentaItems.Bruto
    
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
    mobjPedidoVentaItems.CambiarComision (mComision)
    
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

Public Property Get PedidoVentaItems() As PedidoVentaItems

    Set PedidoVentaItems = mobjPedidoVentaItems

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

Public Sub ActualizarPreciosVenta()

    If Not mflgEditing Then Err.Raise 445
  
    mobjPedidoVentaItems.ActualizarPreciosVenta
    
End Sub

Private Sub CargarDatosEmpresa()
    ' Cargamos los datos de la empresa para comprobar sus condiciones de IVA.

    Set mobjEmpresa = New Empresa
    mobjEmpresa.Load mudtprops.EmpresaID

End Sub

Private Function GetState() As String
    Dim udtData As PedidoVentaData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As PedidoVentaData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
