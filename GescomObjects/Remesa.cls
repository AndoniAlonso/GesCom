VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Remesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As RemesaProps
Private mudtSave As RemesaProps

Private mMoneda As String

Private mobjCobrosPagos As CobrosPagos

' Aqui declaro los ComboBox de la lista de Bancos
Private mobjBancos As TextList

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
    mobjCobrosPagos.BeginEdit Moneda
    
    mflgEditing = True
    
End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    ' restore object state
    LSet mudtprops = mudtSave
    mobjCobrosPagos.CancelEdit
    
End Sub

Public Sub ApplyEdit()
    Dim objPersist As RemesaPersist
    Dim strItems As String
  
    If Not mflgEditing Then Err.Raise 445
    
    Set objPersist = CreateObject("GCServerMTS.RemesaPersist", PERSIST_SERVER)
    
    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.Save mobjCobrosPagos.GetState
        
        mflgNew = True
        mflgDeleted = False

    ElseIf mflgDirty Or mobjCobrosPagos.IsDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
            
            strItems = mobjCobrosPagos.GetState
            'SetState
            objPersist.Save (strItems)
            mobjCobrosPagos.SetState (strItems)
             
            ' save object state
            LSet mudtSave = mudtprops
            mflgNew = False
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
    
    mudtprops.FechaDomiciliacion = Date
    mudtprops.BancoID = 0
    mudtprops.SituacionComercial = "S"
    mudtprops.NumeroEfectos = 0
    mudtprops.ImportePTA = 0
    mudtprops.ImporteEUR = 0
    
    Set mobjBancos = New TextList
    mobjBancos.Load "vBancosEmpresas"
    mudtprops.BancoID = mobjBancos.Key(mobjBancos.IndexItem(1))
    
    mobjValid.RuleBroken "BancoID", True
      
    Set mobjCobrosPagos = New CobrosPagos
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjBancos = Nothing
    Set mobjCobrosPagos = Nothing
    
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

Public Sub Load(BancoID As Long, FechaDomiciliacion As Date, Moneda As String)
    Dim objPersist As RemesaPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.RemesaPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(BancoID, FechaDomiciliacion)
    Set objPersist = Nothing
    
    mobjValid.RuleBroken "BancoID", False
    
    mMoneda = Moneda
    mobjCobrosPagos.LoadRemesa BancoID, FechaDomiciliacion, Moneda
    
    mflgNew = False
    
End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
    
    mflgDeleted = True
    mflgDirty = True
    
    mobjCobrosPagos.DeleteRemesa
    
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

'Public Property Get RemesaID() As Long
'
'    RemesaID = mudtProps.RemesaID
'
'End Property
'

Public Property Get Moneda() As String

    Moneda = Trim(mMoneda)

End Property

Public Property Let FechaDomiciliacion(Value As Variant)

    If Not mflgEditing Then Err.Raise 383
    
    If Len(Trim(Value)) = 0 Then
        mudtprops.FechaDomiciliacion = vbNullString
        mobjValid.RuleBroken "FechaDomiciliacion", True
    
    ElseIf IsDate(Value) Then
        mudtprops.FechaDomiciliacion = CDate(Value)
        mobjValid.RuleBroken "FechaDomiciliacion", False
    
    Else
        mobjValid.RuleBroken "FechaDomiciliacion", True
    
    End If
    
    mflgDirty = True
    
End Property

Public Property Get FechaDomiciliacion() As Variant

    With mudtprops
      FechaDomiciliacion = IIf(.FechaDomiciliacion = 0, vbNullString, .FechaDomiciliacion)
    End With
    
End Property

Public Property Get BancoID() As Long

    BancoID = mudtprops.BancoID
  
End Property

' Esta funcion lee el cliente para asignar sus propiedades por defecto
Public Property Let Banco(Value As String)

    If Not mflgEditing Then Err.Raise 383
    
    mobjValid.RuleBroken "BancoID", (mobjBancos.Key(Value) = 0)
      
    'Salimos si queremos asignar algo invalido
    If mobjBancos.Key(Value) = 0 Then Exit Property
       
    ' Devolvemos error si ya ha habido entregas y queremos cambiar de Banco.
    If mobjCobrosPagos.Remesado Then _
        Err.Raise vbObjectError + 1001, "No se puede cambiar el banco, ya se enviado la remesa"
       
    mudtprops.BancoID = mobjBancos.Key(Value)
    
    mflgDirty = True
    
End Property

Public Property Get Banco() As String

    Banco = mobjBancos.Item(mudtprops.BancoID)

End Property

Public Property Get Bancos() As TextList

    Set Bancos = mobjBancos

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
      Err.Raise vbObjectError + 1001, "Situacion Comercial de la remesa erronea"
    End Select
    
End Property

Public Property Get NumeroEfectos() As Long

    NumeroEfectos = mudtprops.NumeroEfectos
  
End Property

Public Property Let NumeroEfectos(Value As Long)

    If Not mflgEditing Then Err.Raise 383

    If Value <= 0 Then _
        Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"

    mobjValid.RuleBroken "NumeroEfectos", (Value <= 0)

    mudtprops.NumeroEfectos = Value

    mflgDirty = True

End Property

'Public Property Let Importe(Value As Double)
'
'    If Not mflgEditing Then Err.Raise 383
'
'    If Value < 0 Then _
'        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
'
'    mobjValid.RuleBroken "Importe", (Value < 0)
'
'    If EsEUR(mMoneda) Then
'        mudtProps.ImporteEUR = Round(Value, 2)
'        mudtProps.ImportePTA = EUR2PTA(Value)
'    ElseIf EsPTA(mMoneda) Then
'        mudtProps.ImportePTA = Round(Value, 0)
'        mudtProps.ImporteEUR = PTA2EUR(Value)
'    Else
'        mudtProps.ImportePTA = 0
'        mudtProps.ImporteEUR = 0
'    End If
'
'    mflgDirty = True
'
'End Property
'
Public Property Get Importe() As Double

    If EsEUR(mMoneda) Then
        Importe = mudtprops.ImporteEUR
    ElseIf EsPTA(mMoneda) Then
        Importe = mudtprops.ImportePTA
    Else
        Importe = 0
    End If
    
End Property

Public Property Get CobrosPagos() As CobrosPagos

    Set CobrosPagos = mobjCobrosPagos

End Property

Public Sub SetDatosRemesa()

    If Not mflgEditing Then Err.Raise 445
    
    If Not IsValid Then Err.Raise 445
    
    mobjCobrosPagos.SetDatosRemesa mudtprops.BancoID, mudtprops.FechaDomiciliacion

End Sub

Public Sub MarcarRemesado()

    If Not mflgEditing Then Err.Raise 445
    
    If Not IsValid Then Err.Raise 445
    
    mobjCobrosPagos.MarcarRemesado
    
End Sub

Public Property Get Contabilizado() As Boolean
Dim objCobroPago As CobroPago

' Devolvemos contabilizado si ALGUNO de los cobros/pagos está contabilizado.
Contabilizado = False
For Each objCobroPago In mobjCobrosPagos
    If objCobroPago.Contabilizado Then
        Contabilizado = True
        Exit Property
    End If
Next

End Property

Public Sub Contabilizar(TemporadaID As Long, EmpresaID As Long, Optional ForzarContabilidad As Boolean = False)
    Dim objAsiento As Asiento
    Dim objCliente As Cliente
    Dim strDescripcion As String
    Dim objPersist As RemesaPersist
    'Dim strCuentaBase As String
    'Dim strTitular As String
    Dim objCobroPago As CobroPago
    Dim objFacturaVenta As FacturaVenta
    Dim objBanco As Banco
    Dim lngNumero As Long
    Dim dblImporteTotal As Double
    
    If mflgEditing Then Err.Raise 383
    
    ' Realizamos la comprobación de si ya está contabilizada, y damos error si no se ha
    ' marcada la contabilidad como forzada.
    If Me.Contabilizado And _
       Not ForzarContabilidad Then _
        Err.Raise vbObjectError + 1010, "La Remesa ya esta marcada como contabilizada"
     
    'mudtprops.SituacionContable = "C"
    
    ' Aquí por fin contabilizo
    Set objAsiento = New Asiento
    objAsiento.NuevoAsiento Me.BancoID, CStr(Year(Me.FechaDomiciliacion)), "Remesa " & Me.Banco & ", fecha " & CStr(Me.FechaDomiciliacion), _
                            TemporadaID, EmpresaID
                            
    Set objBanco = New Banco
    objBanco.Load Me.BancoID
    
    dblImporteTotal = 0
    For Each objCobroPago In mobjCobrosPagos
        Set objCliente = New Cliente
        objCliente.Load objCobroPago.GetClienteID
        If objCobroPago.HayFactura Then
            Set objFacturaVenta = New FacturaVenta
            objFacturaVenta.Load objCobroPago.FacturaID
            lngNumero = objFacturaVenta.Numero
        Else
            lngNumero = 0
        End If
        
        With objAsiento
            If objCobroPago.Importe > 0 Then
                strDescripcion = "Cobro efecto remesado fra." & CStr(lngNumero)
            Else
                strDescripcion = "Devolucion efecto remesado fra." & CStr(lngNumero)
            End If
            
                                    
            .NuevoApunte objCliente.CuentaContable, "Haber", _
                                   objCobroPago.Importe, strDescripcion, _
                                   Me.FechaDomiciliacion, CStr(lngNumero)
                                   
            dblImporteTotal = dblImporteTotal + objCobroPago.Importe
        End With
        
        Set objCliente = Nothing
        
        If objCobroPago.HayFactura Then Set objFacturaVenta = Nothing
        
    Next
    
    strDescripcion = "Remesa " & Me.Banco & ", fecha " & CStr(Me.FechaDomiciliacion)
    objAsiento.NuevoApunte objBanco.CuentaContable, "Debe", _
                           dblImporteTotal, strDescripcion, _
                           Me.FechaDomiciliacion, objBanco.CuentaBancaria.NombreEntidad & " " & CStr(Me.FechaDomiciliacion)
                           
    objAsiento.CuadrarAsiento objBanco.CuentaContable, Me.FechaDomiciliacion, CStr(lngNumero)
    
    
    Set objPersist = CreateObject("GCServerMTS.RemesaPersist", PERSIST_SERVER)
    objPersist.Contabilizar Me.BancoID, Me.FechaDomiciliacion, _
                            "C", _
                            objAsiento.GetSuperState
    Set objPersist = Nothing

    Set objAsiento = Nothing
    Set objBanco = Nothing
    
End Sub

Private Function GetState() As String
    Dim udtData As RemesaData
    
    With mudtprops
      .IsDeleted = mflgDeleted
      .IsNew = mflgNew
      .IsDirty = mflgDirty
    End With
    
    LSet udtData = mudtprops
    GetState = udtData.Buffer
    
End Function

Private Sub SetState(Buffer As String)
    
    Dim udtData As RemesaData
    
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub

