VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Asiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As AsientoProps
Private mudtSave As AsientoProps

Private mMoneda As String

Private mobjApuntes As Apuntes
Private mobjIVAApuntes As IVAApuntes

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
    mobjApuntes.BeginEdit Moneda
    mobjIVAApuntes.BeginEdit
  
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    
    ' restore object state
    LSet mudtprops = mudtSave
    mobjApuntes.CancelEdit
    mobjIVAApuntes.CancelEdit
  
End Sub

Public Sub ApplyEdit()
    Dim objPersist As AsientoPersist
  
    If Not mflgEditing Then Err.Raise 445
    
    ' comprobar que el saldo del asiento es cero
    If Not mobjApuntes.AsientoEquilibrado Then _
        Err.Raise vbObjectError + 1001, "El asiento no está equilibrado, el saldo no es cero."
    
    ' comprobar que tiene apuntes
    If mobjApuntes.Count = 0 Then _
        Err.Raise vbObjectError + 1001, "El asiento no tiene apuntes contables."


    Set objPersist = CreateObject("GCServerMTS.AsientoPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.Save (GetSuperState)
        
        mflgNew = True
        mflgDeleted = False

    ElseIf mflgDirty Or mobjApuntes.IsDirty Or mobjIVAApuntes.IsDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        ' save object to database if appropriate
        SetSuperState objPersist.Save(GetSuperState)
        
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False

    End If

    mobjApuntes.ApplyEdit 'mobjApuntes.AsientoID
    mobjIVAApuntes.ApplyEdit 'mobjIVAApuntes.AsientoID
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
  
    mudtprops.Numero = 0
    mudtprops.Ejercicio = CStr(Year(Date))
    mudtprops.Concepto = vbNullString
    mudtprops.FechaAlta = Date
    mudtprops.Situacion = "A" ' "A" --> Alta
    mudtprops.TemporadaID = 0
    mudtprops.EmpresaID = 0
      
    mobjValid.RuleBroken "Numero", True
    mobjValid.RuleBroken "Concepto", True
    mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "EmpresaID", True
        
    Set mobjApuntes = New Apuntes
    Set mobjIVAApuntes = New IVAApuntes
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjApuntes = Nothing
    Set mobjIVAApuntes = Nothing
    
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

Public Sub Load(AsientoID As Long, Moneda As String)
    Dim objPersist As AsientoPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.AsientoPersist", PERSIST_SERVER)
    SetSuperState objPersist.Fetch(AsientoID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Numero", False
    mobjValid.RuleBroken "Concepto", False
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "EmpresaID", False
        
  
    mMoneda = Moneda
  
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True
    mflgDirty = True

    mobjApuntes.Delete
    mobjIVAApuntes.Delete

End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty Or mobjApuntes.IsDirty Or mobjIVAApuntes.IsDirty

End Property

Public Property Get AsientoID() As Long

    AsientoID = mobjApuntes.AsientoID
    
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

Public Property Let Ejercicio(Value As String)

    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.Ejercicio) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mobjValid.RuleBroken "Ejercicio", Trim(Value) = vbNullString
      
    mudtprops.Ejercicio = Trim(Value)
    mflgDirty = True
    
End Property

Public Property Get Ejercicio() As String

    Ejercicio = Trim(mudtprops.Ejercicio)

End Property

Public Property Let Concepto(Value As String)

    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.Concepto) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mobjValid.RuleBroken "Concepto", Trim(Value) = vbNullString
      
    mudtprops.Concepto = Trim(Value)
    mflgDirty = True
    
End Property

Public Property Get Concepto() As String

    Concepto = Trim(mudtprops.Concepto)

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

Public Property Get Situacion() As String
  Select Case mudtprops.Situacion
  Case "A"
     Situacion = "Pendiente"
  Case "E"
     Situacion = "Exportado"
  Case Else
    Err.Raise vbObjectError + 1001, "Situacion del asiento erronea"
  End Select

End Property

Public Property Get Exportado() As Boolean
  
    Exportado = (mudtprops.Situacion = "E")
  
End Property

Public Property Get Apuntes() As Apuntes

    Set Apuntes = mobjApuntes

End Property

Public Property Get IVAApuntes() As IVAApuntes

    Set IVAApuntes = mobjIVAApuntes

End Property

Public Sub NuevoAsiento(Numero As Long, Ejercicio As String, Concepto As String, _
                        TemporadaID As Long, EmpresaID As Long)

    With Me
        .BeginEdit mMoneda
        .Numero = Numero
        .Ejercicio = Ejercicio
        .Concepto = Concepto
        .TemporadaID = TemporadaID
        .EmpresaID = EmpresaID
    End With
    
End Sub

Public Sub NuevoApunte(Cuenta As String, TipoImporte As String, _
                      Importe As Double, Descripcion As String, _
                      Fecha As Variant, Documento As String)
    Dim objApunte As Apunte

    If Importe = 0 Then Exit Sub
    
    ' En caso de que me informen de apuntes negativos invierto los signos
    If Importe < 0 Then
        Importe = -Importe
        If TipoImporte = "Debe" Then
            TipoImporte = "Haber"
        Else
            TipoImporte = "Debe"
        End If
    End If
    
    Set objApunte = Me.Apuntes.Add
    With objApunte
        .BeginEdit
    
        .Cuenta = Cuenta
        .TipoImporte = TipoImporte
        .Importe = Importe
        .Descripcion = Descripcion
        .Fecha = Fecha
        .Documento = Documento
        .ApplyEdit
    End With
    
End Sub

Public Sub NuevoIVAApunte(TipoApunte As String, NumeroFactura As String, _
                      CuentaIVA As String, CuentaTotal As String, _
                      CuentaBase As String, Titular As String, _
                      DNINIF As String, BaseImponible As Double, _
                      Total As Double, IVA As Double, _
                      CuotaIVA As Double, RecargoEquivalencia As Double, _
                      CuotaRecargo As Double)
    Dim objIVAApunte As IVAApunte

    Set objIVAApunte = Me.IVAApuntes.Add
    With objIVAApunte
        .BeginEdit
        
        .TipoApunte = TipoApunte
        .NumeroFactura = NumeroFactura
        .CuentaIVA = CuentaIVA
        .CuentaTotal = CuentaTotal
        .CuentaBase = CuentaBase
        .Titular = Titular
        .DNINIF = DNINIF
        .BaseImponible = BaseImponible
        .Total = Total
        .IVA = IVA
        .CuotaIVA = CuotaIVA
        .RecargoEquivalencia = RecargoEquivalencia
        .CuotaRecargo = CuotaRecargo
        .ApplyEdit
    End With
    
End Sub

Public Sub CuadrarAsiento(Cuenta As String, _
                          Fecha As Variant, Documento As String)
    Dim objApunte As Apunte

    If mobjApuntes.AsientoEquilibrado Then Exit Sub
    
    Set objApunte = Me.Apuntes.Add
    With objApunte
        .BeginEdit
    
        .Cuenta = Cuenta
        If mobjApuntes.Saldo < 0 Then
        'mobjApuntes.TotalDebe > mobjApuntes.TotalHaber Then
            .TipoImporte = "Haber"
            .Importe = -mobjApuntes.Saldo   'mobjApuntes.TotalDebe - mobjApuntes.TotalHaber
        Else
            .TipoImporte = "Debe"
            .Importe = mobjApuntes.Saldo 'mobjApuntes.TotalHaber - mobjApuntes.TotalDebe
        End If
        .Descripcion = "Cuadre automático"
        .Fecha = Fecha
        .Documento = Documento
        .ApplyEdit
    End With
    
End Sub

Private Function GetState() As String
    Dim udtData As AsientoData
    
    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With
  
    LSet udtData = mudtprops
    GetState = udtData.Buffer
    
End Function

Private Sub SetState(Buffer As String)
    Dim udtData As AsientoData
    
    udtData.Buffer = Buffer
    LSet mudtprops = udtData
End Sub

Friend Function GetSuperState() As String
    Dim objPB As PropertyBag
    
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty "State", GetState
        .WriteProperty "Items", mobjApuntes.GetSuperState
        .WriteProperty "IVAItems", mobjIVAApuntes.GetSuperState
        GetSuperState = .Contents
    End With
    Set objPB = Nothing
End Function

Friend Sub SetSuperState(Buffer As String)
    Dim objPB As PropertyBag
    Dim arBuffer() As Byte
    
    Set objPB = New PropertyBag
    arBuffer = Buffer
    With objPB
        .Contents = arBuffer
        SetState .ReadProperty("State")
        mobjApuntes.SetSuperState .ReadProperty("Items")
        mobjIVAApuntes.SetSuperState .ReadProperty("IVAItems")
    End With
    Set objPB = Nothing
End Sub


Public Sub Contabilizar(strServidorContawin As String)
    Dim objApunte As Apunte
    Dim strEmpresaContawin As String
    Dim objEmpresa As Empresa
    Dim Empresa  As Object 'cwoEmpresa
    Dim cnnEmpresa As Connection
    Dim cnnPrincipal As Connection
    'Dim rsApuntes As Recordset
    Dim rsEmpresas As Recordset
    Dim ContadorAsiento As Long
    Dim Apuntes As Object 'cwoDiario
    Dim myTipoImporte As Long 'TipoImporte
    Dim objIVAApunte As Object 'IVAApunte
    Dim Compras As Object 'cwoCompras
    Dim Ventas As Object 'cwoVentas
    
    If mflgEditing Then Err.Raise 383
     
    If mudtprops.Situacion = "E" Then _
        Err.Raise 1001 + vbObjectError, "Exportar Contawin", "El asiento a exportar (" & Me.AsientoID & ") ya está exportado."
        
    Set objEmpresa = New Empresa
    objEmpresa.Load Me.EmpresaID
    strEmpresaContawin = objEmpresa.EmpresaContawin
    Set objEmpresa = Nothing
    
    Set Empresa = CreateObject("ContawinDLL.cwoEmpresa") 'New cwoEmpresa
    'Set rsApuntes = New Recordset
    Set rsEmpresas = New Recordset
    Set cnnEmpresa = New Connection
    Set cnnPrincipal = New Connection
    
    cnnEmpresa.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & strEmpresaContawin & ";Data Source=" & strServidorContawin & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False"
    cnnEmpresa.Open
    
    cnnPrincipal.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Contawin;Data Source=" & strServidorContawin & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False"
    cnnPrincipal.Open
    
    If Not Empresa.AbrirEmpresas(cnnPrincipal, rsEmpresas) Then _
        Err.Raise 1001 + vbObjectError, "Exportación Contawin", "Error abriendo empresas Contawin."
    
    Set Apuntes = CreateObject("ContawinDLL.cwoDiario") 'New cwoDiario
    ContadorAsiento = Apuntes.ProximoNumeroDeAsiento(cnnEmpresa, CInt(Me.Ejercicio))
                    
    If Apuntes.CodigoDeError <> 0 Then _
        Err.Raise 1001 + vbObjectError, "Exportación Contawin", "Error obteniendo contador del diario:" & Apuntes.InfoDeError

    For Each objApunte In mobjApuntes
        myTipoImporte = IIf(objApunte.CodigoTipoImporte = TipoImporteDebe, 0, 1) 'Debe, Haber)
        
        If Not Apuntes.Apunte(cnnEmpresa, ContadorAsiento, objApunte.Cuenta, objApunte.Importe, myTipoImporte, , objApunte.Descripcion, objApunte.Fecha, CInt(Me.Ejercicio), objApunte.Documento, , , , 0) Then _
            Err.Raise 1001 + vbObjectError, "Exportación Contawin, " & Apuntes.InfoDeError, "Error generando apunte" & CStr(objApunte.Importe) & " " & objApunte.Cuenta & " " & objApunte.TipoImporte & " " & objApunte.ApunteID

    Next
    
    Set Apuntes = Nothing
    
    Set Compras = CreateObject("ContawinDLL.cwoCompras") 'New cwoCompras
    Set Ventas = CreateObject("ContawinDLL.cwoVentas") 'New cwoVentas
    
    For Each objIVAApunte In mobjIVAApuntes
        With objIVAApunte
            If .TipoApunte = "S" Then _
                Compras.IvaSoportado cnnEmpresa, .NumeroFactura, .CuentaIVA, .CuentaTotal, .CuentaBase, _
                        .Titular, .DNINIF, .BaseImponible, .Total, .IVA, .CuotaIVA, .RecargoEquivalencia, .CuotaRecargo, _
                        ContadorAsiento, mobjApuntes(1).Fecha, Me.Ejercicio
            If .TipoApunte = "R" Then _
                Ventas.IvaRepercutido cnnEmpresa, .NumeroFactura, .CuentaIVA, .CuentaTotal, .CuentaBase, _
                         .Titular, .DNINIF, .BaseImponible, .Total, .IVA, .CuotaIVA, .RecargoEquivalencia, .CuotaRecargo, _
                        ContadorAsiento, mobjApuntes(1).Fecha, Me.Ejercicio
        End With

    Next
    
    Set Compras = Nothing
    Set Ventas = Nothing
    
    Set Empresa = Nothing
    'Set rsApuntes = Nothing
    
    rsEmpresas.Close
    Set rsEmpresas = Nothing
    
    cnnEmpresa.Close
    Set cnnEmpresa = Nothing
    
    cnnPrincipal.Close
    Set cnnPrincipal = Nothing
    

    ' hay que informar del tipo de documento?????.
    ' tdaFacturaDeVentas
    
    Me.BeginEdit "EUR"
    mudtprops.Situacion = "E"
    mflgDirty = True
    Me.ApplyEdit
    
End Sub
