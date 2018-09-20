VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Traspaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As TraspasoProps
Private mudtSave As TraspasoProps

Private mobjTraspasoItems As TraspasoItems

' Aqui declaro los ComboBox de la lista de clientes, representantes, transportistas y formas de pago.
Private mobjAlmacenesOrigen As TextList
Private mobjAlmacenesDestino As TextList

Event Valid(IsValid As Boolean)

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean

Public Enum enuTraspasoSituacion
    enuTraspasoSituacionAlta = 0
    enuTraspasoSituacionEnTransito = 1
    enuTraspasoSituacionRecepcionado = 2
End Enum

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit()

    If mflgEditing Then Err.Raise 445
  
    ' save object state
    LSet mudtSave = mudtprops
  
    mobjTraspasoItems.BeginEdit
  
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    
    ' restore object state
    LSet mudtprops = mudtSave
    mobjTraspasoItems.CancelEdit
  
End Sub

Public Sub ApplyEdit()
    Dim objPersist As TraspasoPersist
    Dim strItems As String
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.TraspasoPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
       
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.TraspasoID, mobjTraspasoItems.GetState
    
        mflgNew = True
        mflgDeleted = False
    
    ElseIf IsDirty Or mobjTraspasoItems.IsDirty Or mflgNew Then
            If Not IsValid Then Err.Raise 445
            
            strItems = mobjTraspasoItems.GetState
            SetState objPersist.Save(GetState, strItems)
            mobjTraspasoItems.SetState (strItems)
            
            ' save object state
            LSet mudtSave = mudtprops
            mflgNew = False
    
    End If

    mobjTraspasoItems.ApplyEdit  ' mudtprops.TraspasoID
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
      
    mudtprops.AlmacenOrigenID = 0
    mudtprops.AlmacenDestinoID = 0
    mudtprops.Situacion = enuTraspasoSituacionAlta
    mudtprops.FechaAlta = Now
    mudtprops.FechaTransito = CDate("00:00:00")
    mudtprops.FechaRecepcion = CDate("00:00:00")
      
    Set mobjAlmacenesOrigen = New TextList
    mobjAlmacenesOrigen.Load "Almacenes"
    mudtprops.AlmacenOrigenID = mobjAlmacenesOrigen.Key(mobjAlmacenesOrigen.IndexItem(1))
      
    Set mobjAlmacenesDestino = New TextList
    mobjAlmacenesDestino.Load "Almacenes"
    mudtprops.AlmacenDestinoID = mobjAlmacenesDestino.Key(mobjAlmacenesDestino.IndexItem(1))
      
    mobjValid.RuleBroken "AlmacenOrigenID", True
    mobjValid.RuleBroken "AlmacenDestinoID", True
    
    Set mobjTraspasoItems = New TraspasoItems
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjAlmacenesOrigen = Nothing
    Set mobjAlmacenesDestino = Nothing
    
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

Public Sub Load(TraspasoID As Long)
    Dim objPersist As TraspasoPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.TraspasoPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(TraspasoID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "AlmacenOrigenID", False
    mobjValid.RuleBroken "AlmacenDestinoID", False
  
    mobjTraspasoItems.Load TraspasoID
    
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
      
    If Me.Situacion = enuTraspasoSituacionEnTransito Or _
       Me.Situacion = enuTraspasoSituacionRecepcionado Then _
       Err.Raise vbObjectError + 1001, "El movimiento entre almacenes ya se ha ejecutado, no se puede eliminar."

    mflgDeleted = True
    mflgDirty = True

    mobjTraspasoItems.Delete

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

Public Property Get TraspasoID() As Long

    TraspasoID = mudtprops.TraspasoID
  
End Property

Public Property Get Situacion() As enuTraspasoSituacion  'Long

    Situacion = mudtprops.Situacion
  
End Property

Public Property Get AlmacenOrigenID() As Long

    AlmacenOrigenID = mudtprops.AlmacenOrigenID
  
End Property

' Esta funcion lee el cliente para asignar sus propiedades por defecto
Public Property Let AlmacenOrigen(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Me.Situacion = enuTraspasoSituacionEnTransito Or _
       Me.Situacion = enuTraspasoSituacionRecepcionado Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar el almacen origen, ya se ha entregado."
    
    mobjValid.RuleBroken "AlmacenOrigenID", (mobjAlmacenesOrigen.Key(Value) = 0)
    
    mudtprops.AlmacenOrigenID = mobjAlmacenesOrigen.Key(Value)
    
    mflgDirty = True

End Property

Public Property Get AlmacenOrigen() As String

    AlmacenOrigen = mobjAlmacenesOrigen.Item(mudtprops.AlmacenOrigenID)

End Property

Public Property Get AlmacenesOrigen() As TextList

    Set AlmacenesOrigen = mobjAlmacenesOrigen

End Property

Public Property Get FechaAlta() As Date

    FechaAlta = mudtprops.FechaAlta
    
End Property

Public Property Get FechaTransito() As Variant

    With mudtprops
        FechaTransito = IIf(.FechaTransito = 0, vbNullString, .FechaTransito)
    End With

End Property

Public Property Get FechaRecepcion() As Variant

    With mudtprops
        FechaRecepcion = IIf(.FechaRecepcion = 0, vbNullString, .FechaRecepcion)
    End With

End Property

Public Property Get AlmacenDestinoID() As Long

    AlmacenDestinoID = mudtprops.AlmacenDestinoID
      
End Property

Public Property Let AlmacenDestino(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Me.Situacion = enuTraspasoSituacionEnTransito Or _
       Me.Situacion = enuTraspasoSituacionRecepcionado Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar el almacen destino, ya se ha entregado."
    
    mudtprops.AlmacenDestinoID = mobjAlmacenesDestino.Key(Value)
  
    mobjValid.RuleBroken "AlmacenDestinoID", (mudtprops.AlmacenDestinoID = 0)
    
    mflgDirty = True

End Property

Public Property Get AlmacenDestino() As String

    AlmacenDestino = mobjAlmacenesDestino.Item(mudtprops.AlmacenDestinoID)

End Property

Public Property Get AlmacenesDestino() As TextList

    Set AlmacenesDestino = mobjAlmacenesDestino

End Property

Public Property Get TraspasoItems() As TraspasoItems

    Set TraspasoItems = mobjTraspasoItems

End Property

Public Property Get Cantidad() As Double
  
  Cantidad = mobjTraspasoItems.Cantidad
  
End Property

Public Sub Enviar()
    Dim objPersist As TraspasoPersist
    Dim strItems As String
    Dim objParametro As ParametroAplicacion
    Dim lngAlmacenTransito As Long
    
    ' Error si estamos editando el registro
    If mflgEditing Then Err.Raise 383
    
    ' Error si es un registro nuevo
    If Me.IsNew Then Err.Raise 383
    
    If Me.Situacion = enuTraspasoSituacionRecepcionado Then _
        Err.Raise vbObjectError + 1001, "No se puede realizar el envío, ya se ha recepcionado."
    
    If Me.Situacion = enuTraspasoSituacionEnTransito Then _
        Err.Raise vbObjectError + 1001, "No se puede realizar el envío, ya se ha enviado."
    
    Set objPersist = CreateObject("GCServerMTS.TraspasoPersist", PERSIST_SERVER)
    
    Set objParametro = New ParametroAplicacion
    lngAlmacenTransito = objParametro.AlmacenTransito
    
    mudtprops.FechaTransito = Now
    mudtprops.Situacion = enuTraspasoSituacionEnTransito
    
    mobjTraspasoItems.Enviar
    strItems = mobjTraspasoItems.GetState
    SetState objPersist.Enviar(GetState, strItems, lngAlmacenTransito)
            
    Set objParametro = Nothing
    Set objPersist = Nothing

End Sub


Public Sub Recepcionar()
    Dim objPersist As TraspasoPersist
    Dim strItems As String
    Dim objParametro As ParametroAplicacion
    Dim lngAlmacenTransito As Long
    
    ' Error si estamos editando el registro
    If mflgEditing Then Err.Raise 383
    
    ' Error si es un registro nuevo
    If Me.IsNew Then Err.Raise 383
    
    If Me.Situacion = enuTraspasoSituacionAlta Then _
        Err.Raise vbObjectError + 1001, "No se puede realizar la recepción, está pendiente de enviarse."
    
    If Me.Situacion = enuTraspasoSituacionRecepcionado Then _
        Err.Raise vbObjectError + 1001, "No se puede realizar la recepción, ya se ha recepcionado."
    
    Set objPersist = CreateObject("GCServerMTS.TraspasoPersist", PERSIST_SERVER)
    
    Set objParametro = New ParametroAplicacion
    lngAlmacenTransito = objParametro.AlmacenTransito
    
    mudtprops.FechaRecepcion = Now
    mudtprops.Situacion = enuTraspasoSituacionRecepcionado
    
    mobjTraspasoItems.Recepcionar
    strItems = mobjTraspasoItems.GetState
    SetState objPersist.Recepcionar(GetState, strItems, lngAlmacenTransito)
            
    Set objParametro = Nothing
    Set objPersist = Nothing

End Sub

Public Sub EnviarYRecepcionar()
    Dim objPersist As TraspasoPersist
    Dim strItems As String
    Dim objParametro As ParametroAplicacion
    'Dim lngAlmacenTransito As Long
    
    ' Error si estamos editando el registro
    If mflgEditing Then Err.Raise 383
    
    ' Error si es un registro nuevo
    If Me.IsNew Then Err.Raise 383
    
    If Me.Situacion = enuTraspasoSituacionRecepcionado Then _
        Err.Raise vbObjectError + 1001, "No se puede realizar el envío, ya se ha recepcionado."
    
    If Me.Situacion = enuTraspasoSituacionEnTransito Then _
        Err.Raise vbObjectError + 1001, "No se puede realizar el envío, ya se ha enviado."
    
    Set objPersist = CreateObject("GCServerMTS.TraspasoPersist", PERSIST_SERVER)
    
    Set objParametro = New ParametroAplicacion
    'lngAlmacenTransito = objParametro.AlmacenTransito
    
    mudtprops.FechaTransito = Now
    mudtprops.FechaRecepcion = mudtprops.FechaTransito
    mudtprops.Situacion = enuTraspasoSituacionRecepcionado
    
    mobjTraspasoItems.Recepcionar
    strItems = mobjTraspasoItems.GetState
    SetState objPersist.EnviarYRecepcionar(GetState, strItems)
            
    Set objParametro = Nothing
    Set objPersist = Nothing

End Sub

' A partir de los items de albaran, un articulo a ubicar y la lista de pedidos pendientes,
' ubicar el albaran en una linea:
'4.- Ubicar el artículo en una nueva linea no relacionada con pedidos.
' La función devuelve el string informativo del error en caso de que haya alguna incidencia.
Public Function TraspasoItemCodigoBarras(ByVal intTalla As Integer, _
                                        ByVal ArticuloColorID As Long) As String
Dim objTraspasoItem As TraspasoItem
Dim dblCantidadActual As Double
Dim strResultadoFuncion As String


    strResultadoFuncion = vbNullString
    
    Set objTraspasoItem = Me.TraspasoItems.Add
    
    With objTraspasoItem
        .BeginEdit
        .ArticuloColorID = ArticuloColorID
        '.AlmacenID = mudtprops.AlmacenID
        ' Incrementamos la cantidad actual por talla
        dblCantidadActual = CallByName(objTraspasoItem, "CantidadT" & CStr(intTalla), VbGet)
        CallByName objTraspasoItem, "CantidadT" & CStr(intTalla), VbLet, dblCantidadActual + 1

        .ApplyEdit

    End With
    
    TraspasoItemCodigoBarras = strResultadoFuncion
    Set objTraspasoItem = Nothing
End Function

Private Function GetState() As String
    Dim udtData As TraspasoData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As TraspasoData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
