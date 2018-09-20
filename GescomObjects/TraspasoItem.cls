VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TraspasoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private mobjArticuloColor As ArticuloColor

Private mstrCodigoArticuloColor As String

Private mudtprops As TraspasoItemProps
Private mudtSave As TraspasoItemProps
Private mudtChild As TraspasoItemProps
Private mobjParent As TraspasoItems
Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChild As Boolean
Private mflgChildEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    
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
    ' restore object state
    LSet mudtprops = mudtSave
    
End Sub

Public Sub ApplyEdit()
    Dim objPersist As TraspasoItemPersist
    
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    Set objPersist = CreateObject("GCServerMTS.TraspasoItemPersist", PERSIST_SERVER)
    
    If Not mflgChild Then
        If mflgDeleted And Not mflgNew Then
            ' code to delete the object's data goes here
            objPersist.DeleteObject mudtprops.TraspasoItemID
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
    
    ElseIf mflgNew Then
        If Not mobjParent Is Nothing Then
            mobjParent.AddTraspasoItem Me
            Set mobjParent = Nothing
        End If
    
    End If
      
    Set objPersist = Nothing
    mflgEditing = False
    
End Sub

Private Sub Class_Initialize()
    
    mflgNew = True
    Set mobjValid = New BrokenRules
    
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    mobjValid.RuleBroken "ArticuloColorID", True
    mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "Cantidad", True
    
    With mudtprops
        .Situacion = enuTraspasoSituacionAlta    ' Marco la situacion del pedido como --> alta,
                                                 ' pendiente de codificar las posibles situaciones.
        .ArticuloColorID = 0
        .CantidadT36 = 0
        .CantidadT38 = 0
        .CantidadT40 = 0
        .CantidadT42 = 0
        .CantidadT44 = 0
        .CantidadT46 = 0
        .CantidadT48 = 0
        .CantidadT50 = 0
        .CantidadT52 = 0
        .CantidadT54 = 0
        .CantidadT56 = 0
        .Observaciones = vbNullString
    End With
    
End Sub

Private Sub Class_Terminate()
    
    Set mobjValid = Nothing
    
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

Public Sub Load(TraspasoItemID As Long)
    Dim objPersist As TraspasoItemPersist
    
    If mflgChild Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.TraspasoItemPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(TraspasoItemID)
    Set objPersist = Nothing
    
    mobjValid.RuleBroken "ArticuloColorID", False
    mobjValid.RuleBroken "Cantidad", False
    
    mflgNew = False
    
End Sub

Public Sub Delete()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    If mudtprops.Situacion <> enuTraspasoSituacionAlta Then _
        Err.Raise vbObjectError + 1001, "No se puede borrar, ya se han servido entregas"
    
    mflgDeleted = True
    mflgDirty = True
    
End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew
    
End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty Or mflgNew

End Property

Friend Sub SetAsChild(objTraspasoItems As TraspasoItems)

    If mflgNew Then Set mobjParent = objTraspasoItems
    mflgChild = True
    
End Sub

Friend Sub ChildBeginEdit()

    If mflgChildEditing Then Err.Raise 445
    
    ' save object state
    LSet mudtChild = mudtprops
    mflgChildEditing = True
    
End Sub

Friend Sub ChildCancelEdit()

    If Not mflgChildEditing Then Err.Raise 445
    
    mflgChildEditing = False
    mflgDeleted = False
    ' restore object state
    LSet mudtprops = mudtChild
    
End Sub

Public Property Get TraspasoItemID() As Long

    TraspasoItemID = mudtprops.TraspasoItemID

End Property

Public Property Get TraspasoID() As Long

    TraspasoID = mudtprops.TraspasoID

End Property

Public Property Get Situacion() As String
    
'    Select Case mudtprops.Situacion
'    Case "A"
'        Situacion = "Pendiente"
'    Case "C"
'        Situacion = "Servido"
'    Case Else
'        Err.Raise vbObjectError + 1001, "Situacion de pedido erronea"
'    End Select
    
End Property

Public Property Get ArticuloColorID() As Long

    ArticuloColorID = mudtprops.ArticuloColorID
  
End Property

Public Property Let ArticuloColorID(Value As Long)

    If Not mflgEditing Then Err.Raise 383
    
    Set mobjArticuloColor = New ArticuloColor
    
    mobjArticuloColor.Load Value
    
    mobjValid.RuleBroken "ArticuloColorID", False

'    mudtprops.NombreArticuloColor = mobjArticuloColor.ObjArticulo.NombreCompleto & " " & mobjArticuloColor.NombreColor

    mudtprops.ArticuloColorID = mobjArticuloColor.ArticuloColorID
 
    mflgDirty = True
 
End Property

Public Property Get CodigoArticuloColor() As String
    
    CodigoArticuloColor = mstrCodigoArticuloColor
    
End Property

'Public Property Get NombreArticuloColor() As String
'
'    NombreArticuloColor = mudtprops.NombreArticuloColor
'
'End Property

'Public Property Get objArticuloColor() As ArticuloColor
'
'    If mudtprops.ArticuloColorID = 0 Then
'        Set objArticuloColor = Nothing
'        Exit Property
'    End If
'
'    If mobjArticuloColor Is Nothing Then _
'        Set mobjArticuloColor = New ArticuloColor
'
'    mobjArticuloColor.Load mudtprops.ArticuloColorID, "EUR"
'    Set objArticuloColor = mobjArticuloColor
'
'End Property

'Public Property Get ArticuloColor() As String
'
'    ArticuloColor = Trim(mudtprops.NombreArticuloColor)
'
'End Property

Public Property Let CantidadT36(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
'    If Servido Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
'    If mudtprops.SituacionCorte = "C" Then _
'        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT36", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT36 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    mflgDirty = True
    
End Property

Public Property Get CantidadT36() As Double

    CantidadT36 = mudtprops.CantidadT36

End Property

Public Property Let CantidadT38(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
'    If Servido Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
'    If mudtprops.SituacionCorte = "C" Then _
'        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT38", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT38 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    
    mflgDirty = True
    
End Property

Public Property Get CantidadT38() As Double

    CantidadT38 = mudtprops.CantidadT38

End Property

Public Property Let CantidadT40(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
'    If Servido Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
'    If mudtprops.SituacionCorte = "C" Then _
'        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT40", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT40 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    
    mflgDirty = True
    
End Property

Public Property Get CantidadT40() As Double

    CantidadT40 = mudtprops.CantidadT40

End Property

Public Property Let CantidadT42(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
'    If Servido Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
'    If mudtprops.SituacionCorte = "C" Then _
'        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT42", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT42 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    
    mflgDirty = True
    
End Property

Public Property Get CantidadT42() As Double

    CantidadT42 = mudtprops.CantidadT42

End Property

Public Property Let CantidadT44(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
'    If Servido Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
'    If mudtprops.SituacionCorte = "C" Then _
'        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT44", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT44 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    
    mflgDirty = True
    
End Property

Public Property Get CantidadT44() As Double

    CantidadT44 = mudtprops.CantidadT44

End Property

Public Property Let CantidadT46(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
'    If Servido Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
'    If mudtprops.SituacionCorte = "C" Then _
'        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT46", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT46 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    
    mflgDirty = True
    
End Property

Public Property Get CantidadT46() As Double

    CantidadT46 = mudtprops.CantidadT46

End Property

Public Property Let CantidadT48(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
'    If Servido Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
'    If mudtprops.SituacionCorte = "C" Then _
'        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT48", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT48 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    mflgDirty = True
    
End Property

Public Property Get CantidadT48() As Double

    CantidadT48 = mudtprops.CantidadT48

End Property

Public Property Let CantidadT50(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
'    If Servido Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
'    If mudtprops.SituacionCorte = "C" Then _
'        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT50", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT50 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    
    mflgDirty = True
    
End Property

Public Property Get CantidadT50() As Double

    CantidadT50 = mudtprops.CantidadT50

End Property

Public Property Let CantidadT52(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
'    If Servido Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
'    If mudtprops.SituacionCorte = "C" Then _
'        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT52", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT52 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    
    mflgDirty = True
    
End Property

Public Property Get CantidadT52() As Double

    CantidadT52 = mudtprops.CantidadT52

End Property

Public Property Let CantidadT54(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
'    If Servido Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
'    If mudtprops.SituacionCorte = "C" Then _
'        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT54", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT54 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    
    mflgDirty = True
    
End Property

Public Property Get CantidadT54() As Double

    CantidadT54 = mudtprops.CantidadT54

End Property

Public Property Let CantidadT56(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
'    If Servido Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han servido entregas"
    
'    If mudtprops.SituacionCorte = "C" Then _
'        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha generado orden de corte"
    
    ' No se permiten Cantidades negativas
    ' y la cantidad de prendas debe ser mayor que uno
    mobjValid.RuleBroken "CantidadT56", (Value < 0)
    
    ' Las cantidades son enteras
    mudtprops.CantidadT56 = Round(Value, 0)
    
    ' Se valida la cantidad total de la linea de pedidos
    mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    
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

Public Sub Enviar()
    mudtprops.Situacion = enuTraspasoSituacionEnTransito
End Sub

Public Sub Recepcionar()
    mudtprops.Situacion = enuTraspasoSituacionRecepcionado
End Sub

Friend Function GetState() As String

  Dim udtData As TraspasoItemData
  
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
    Dim udtData As TraspasoItemData
    
    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    mobjValid.RuleBroken "ArticuloColorID", False
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "Cantidad", False
    'mstrCodigoArticuloColor = Left(mudtprops.NombreArticuloColor, 8)
    
End Sub
