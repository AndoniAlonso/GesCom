VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Apunte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private mudtprops As ApunteProps
Private mudtSave As ApunteProps
Private mudtChild As ApunteProps
Private mobjParent As Apuntes

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
    
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    If Not mflgChild Then
'        If mflgDeleted And Not mflgNew Then
'            ' code to delete the object's data goes here
'            objPersist.DeleteObject mudtprops.ApunteID
'            mflgNew = True
'            mflgDeleted = False
'
'        ElseIf mflgDirty Or mflgNew Then
'            If Not IsValid Then Err.Raise 445
'            ' save object to database if appropriate
'            SetState objPersist.Save(GetState)
'            ' save object state
'            LSet mudtSave = mudtprops
'            mflgNew = False
'
'        End If
'
'        mflgDirty = False
'
    ElseIf mflgNew Then
        If Not mobjParent Is Nothing Then
            mobjParent.AddApunte Me
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
    '
    mobjValid.RuleBroken "Cuenta", True
    mobjValid.RuleBroken "Importe", True
    mobjValid.RuleBroken "Descripcion", True
    
    With mudtprops
        .ApunteID = 0
        .AsientoID = 0
        .Cuenta = vbNullString
        .Descripcion = vbNullString
        .Documento = vbNullString
        .Fecha = Date
        .ImporteEUR = 0
        .ImportePTA = 0
        .TipoApunte = TipoApunteEuros
        .TipoImporte = TipoImporteDebe
    
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

Public Sub Load()
End Sub

Public Sub Delete()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
  
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

Friend Sub SetAsChild(objApuntes As Apuntes)

    If mflgNew Then Set mobjParent = objApuntes
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

Public Property Get ApunteID() As Long

    ApunteID = mudtprops.ApunteID

End Property

Public Property Get AsientoID() As Long

    AsientoID = mudtprops.AsientoID

End Property

Public Property Get Cuenta() As String

    Cuenta = Trim(mudtprops.Cuenta)

End Property

Public Property Let Cuenta(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Cuenta) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - todos numericos
    '     - longitud exacta a 10
    mobjValid.RuleBroken "Cuenta", _
        ((Len(Trim(Value)) < Len(mudtprops.Cuenta)) Or (Not IsNumeric(Value)))
  
    mudtprops.Cuenta = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get CodigoTipoImporte() As Long

    CodigoTipoImporte = mudtprops.TipoImporte

End Property

Public Property Get TipoImporte() As String

    Select Case mudtprops.TipoImporte
    Case TipoImporteDebe
        TipoImporte = "Debe"
    Case TipoImporteHaber
        TipoImporte = "Haber"
    Case Else
        Err.Raise vbObjectError + 1001, "Error en tipo de importe solo puede ser debe/haber."
    End Select

End Property

Public Property Let TipoImporte(Value As String)

    If Not mflgEditing Then Err.Raise 383
    
    Select Case Value
    Case "Debe"
        mudtprops.TipoImporte = TipoImporteDebe
    Case "Haber"
        mudtprops.TipoImporte = TipoImporteHaber
    Case Else
  ' Solo puede ser debe o haber
        mobjValid.RuleBroken "TipoImporte", Not (Value = TipoImporteDebe Or _
                                            Value = TipoImporteHaber)
    End Select
    
    mflgDirty = True

End Property

Public Property Let Importe(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "Importe", (Value < 0)
  
    mudtprops.ImporteEUR = Round(Value, 2)
    mudtprops.ImportePTA = EUR2PTA(Value)
    
    mflgDirty = True

End Property

Public Property Get Importe() As Double

    Importe = mudtprops.ImporteEUR
  
End Property

Public Property Get Descripcion() As String
  
    Descripcion = Trim(mudtprops.Descripcion)

End Property

Public Property Let Descripcion(Value As String)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.Descripcion) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.Descripcion = Trim(Value)
    mflgDirty = True
    mobjValid.RuleBroken "Descripcion", (Len(Trim(Value)) = 0)

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

Public Property Get Documento() As String
  
    Documento = Trim(mudtprops.Documento)

End Property

Public Property Let Documento(Value As String)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.Documento) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.Documento = Trim(Value)
    mflgDirty = True

End Property

Public Property Get TipoApunte() As Long

    TipoApunte = mudtprops.TipoApunte

End Property

Friend Function GetSuperState() As String
    
    GetSuperState = GetState

End Function

Friend Sub SetSuperState(Buffer As String)
    
    SetState Buffer

End Sub

Private Function GetState() As String

    Dim udtData As ApunteData
  
    With mudtprops
        .IsNew = mflgNew
        .IsDeleted = mflgDeleted
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer
    mflgNew = False

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As ApunteData

    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    mobjValid.RuleBroken "Cuenta", False
    mobjValid.RuleBroken "Importe", False
    mobjValid.RuleBroken "Descripcion", False

End Sub

