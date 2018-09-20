VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Modelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As ModeloProps
Private mudtSave As ModeloProps

Private mobjEstrModelos As EstrModelos

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
  
    If mobjEstrModelos Is Nothing Then
        CargarEstrModelos
    End If
    mobjEstrModelos.BeginEdit
  
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    
    ' restore object state
    LSet mudtprops = mudtSave
    
    If mobjEstrModelos Is Nothing Then
        CargarEstrModelos
    End If
    mobjEstrModelos.CancelEdit
  
End Sub

Public Sub ApplyEdit()
    Dim objPersist As ModeloPersist
    Dim strItems As String
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.ModeloPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.ModeloID, mobjEstrModelos.GetState
        mflgNew = True
        mflgDeleted = False

    ElseIf mflgDirty Or mobjEstrModelos.IsDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        ' save object to database if appropriate
        strItems = mobjEstrModelos.GetState
        SetState objPersist.Save(GetState, strItems)
        mobjEstrModelos.SetState (strItems)
        
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False

    End If

    mobjEstrModelos.ApplyEdit 'mudtprops.ModeloID
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
  
    mudtprops.TemporadaID = 0
    mudtprops.Beneficio = 0
    mudtprops.BeneficioPVP = 0
    mudtprops.CantidadTela = 0
    mudtprops.CorteEUR = 0
    mudtprops.TallerEUR = 0
      
    mobjValid.RuleBroken "Nombre", True
    mobjValid.RuleBroken "Codigo", True
    mobjValid.RuleBroken "TemporadaID", True
    'mobjValid.RuleBroken "Moneda", True
        
    ' Ya no hace falta inicializarlo aquí.
    'Set mobjEstrModelos = New EstrModelos
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjEstrModelos = Nothing
    
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

Public Sub Load(ModeloID As Long, Optional Codigo As String, _
                Optional TemporadaID As Long, Optional Nombre As String)

    Dim objPersist As ModeloPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.ModeloPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(ModeloID, Codigo, TemporadaID, Nombre)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Nombre", False
    mobjValid.RuleBroken "Codigo", False
    mobjValid.RuleBroken "TemporadaID", False
    'mobjValid.RuleBroken "Moneda", False
  
    'mobjEstrModelos.Load ModeloID
  
    mflgNew = False

End Sub

' Con este método implementamos un Lazy load de las estructuras de modelos.
Private Sub CargarEstrModelos()
    
    Set mobjEstrModelos = New EstrModelos
    If Not mflgNew Then
        mobjEstrModelos.Load ModeloID
    End If
End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True
    mflgDirty = True

    If mobjEstrModelos Is Nothing Then
        CargarEstrModelos
    End If
    mobjEstrModelos.Delete

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

Public Property Get ModeloID() As Long

    ModeloID = mudtprops.ModeloID
  
End Property

Public Property Let Nombre(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Nombre) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    mobjValid.RuleBroken "Nombre", (Len(Trim(Value)) = 0)

    mudtprops.Nombre = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Nombre() As String

    Nombre = Trim(mudtprops.Nombre)

End Property

Public Property Let Codigo(Value As String)

    If Not mflgNew Then Err.Raise 383
    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Codigo) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    mobjValid.RuleBroken "Codigo", (Len(Trim(Value)) <> Len(mudtprops.Codigo))

    mudtprops.Codigo = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Codigo() As String

    Codigo = Trim(mudtprops.Codigo)

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

Public Property Let Beneficio(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    mobjValid.RuleBroken "Beneficio", (Value < 0)
  
    mudtprops.Beneficio = Round(Value, 2)
  
    mflgDirty = True

End Property

Public Property Get Beneficio() As Double

    Beneficio = mudtprops.Beneficio

End Property

Public Property Let BeneficioPVP(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    mobjValid.RuleBroken "BeneficioPVP", (Value < 0)
  
    mudtprops.BeneficioPVP = Round(Value, 2)
  
    mflgDirty = True

End Property

Public Property Get BeneficioPVP() As Double

    BeneficioPVP = mudtprops.BeneficioPVP

End Property

Public Property Let CantidadTela(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    mobjValid.RuleBroken "CantidadTela", (Value < 0)
  
    mudtprops.CantidadTela = Round(Value, 2)
  
    mflgDirty = True

End Property

Public Property Get CantidadTela() As Double

    CantidadTela = mudtprops.CantidadTela

End Property

Public Property Let Corte(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "Corte", (Value < 0)
  
    mudtprops.CorteEUR = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get Corte() As Double

    Corte = mudtprops.CorteEUR
  
End Property

Public Property Let Taller(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "Taller", (Value < 0)
  
    mudtprops.TallerEUR = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get Taller() As Double

    Taller = mudtprops.TallerEUR
  
End Property

Public Property Get PrecioCoste() As Double

    If mobjEstrModelos Is Nothing Then
        CargarEstrModelos
    End If
    PrecioCoste = Taller + Corte + mobjEstrModelos.Precio
  
End Property

Public Function CrearCodigoAutomatico(Nombre As String, TemporadaID As Long) As String
    Dim objModelo As Modelo
    Dim strCodigo As String
        
    On Error GoTo ErrorManager
    
    strCodigo = Left(Nombre, 3)
    CrearCodigoAutomatico = strCodigo
    
    Do
        CrearCodigoAutomatico = strCodigo
        Set objModelo = New Modelo
        
        On Error Resume Next
        objModelo.Load 0, strCodigo, TemporadaID, vbNullString
        On Error GoTo ErrorManager
    
        strCodigo = ObtenerSiguienteCodigo(strCodigo, 3)
    Loop Until objModelo.ModeloID = 0
    
    Set objModelo = Nothing
    
    Exit Function

ErrorManager:
    CrearCodigoAutomatico = vbNullString
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Exit Function
End Function


Public Property Get EstrModelos() As EstrModelos

    If mobjEstrModelos Is Nothing Then
        CargarEstrModelos
    End If
    Set EstrModelos = mobjEstrModelos

End Property

Private Function GetState() As String

    Dim udtData As ModeloData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As ModeloData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
