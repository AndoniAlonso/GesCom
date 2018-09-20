VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Serie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As SerieProps
Private mudtSave As SerieProps

Event Valid(IsValid As Boolean)

' Aqui declaro el ComboBox de la lista de materiales.
Private mobjMateriales As TextList

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
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    ' restore object state
    LSet mudtprops = mudtSave
  
End Sub

Public Sub ApplyEdit()
    Dim objPersist As SeriePersist
    
    If Not mflgEditing Then Err.Raise 445
    
    Set objPersist = CreateObject("GCServerMTS.SeriePersist", PERSIST_SERVER)
    
    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.SerieID
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
'    Set mobjMateriales = New TextList
'    mobjMateriales.Load "vNombreMateriales"
'    mudtprops.MaterialID = mobjMateriales.Key(mobjMateriales.IndexItem(1))
    Set mobjMateriales = Nothing
    mudtprops.MaterialID = 0
  
    mudtprops.TemporadaID = 0
    
    'Anchura de la tela estandar, provendrá del material de la serie.
    mudtprops.AnchuraTela = dblAnchuraTelaEstandar
      
    mobjValid.RuleBroken "Nombre", True
    mobjValid.RuleBroken "Codigo", True
    mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "MaterialID", True
        
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjMateriales = Nothing

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

Public Sub Load(SerieID As Long, Optional Codigo As String, _
                Optional TemporadaID As Long, Optional Nombre As String)
    Dim objPersist As SeriePersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.SeriePersist", PERSIST_SERVER)
    SetState objPersist.Fetch(SerieID, Codigo, TemporadaID, Nombre)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Nombre", False
    mobjValid.RuleBroken "Codigo", False
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "MaterialID", False
  
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True

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

Public Property Get SerieID() As Long

    SerieID = mudtprops.SerieID
  
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

Public Property Get MaterialID() As Long

    MaterialID = mudtprops.MaterialID
  
End Property

Public Property Let Material(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If mobjMateriales Is Nothing Then CargarListaMateriales
    
    mudtprops.MaterialID = mobjMateriales.Key(Value)
  
    mobjValid.RuleBroken "MaterialID", (mudtprops.MaterialID = 0)
  
    mflgDirty = True

End Property

Public Property Get Material() As String

    If mobjMateriales Is Nothing Then CargarListaMateriales
    
    Material = mobjMateriales.Item(mudtprops.MaterialID)

End Property

Public Property Get Materiales() As TextList

    If mobjMateriales Is Nothing Then CargarListaMateriales
    
    Set Materiales = mobjMateriales

End Property

Public Property Get objMaterial(strMoneda As String) As Material
    Dim MyMaterial As Material
    
    If mudtprops.MaterialID = 0 Then Exit Property
    
    Set MyMaterial = New Material
        
    MyMaterial.Load mudtprops.MaterialID, strMoneda
    
    Set objMaterial = MyMaterial
    
    Set MyMaterial = Nothing
    
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

Private Sub CargarListaMateriales()

    Set mobjMateriales = New TextList
    mobjMateriales.Load "vNombreMateriales"

End Sub
  

Public Function CrearCodigoAutomatico(Nombre As String, TemporadaID As Long) As String
    Dim objSerie As Serie
    Dim strCodigo As String
        
    On Error GoTo ErrorManager
    
    strCodigo = Left(Nombre, 2)
    CrearCodigoAutomatico = strCodigo
    
    Do
        CrearCodigoAutomatico = strCodigo
        Set objSerie = New Serie
        
        On Error Resume Next
        objSerie.Load 0, strCodigo, TemporadaID, vbNullString
        On Error GoTo ErrorManager
        
        strCodigo = ObtenerSiguienteCodigo(strCodigo, 2)
    Loop Until objSerie.SerieID = 0
    
    Set objSerie = Nothing
    
    Exit Function

ErrorManager:
    CrearCodigoAutomatico = vbNullString
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Exit Function
End Function

Public Sub AsignarMaterialPredeterminado()
    Dim objParametro As ParametroAplicacion
    
    If Not mflgEditing Then Err.Raise 383
  
    ' Obtenemos el material predeterminado.
    Set objParametro = New ParametroAplicacion
    mudtprops.MaterialID = objParametro.MaterialPredeterminado
    Set objParametro = Nothing
        
    mobjValid.RuleBroken "MaterialID", (mudtprops.MaterialID = 0)
  
    mflgDirty = True

End Sub


Private Function GetState() As String
    Dim udtData As SerieData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As SerieData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
