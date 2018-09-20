VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Prenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private mudtprops As PrendaProps
Private mudtSave As PrendaProps

Event Valid(IsValid As Boolean)

Private mMoneda As String

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit(Moneda As String)

    If mflgEditing Then Err.Raise 445
  
    mMoneda = Moneda
    
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
    Dim objPersist As PrendaPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.PrendaPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.PrendaID
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
    mobjValid.RuleBroken "Nombre", True
    mobjValid.RuleBroken "Codigo", True
  
    mudtprops.PlanchaPTA = 0
    mudtprops.PlanchaEUR = 0
    mudtprops.TransportePTA = 0
    mudtprops.TransporteEUR = 0
    mudtprops.PerchaPTA = 0
    mudtprops.PerchaEUR = 0
    mudtprops.CartonPTA = 0
    mudtprops.CartonEUR = 0
    mudtprops.EtiquetaPTA = 0
    mudtprops.EtiquetaEUR = 0
    mudtprops.Administracion = 0
      
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

Public Sub Load(PrendaID As Long, Moneda As String, Optional Codigo As String)

    Dim objPersist As PrendaPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    mMoneda = Moneda
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.PrendaPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(PrendaID, Codigo)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Nombre", False
    mobjValid.RuleBroken "Codigo", False
  
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

Public Property Get PrendaID() As Long

    PrendaID = mudtprops.PrendaID
  
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

Public Property Let Plancha(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    'If Value < 0 Then _
    '   Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "Plancha", (Value < 0)
    If EsEUR(mMoneda) Then
        mudtprops.PlanchaEUR = Round(Value, 2)
        mudtprops.PlanchaPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.PlanchaPTA = Round(Value, 0)
        mudtprops.PlanchaEUR = PTA2EUR(Value)
    Else
        mudtprops.PlanchaPTA = 0
        mudtprops.PlanchaEUR = 0
    End If
        
    mflgDirty = True

End Property

Public Property Get Plancha() As Double

    If EsEUR(mMoneda) Then
        Plancha = mudtprops.PlanchaEUR
    ElseIf EsPTA(mMoneda) Then
        Plancha = mudtprops.PlanchaPTA
    Else
        Plancha = 0
    End If
  
End Property

Public Property Let Transporte(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    'If Value < 0 Then _
    '   Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "transporte", (Value < 0)
      
    If EsEUR(mMoneda) Then
        mudtprops.TransporteEUR = Round(Value, 2)
        mudtprops.TransportePTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.TransportePTA = Round(Value, 0)
        mudtprops.TransporteEUR = PTA2EUR(Value)
    Else
        mudtprops.TransportePTA = 0
        mudtprops.TransporteEUR = 0
    End If
    
    mflgDirty = True

End Property

Public Property Get Transporte() As Double

    If EsEUR(mMoneda) Then
        Transporte = mudtprops.TransporteEUR
    ElseIf EsPTA(mMoneda) Then
        Transporte = mudtprops.TransportePTA
    Else
        Transporte = 0
    End If
  
End Property

Public Property Let Percha(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    'If Value < 0 Then _
    '   Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "percha", (Value < 0)
      
    If EsEUR(mMoneda) Then
        mudtprops.PerchaEUR = Round(Value, 2)
        mudtprops.PerchaPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.PerchaPTA = Round(Value, 0)
        mudtprops.PerchaEUR = PTA2EUR(Value)
    Else
        mudtprops.PerchaPTA = 0
        mudtprops.PerchaEUR = 0
    End If
    
    mflgDirty = True

End Property

Public Property Get Percha() As Double

    If EsEUR(mMoneda) Then
        Percha = mudtprops.PerchaEUR
    ElseIf EsPTA(mMoneda) Then
        Percha = mudtprops.PerchaPTA
    Else
        Percha = 0
    End If
  
End Property

Public Property Let Carton(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    'If Value < 0 Then _
    '   Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "Carton", (Value < 0)
      
    If EsEUR(mMoneda) Then
        mudtprops.CartonEUR = Round(Value, 2)
        mudtprops.CartonPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.CartonPTA = Round(Value, 0)
        mudtprops.CartonEUR = PTA2EUR(Value)
    Else
        mudtprops.CartonPTA = 0
        mudtprops.CartonEUR = 0
    End If
        
    mflgDirty = True

End Property

Public Property Get Carton() As Double

    If EsEUR(mMoneda) Then
        Carton = mudtprops.CartonEUR
    ElseIf EsPTA(mMoneda) Then
        Carton = mudtprops.CartonPTA
    Else
        Carton = 0
    End If
  
End Property

Public Property Let Etiqueta(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    'If Value < 0 Then _
    '   Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "Etiqueta", (Value < 0)
  
    If EsEUR(mMoneda) Then
        mudtprops.EtiquetaEUR = Round(Value, 2)
        mudtprops.EtiquetaPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.EtiquetaPTA = Round(Value, 0)
        mudtprops.EtiquetaEUR = PTA2EUR(Value)
    Else
        mudtprops.EtiquetaPTA = 0
        mudtprops.EtiquetaEUR = 0
    End If
    
    mflgDirty = True

End Property

Public Property Get Etiqueta() As Double

    If EsEUR(mMoneda) Then
        Etiqueta = mudtprops.EtiquetaEUR
    ElseIf EsPTA(mMoneda) Then
        Etiqueta = mudtprops.EtiquetaPTA
    Else
        Etiqueta = 0
    End If
      
End Property

Public Property Let Administracion(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    'If Value < 0 Then _
    '   Err.Raise vbObjectError + 1001, "Porcentaje no puede ser negativo"
  
    mobjValid.RuleBroken "Porcentaje", (Value < 0)
  
    mudtprops.Administracion = Round(Value, 3)
    
    mflgDirty = True

End Property

Public Property Get Administracion() As Double

    Administracion = mudtprops.Administracion
  
End Property

' Devuelve el precio de coste de la prenda
Public Property Get PrecioCoste() As Double

    With mudtprops
        If EsEUR(mMoneda) Then
            PrecioCoste = .PlanchaEUR + .TransporteEUR + .PerchaEUR + .CartonEUR + .EtiquetaEUR
        ElseIf EsPTA(mMoneda) Then
            PrecioCoste = .PlanchaPTA + .TransportePTA + .PerchaPTA + .CartonPTA + .EtiquetaPTA
        Else
            PrecioCoste = 0
        End If
    End With
    
End Property

Private Function GetState() As String

    Dim udtData As PrendaData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As PrendaData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
