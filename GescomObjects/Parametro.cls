VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As ParametroProps
Private mudtSave As ParametroProps

' Estas propiedades no se guardan en la BD.
Private mEmpresaActualID As Long
Private mTemporadaActualID As Long

' Aqui declaro el objeto 'child' Direccion
Private mobjDireccion As Direccion

' Aqui declaro los ComboBox de empresa y temporada
Private mobjEmpresas As TextList
Private mobjTemporadas As TextList

Event Valid(IsValid As Boolean)

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Private Const cnProyectoPersist = "GCServerMTS"

Public Sub BeginEdit()

    If mflgEditing Then Err.Raise 445
  
    ' save object state
    LSet mudtSave = mudtprops
  
    mobjDireccion.BeginEdit
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    
    ' restore object state
    LSet mudtprops = mudtSave
  
    mobjDireccion.CancelEdit
  
End Sub

Public Sub ApplyEdit()

    Dim objPersist As ParametroPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.ParametroPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.ParametroID, mudtprops.DireccionID
            
        ' tratar con los datos de direcciones
        mobjDireccion.Delete
        mobjDireccion.ApplyEdit
    
        mflgNew = True
        mflgDeleted = False

    ElseIf mflgDirty Or mflgNew Or mobjDireccion.IsDirty Then
        If Not IsValid Then Err.Raise 445
        mudtprops.Direccion.Buffer = mobjDireccion.GetState
    
        ' save object to database if appropriate
        SetState objPersist.Save(GetState)
        
        ' salvar el objeto hijo primero (en este caso)
        mobjDireccion.SetState (mudtprops.Direccion.Buffer)
        mobjDireccion.ApplyEdit
        mudtprops.DireccionID = mobjDireccion.DireccionID
    
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False
    Else
        mobjDireccion.ApplyEdit

    End If
    
    Set objPersist = Nothing
    mflgDirty = False
    mflgEditing = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    Set mobjValid = New BrokenRules
 
    Set mobjDireccion = New Direccion
  
    Set mobjEmpresas = New TextList
    mobjEmpresas.Load "Empresas"
    mudtprops.EmpresaID = mobjEmpresas.Key(mobjEmpresas.IndexItem(1))
  
    Set mobjTemporadas = New TextList
    mobjTemporadas.Load "Temporadas"
    mudtprops.TemporadaID = mobjTemporadas.Key(mobjTemporadas.IndexItem(1))
  
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    '
  
    mobjValid.RuleBroken "Clave", True
    mobjValid.RuleBroken "Propietario", True
    mobjValid.RuleBroken "Usuario", True
    mobjValid.RuleBroken "EmpresaID", True
    mobjValid.RuleBroken "EmpresaActualID", True
    mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "TemporadaActualID", True
    mobjValid.RuleBroken "Moneda", True
    'mobjValid.RuleBroken "DireccionID", True
  
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjDireccion = Nothing
    Set mobjEmpresas = Nothing
    Set mobjTemporadas = Nothing
  
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
    Dim objPersist As ParametroPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.ParametroPersist", PERSIST_SERVER)
    SetState objPersist.Fetch()
    Set objPersist = Nothing

    mobjValid.RuleBroken "Clave", False
    mobjValid.RuleBroken "Propietario", False
    mobjValid.RuleBroken "Usuario", False
    mobjValid.RuleBroken "EmpresaID", False
    mobjValid.RuleBroken "EmpresaActualID", False
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "TemporadaActualID", False
    mobjValid.RuleBroken "Moneda", False
    'mobjValid.RuleBroken "DireccionID", false
      
    ' Aprovecho para inicializar los valores de TemporadaActual y EmpresaActual
    mEmpresaActualID = mudtprops.EmpresaID
    mTemporadaActualID = mudtprops.TemporadaID
  
    mobjDireccion.SetState (mudtprops.Direccion.Buffer)
  
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True
    'mflgdirty?????
  
    mobjDireccion.Delete

End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty Or mobjDireccion.IsDirty

End Property

Public Property Get ParametroID() As Long

    ParametroID = mudtprops.ParametroID
  
End Property

Public Property Let Alfanumero(Value As String)
    
    Dim ReglaRota As Boolean

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Alfanumero) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    If ((Len(Trim(Value)) > 0) And (Len(Trim(Value)) < Len(mudtprops.Alfanumero))) Then
        ReglaRota = True
    Else
        ReglaRota = False
    End If

    ' Regla de longitud de Alfanumero si es mayor de cero -->
    mobjValid.RuleBroken "Alfanumero", ReglaRota

    mudtprops.Alfanumero = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Alfanumero() As String

    Alfanumero = Trim(mudtprops.Alfanumero)

End Property

Public Property Let Clave(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Clave) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    ' Por lo menos deberia ser de siete caracteres.
    mobjValid.RuleBroken "Clave", (Len(Trim(Value)) < 7)

    mudtprops.Clave = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Clave() As String

    Clave = Trim(mudtprops.Clave)

End Property

Public Property Let Propietario(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Propietario) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    mobjValid.RuleBroken "Propietario", (Len(Trim(Value)) = 0)

    mudtprops.Propietario = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Propietario() As String

    Propietario = Trim(mudtprops.Propietario)

End Property

Public Property Let Usuario(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Usuario) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    mobjValid.RuleBroken "Usuario", (Len(Trim(Value)) = 0)

    mudtprops.Usuario = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Usuario() As String

    Usuario = Trim(mudtprops.Usuario)

End Property

Public Property Let Moneda(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    'If Not EsMonedaValida(Value) Then _
    '   Err.Raise vbObjectError + 1002, "Moneda debe ser EUR ó PTA"
  
    mobjValid.RuleBroken "Moneda", (Not EsMonedaValida(Value))
    mudtprops.Moneda = UCase$(Value)
    mflgDirty = True

End Property

Public Property Get Moneda() As String

    Moneda = Trim(mudtprops.Moneda)

End Property

Public Property Get EmpresaID() As Long

    EmpresaID = mudtprops.EmpresaID
  
End Property

Public Property Let Empresa(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.EmpresaID = mobjEmpresas.Key(Value)
  
    mobjValid.RuleBroken "EmpresaID", (mudtprops.EmpresaID = 0)
  
    mflgDirty = True

End Property

Public Property Get Empresa() As String

    Empresa = mobjEmpresas.Item(mudtprops.EmpresaID)

End Property

Public Property Let EmpresaActual(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mEmpresaActualID = mobjEmpresas.Key(Value)
  
    mobjValid.RuleBroken "EmpresaActualID", (mEmpresaActualID = 0)
    mflgDirty = True

End Property

Public Property Get EmpresaActual() As String

    EmpresaActual = mobjEmpresas.Item(mEmpresaActualID)

End Property

Public Property Get Empresas() As TextList

    Set Empresas = mobjEmpresas

End Property

Public Property Get TemporadaID() As Long

    TemporadaID = mudtprops.TemporadaID
  
End Property

Public Property Let Temporada(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.TemporadaID = mobjTemporadas.Key(Value)
  
    mobjValid.RuleBroken "TemporadaID", (mudtprops.TemporadaID = 0)
    mflgDirty = True

End Property

Public Property Get Temporada() As String

    Temporada = mobjTemporadas.Item(mudtprops.TemporadaID)

End Property

Public Property Let TemporadaActual(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mTemporadaActualID = mobjTemporadas.Key(Value)
  
    mobjValid.RuleBroken "TemporadaActualID", (mTemporadaActualID = 0)
    mflgDirty = True

End Property

Public Property Get TemporadaActual() As String

    TemporadaActual = mobjTemporadas.Item(mTemporadaActualID)

End Property

Public Property Get Temporadas() As TextList

    Set Temporadas = mobjTemporadas

End Property

Public Property Let DireccionID(Value As Long)
  
    If Not mflgEditing Then Err.Raise 383
  
    'mobjValid.RuleBroken "DireccionID", (Value=0)
    mudtprops.DireccionID = Value
    mflgDirty = True
    
End Property

Public Property Get DireccionID() As Long

    DireccionID = mudtprops.DireccionID
  
End Property

Public Property Let EmpresaActualID(Value As Long)
  
    If Not mflgEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "EmpresaActualID", (Value = 0)
    mEmpresaActualID = Value
    mflgDirty = True
    
End Property

Public Property Get EmpresaActualID() As Long

    EmpresaActualID = mEmpresaActualID
  
End Property

Public Property Let TemporadaActualID(Value As Long)
  
    If Not mflgEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "TemporadaActualID", (Value = 0)
    mTemporadaActualID = Value
    mflgDirty = True

End Property

Public Property Get TemporadaActualID() As Long

    TemporadaActualID = mTemporadaActualID
  
End Property

Public Property Get AnchuraTelaEstandar() As Double

    AnchuraTelaEstandar = dblAnchuraTelaEstandar
  
End Property

Public Property Get Direccion() As Direccion

    Set Direccion = mobjDireccion
  
End Property

Public Property Get ObjEmpresaActual() As Empresa
    Dim auxEmpresa As Empresa
  
    Set auxEmpresa = New Empresa
    auxEmpresa.Load mEmpresaActualID
    'mudtprops.EmpresaID
    Set ObjEmpresaActual = auxEmpresa
    Set auxEmpresa = Nothing

End Property

Public Property Let ServidorContawin(Value As String)
  
    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.ServidorContawin) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.ServidorContawin = Trim(Value)
    mflgDirty = True

End Property

Public Property Get ServidorContawin() As String

    ServidorContawin = Trim(mudtprops.ServidorContawin)

End Property

Public Property Let Proyecto(Value As String)
  
    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Proyecto) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.Proyecto = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Proyecto() As String

    'Proyecto = Trim(mudtprops.Proyecto)
    'Tomamos el valor constante "GCServerMTS"
    Proyecto = cnProyectoPersist

End Property

Public Property Let ServidorPersist(Value As String)
  
    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.ServidorPersist) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.ServidorPersist = Trim(Value)
    mflgDirty = True

End Property

Public Property Get ServidorPersist() As String

    'ServidorPersist = Trim(mudtprops.ServidorPersist)
    'Tomamos el valor del servidor PERSIST que tenemos ya obtenido
    ServidorPersist = PERSIST_SERVER

End Property

Public Property Let Sufijo(Value As String)
  
    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Sufijo) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.Sufijo = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Sufijo() As String

    Sufijo = Trim(mudtprops.Sufijo)

End Property

Private Function GetState() As String

    Dim udtData As ParametroData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As ParametroData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
