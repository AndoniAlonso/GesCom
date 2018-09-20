VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParametroAplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As ParametroAplicacionProps
Private mudtSave As ParametroAplicacionProps
Private Const ALMAPRED As String = "ALMAPRED"
Private Const ALMATRAN As String = "ALMATRAN"  'ALMACEN EN TRANSITO
Private Const TIPOINST As String = "TIPOINST"
Private Const MATEPRED As String = "MATEPRED"
Private Const CLIEPRED As String = "CLIEPRED"
Private Const TALLPRED As String = "TALLPRED"


Event Valid(IsValid As Boolean)

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Enum enTipoInstalacion
  TIPOINSTALACION_FABRICA = 0
  TIPOINSTALACION_PUNTOVENTA = 1
End Enum

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

    Dim objPersist As ParametroAplicacionPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.ParametroAplicacionPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.ParametroAplicacionID
            
        ' tratar con los datos de direcciones
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
  
    mobjValid.RuleBroken "Valor", True
    mobjValid.RuleBroken "Nombre", True
    mobjValid.RuleBroken "Sistema", True
    mobjValid.RuleBroken "TipoParametro", True
  
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

Public Sub Load(ParametroAplicacionID As String)
    Dim objPersist As ParametroAplicacionPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.ParametroAplicacionPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(ParametroAplicacionID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Valor", False
    mobjValid.RuleBroken "Nombre", False
    mobjValid.RuleBroken "Sistema", False
    mobjValid.RuleBroken "TipoParametro", False
      
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True
    'mflgdirty?????
  
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

Public Property Get ParametroAplicacionID() As Long

    ParametroAplicacionID = mudtprops.ParametroAplicacionID
  
End Property

Public Property Let Nombre(Value As String)
    
    Dim ReglaRota As Boolean

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Nombre) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    If ((Len(Trim(Value)) > 0) And (Len(Trim(Value)) < Len(mudtprops.Nombre))) Then
        ReglaRota = True
    Else
        ReglaRota = False
    End If

    ' Regla de longitud de Nombre si es mayor de cero -->
    mobjValid.RuleBroken "Nombre", ReglaRota

    mudtprops.Nombre = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Nombre() As String

    Nombre = Trim(mudtprops.Nombre)

End Property

Public Property Let Valor(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Valor) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    mudtprops.Valor = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Valor() As String

    Valor = Trim(mudtprops.Valor)

End Property

Public Property Let Sistema(Value As Boolean)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.Sistema = Value
    mflgDirty = True

End Property

Public Property Get Sistema() As Boolean

    Sistema = mudtprops.Sistema

End Property

Public Property Let TipoParametro(Value As Long)

    If Not mflgEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "TipoParametro", (Value < 0)

    mudtprops.TipoParametro = Value
    mflgDirty = True

End Property

Public Property Get TipoParametro() As Long

    TipoParametro = mudtprops.TipoParametro

End Property

Public Property Get AlmacenPredeterminado() As Long
    Dim objParametroAplicacion As ParametroAplicacion
    
    Set objParametroAplicacion = New ParametroAplicacion
    objParametroAplicacion.Load (ALMAPRED)
    AlmacenPredeterminado = CLng(objParametroAplicacion.Valor)
    Set objParametroAplicacion = Nothing
    
End Property

Public Property Get TipoInstalacion() As enTipoInstalacion
    Dim objParametroAplicacion As ParametroAplicacion
    
    Set objParametroAplicacion = New ParametroAplicacion
    objParametroAplicacion.Load (TIPOINST)
    TipoInstalacion = CLng(objParametroAplicacion.Valor)
    Set objParametroAplicacion = Nothing
    
End Property

Public Property Get MaterialPredeterminado() As Long
    Dim objParametroAplicacion As ParametroAplicacion
    
    Set objParametroAplicacion = New ParametroAplicacion
    objParametroAplicacion.Load (MATEPRED)
    MaterialPredeterminado = CLng(objParametroAplicacion.Valor)
    Set objParametroAplicacion = Nothing
    
End Property

Public Property Get ClientePredeterminado() As Long
    Dim objParametroAplicacion As ParametroAplicacion
    
    Set objParametroAplicacion = New ParametroAplicacion
    objParametroAplicacion.Load (CLIEPRED)
    ClientePredeterminado = CLng(objParametroAplicacion.Valor)
    Set objParametroAplicacion = Nothing
    
End Property

Public Property Get AlmacenTransito() As Long
    Dim objParametroAplicacion As ParametroAplicacion
    
    Set objParametroAplicacion = New ParametroAplicacion
    objParametroAplicacion.Load (ALMATRAN)
    AlmacenTransito = CLng(objParametroAplicacion.Valor)
    Set objParametroAplicacion = Nothing
    
End Property

Public Property Get TallajePredeterminado() As Long
    Dim objParametroAplicacion As ParametroAplicacion
    
    Set objParametroAplicacion = New ParametroAplicacion
    objParametroAplicacion.Load (TALLPRED)
    TallajePredeterminado = CLng(objParametroAplicacion.Valor)
    Set objParametroAplicacion = Nothing
    
End Property

Private Function GetState() As String

    Dim udtData As ParametroAplicacionData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As ParametroAplicacionData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
