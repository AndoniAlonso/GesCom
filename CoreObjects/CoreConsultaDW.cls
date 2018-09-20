VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoreConsultaDW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As CoreConsultaDWProps
Private mudtSave As CoreConsultaDWProps

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
    Dim objPersist As CoreConsultaDWPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("CoreServer.CoreConsultaDWPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.ConsultaID

        mflgNew = True
        mflgDeleted = False

    ElseIf IsDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        
        ' save object to database if appropriate
        SetState objPersist.Save(GetState)
        
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False
    Else
    
    End If
    
    Set objPersist = Nothing
    mflgDirty = False
    mflgEditing = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    
    Set mobjValid = New BrokenRules
 
         'TODO:Incuir aqui las reglas rotas
    'mobjValid.RuleBroken "Nombre", True
   
    mudtprops.Descripcion = vbNullString
    mudtprops.TiempoRefresco = 0
    mudtprops.DatePartRefresco = vbNullString
    mudtprops.VistaOrigen = vbNullString
    mudtprops.TablaDestino = vbNullString
    
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing

    'TODO: Liberar los objetos child
                    
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

Public Sub Load(ByVal ConsultaID As String)
    Dim objPersist As CoreConsultaDWPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("CoreServer.CoreConsultaDWPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(ConsultaID)
    Set objPersist = Nothing

    'mobjValid.RuleBroken "Nombre", False
    'TODO: Poner las reglas rotas correspondientes
         
         'TODO: Hacer el SetState de los objetos Child

    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True
  
    'TODO: Hacer delete de los objetos child

End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty
    'TODO: Tener en cuenta que si hay objetos child hay que marcar la propiedad Dirty teniendolos en cuenta

End Property


Public Property Get ConsultaID() As String
    ConsultaID = mudtprops.ConsultaID
End Property

Public Property Get Descripcion() As String
    Descripcion = mudtprops.Descripcion
End Property

Public Property Let Descripcion(ByVal Value As String)
    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.Descripcion) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    'TODO:Insert the Rule for the column here ..........
    mudtprops.Descripcion = Value
    
    mflgDirty = True
End Property

Public Property Get TiempoRefresco() As Long
    TiempoRefresco = mudtprops.TiempoRefresco
End Property

Public Property Let TiempoRefresco(ByVal Value As Long)
    If Not mflgEditing Then Err.Raise 383
    
    
    
    'TODO:Insert the Rule for the column here ..........
    mudtprops.TiempoRefresco = Value
    
    mflgDirty = True
End Property

Public Property Get DatePartRefresco() As String
    DatePartRefresco = mudtprops.DatePartRefresco
End Property

Public Property Let DatePartRefresco(ByVal Value As String)
    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.DatePartRefresco) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    'TODO:Insert the Rule for the column here ..........
    mudtprops.DatePartRefresco = Value
    
    mflgDirty = True
End Property

Public Property Get VistaOrigen() As String
    VistaOrigen = mudtprops.VistaOrigen
End Property

Public Property Let VistaOrigen(ByVal Value As String)
    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.VistaOrigen) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    'TODO:Insert the Rule for the column here ..........
    mudtprops.VistaOrigen = Value
    
    mflgDirty = True
End Property

Public Property Get TablaDestino() As String
    TablaDestino = mudtprops.TablaDestino
End Property

Public Property Let TablaDestino(ByVal Value As String)
    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.TablaDestino) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    'TODO:Insert the Rule for the column here ..........
    mudtprops.TablaDestino = Value
    
    mflgDirty = True
End Property

Private Function GetState() As String
    Dim udtData As CoreConsultaDWData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As CoreConsultaDWData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub

Public Sub UpdateTable()
    Dim objPersist As CoreConsultaDWPersist
  
    ' No puede ser una consulta nueva
    If mflgNew Then Err.Raise 383

    Set objPersist = CreateObject("CoreServer.CoreConsultaDWPersist", PERSIST_SERVER)

    objPersist.InsertTable mudtprops.VistaOrigen, mudtprops.TablaDestino, vbNullString
    
    Set objPersist = Nothing
End Sub
