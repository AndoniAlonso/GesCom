VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Esta clase es de tipo child puro, siempre es un
'child y tendra un unico padre
Event Valid(IsValid As Boolean)

Private mudtprops As DireccionProps
Private mudtSave As DireccionProps
Private mudtChild As DireccionProps

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChildEditing As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub ChildApplyEdit()

    If Not mflgChildEditing Then Err.Raise 445
  
    mflgChildEditing = False

End Sub

Public Sub ApplyEdit()

'    Dim objPersist As DireccionPersist
  
    If Not mflgEditing Then Err.Raise 445

'    Set objPersist = CreateObject("GCServerMTS.DireccionPersist", PERSIST_SERVER)
  
    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
'        objPersist.DeleteObject mudtProps.DireccionID
        mflgNew = True
        mflgDeleted = False

    ElseIf mflgDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        
        ' save object to database if appropriate
'        SetState objPersist.Save(GetState)
    
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False

    End If
    
'    Set objPersist = Nothing
    mflgEditing = False
    mflgDirty = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    
    Set mobjValid = New BrokenRules
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    ' En principio no hay ninguna regla rota, es decir el formato es libre
    
    mudtprops.Calle = vbNullString
    mudtprops.Poblacion = vbNullString
    mudtprops.CodigoPostal = vbNullString
    mudtprops.Provincia = vbNullString
    mudtprops.Pais = vbNullString
    mudtprops.Telefono1 = vbNullString
    mudtprops.Telefono2 = vbNullString
    mudtprops.Telefono3 = vbNullString
    mudtprops.Fax = vbNullString
    mudtprops.EMAIL = vbNullString
 
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

Public Sub Load(DireccionID As Long)

    Dim objPersist As DireccionPersist
  
    If mflgChildEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.DireccionPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(DireccionID)
    Set objPersist = Nothing
  
    mflgNew = False

End Sub

Public Sub Delete()

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

Public Sub ChildBeginEdit()

    If mflgChildEditing Then Err.Raise 445
  
    ' save object state
    LSet mudtChild = mudtprops
    mflgChildEditing = True

End Sub

Public Sub ChildCancelEdit()

    If Not mflgChildEditing Then Err.Raise 445
  
    mflgChildEditing = False
    mflgDeleted = False
    
    ' restore object state
    LSet mudtprops = mudtChild

End Sub

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
  
    ' restore object state
    LSet mudtprops = mudtSave

End Sub

Public Property Get DireccionID() As Long

    DireccionID = mudtprops.DireccionID

End Property

Public Property Get Calle() As String

    Calle = Trim(mudtprops.Calle)

End Property

Public Property Let Calle(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Calle) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    mobjValid.RuleBroken "Calle", (Len(Trim(Value)) = 0)
  
    mudtprops.Calle = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get Poblacion() As String

    Poblacion = Trim(mudtprops.Poblacion)

End Property

Public Property Let Poblacion(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    mudtprops.Poblacion = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get CodigoPostal() As String

    CodigoPostal = Trim(mudtprops.CodigoPostal)

End Property

Public Property Let CodigoPostal(Value As String)
    
    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.CodigoPostal) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - longitud nula   o
    '     - todos numericos
    mobjValid.RuleBroken "CodigoPostal", _
        Not ((Len(Trim(Value)) = 0) Or _
               (Len(Value) = Len(mudtprops.CodigoPostal) _
               And _
               IsNumeric(Value)) _
            )
  
    mudtprops.CodigoPostal = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get Provincia() As String

    Provincia = Trim(mudtprops.Provincia)

End Property

Public Property Let Provincia(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Provincia) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    mobjValid.RuleBroken "Provincia", (Len(Trim(Value)) = 0)
  
    mudtprops.Provincia = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get Pais() As String

    Pais = Trim(mudtprops.Pais)

End Property

Public Property Let Pais(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Pais) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
'    mobjValid.RuleBroken "Pais", (Len(Trim(Value)) = 0)
  
    mudtprops.Pais = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get Telefono1() As String

    Telefono1 = Trim(mudtprops.Telefono1)

End Property

Public Property Let Telefono1(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Telefono1) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    'mobjValid.RuleBroken "Telefono1", (Len(trim(Value)) = 0)
  
    mudtprops.Telefono1 = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get Telefono2() As String

    Telefono2 = Trim(mudtprops.Telefono2)

End Property

Public Property Let Telefono2(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Telefono2) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    'mobjValid.RuleBroken "Telefono2", (Len(trim(Value)) = 0)
  
    mudtprops.Telefono2 = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get Telefono3() As String

    Telefono3 = Trim(mudtprops.Telefono3)

End Property

Public Property Let Telefono3(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Telefono3) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    'mobjValid.RuleBroken "Telefono3", (Len(trim(Value)) = 0)
  
    mudtprops.Telefono3 = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get Fax() As String

    Fax = Trim(mudtprops.Fax)

End Property

Public Property Let Fax(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Fax) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    'mobjValid.RuleBroken "Fax", (Len(trim(Value)) = 0)
  
    mudtprops.Fax = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get EMAIL() As String

    EMAIL = Trim(mudtprops.EMAIL)

End Property

Public Property Let EMAIL(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.EMAIL) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    mobjValid.RuleBroken "EMAIL", (Len(Trim(Value)) = 0)
  
    mudtprops.EMAIL = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get DireccionText() As String
    
    Dim Texto As String
   
    Texto = vbNullString
    If Len(Trim(mudtprops.Calle)) <> 0 Then
        Texto = Trim(mudtprops.Calle)
    End If
   
    If Len(Trim(mudtprops.Poblacion)) <> 0 Then
        Texto = Texto & vbCrLf & Trim(mudtprops.Poblacion)
    End If
   
    If Len(Trim(mudtprops.CodigoPostal)) <> 0 Then
        Texto = Texto & vbCrLf & Trim(mudtprops.CodigoPostal)
    End If
    
    If Len(Trim(mudtprops.CodigoPostal)) <> 0 And Len(Trim(mudtprops.Provincia)) <> 0 Then
        Texto = Texto & "-"
    End If
    
    If Len(Trim(mudtprops.Provincia)) <> 0 Then
        Texto = Texto & " " & Trim(mudtprops.Provincia)
    End If
    
    If Len(Trim(mudtprops.Pais)) <> 0 Then
        Texto = Texto & vbCrLf & Trim(mudtprops.Pais)
    End If
   
    If Len(Trim(mudtprops.Telefono1)) <> 0 Then
        Texto = Texto & vbCrLf & "Teléfono:" & Trim(mudtprops.Telefono1)
    End If
    
    If Len(Trim(mudtprops.Telefono2)) <> 0 Then
        Texto = Texto & vbCrLf & Trim(mudtprops.Telefono2)
    End If
    
    If Len(Trim(mudtprops.Telefono3)) <> 0 Then
        Texto = Texto & vbCrLf & Trim(mudtprops.Telefono3)
    End If
    
    If Len(Trim(mudtprops.Fax)) <> 0 Then
        Texto = Texto & vbCrLf & "Fax:" & Trim(mudtprops.Fax)
    End If
   
    If Len(Trim(mudtprops.EMAIL)) <> 0 Then
        Texto = Texto & vbCrLf & "E-MAIL:" & Trim(mudtprops.EMAIL)
    End If
   
    DireccionText = Texto
    
End Property

Friend Function GetState() As String

    Dim udtData As DireccionData
  
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

    Dim udtData As DireccionData

    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False

End Sub
