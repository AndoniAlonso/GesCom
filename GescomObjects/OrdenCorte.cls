VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrdenCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As OrdenCorteProps
Private mudtSave As OrdenCorteProps

Private mMoneda As String

Private mobjOrdenCorteItems As OrdenCorteItems

' Aqui declaro el ComboBox de la lista de articulos,
Private mobjArticulos As TextList

Event Valid(IsValid As Boolean)

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit(Moneda As String)

    If mflgEditing Then Err.Raise 445
    
    ' save object state
    LSet mudtSave = mudtprops
    
    mMoneda = Moneda
    mobjOrdenCorteItems.BeginEdit Moneda
    
    mflgEditing = True
    
End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    ' restore object state
    LSet mudtprops = mudtSave
    mobjOrdenCorteItems.CancelEdit
    
End Sub

Public Sub ApplyEdit()
    Dim objPersist As OrdenCortePersist
    Dim strItems As String
  
    If Not mflgEditing Then Err.Raise 445
    
    Set objPersist = CreateObject("GCServerMTS.OrdenCortePersist", PERSIST_SERVER)
    
    mudtprops.ArticuloID = mobjOrdenCorteItems.ArticuloID
    
    ' Si la orden de corte se ha modificado y no tiene articulos, se borra.
    If mudtprops.ArticuloID = 0 Then mflgDeleted = True
        
    If mflgDeleted And Not mflgNew Then
        
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.OrdenCorteID, mobjOrdenCorteItems.GetState
        
        mflgNew = True
        mflgDeleted = False

    ElseIf mflgDirty Or mobjOrdenCorteItems.IsDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        
        strItems = mobjOrdenCorteItems.GetState
        SetState objPersist.Save(GetState, strItems)
        mobjOrdenCorteItems.SetState (strItems)
         
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False
    End If
    
    mobjOrdenCorteItems.ApplyEdit 'mudtprops.OrdenCorteID
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
    
    mudtprops.TemporadaID = 0
    mudtprops.EmpresaID = 0
    mudtprops.ArticuloID = 0
    mudtprops.Fecha = Date
    mudtprops.Observaciones = vbNullString
    mudtprops.Nombre = vbNullString
    
    mobjValid.RuleBroken "TemporadaID", True
    'mobjValid.RuleBroken "ArticuloID", True
    mobjValid.RuleBroken "EmpresaID", True
      
    Set mobjOrdenCorteItems = New OrdenCorteItems
    
End Sub

Private Sub Class_Terminate()

  Set mobjValid = Nothing
  
  Set mobjOrdenCorteItems = Nothing
  
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

Public Sub Load(OrdenCorteID As Long, Moneda As String)
    Dim objPersist As OrdenCortePersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.OrdenCortePersist", PERSIST_SERVER)
    SetState objPersist.Fetch(OrdenCorteID)
    Set objPersist = Nothing
    
    mobjValid.RuleBroken "TemporadaID", False
    'mobjValid.RuleBroken "ArticuloID", False
    mobjValid.RuleBroken "EmpresaID", False
    
    mMoneda = Moneda
    mobjOrdenCorteItems.Load OrdenCorteID, Moneda
    
    mflgNew = False
    
End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
    
    mflgDeleted = True
    mflgDirty = True
    
    mobjOrdenCorteItems.Delete

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

Public Property Get OrdenCorteID() As Long

    OrdenCorteID = mudtprops.OrdenCorteID
  
End Property

Public Property Get ArticuloID() As Long

    ArticuloID = mudtprops.ArticuloID
  
End Property

Public Property Let Articulo(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If mobjArticulos Is Nothing Then CargarListaArticulos
    
    mobjValid.RuleBroken "ArticuloID", (mobjArticulos.Key(Value) = 0)
    
    'Salimos si queremos asignar algo invalido
    If mobjArticulos.Key(Value) = 0 Then Exit Property
     
    ' Devolvemos error si ya hay hijos con otro articulo y queremos modificarlo.
    If mobjOrdenCorteItems.ArticuloID <> 0 Then _
        Err.Raise vbObjectError + 1001, "No se puede cambiar el articulo, ya hay lineas introducidas"
     
    mudtprops.ArticuloID = mobjArticulos.Key(Value)
  
    mflgDirty = True

End Property

Public Property Get Articulo() As String

    If mobjArticulos Is Nothing Then CargarListaArticulos
    
    Articulo = mobjArticulos.Item(mudtprops.ArticuloID)

End Property

Public Property Get Articulos() As TextList

    If mobjArticulos Is Nothing Then CargarListaArticulos
    
    Set Articulos = mobjArticulos

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

Public Property Let EmpresaID(Value As Long)

    If Not mflgNew Then Err.Raise 383
    
    mobjValid.RuleBroken "EmpresaID", (Value = 0)
    
    mudtprops.EmpresaID = Value
    mflgDirty = True

End Property

Public Property Get EmpresaID() As Long

    EmpresaID = mudtprops.EmpresaID
  
End Property


Public Property Get Moneda() As String

    Moneda = Trim(mMoneda)

End Property

Public Property Get Numero() As Long

    Numero = mudtprops.Numero
  
End Property

Public Property Let Numero(Value As Long)

    If Not mflgEditing Then Err.Raise 383
    
    If Value <= 0 Then _
        Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"
    
    mobjValid.RuleBroken "Numero", (Value <= 0)
    
    mudtprops.Numero = Value
    
    mflgDirty = True
    
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

Public Property Let FechaCorte(Value As Variant)

    If Not mflgEditing Then Err.Raise 383
    
    If Len(Trim(Value)) = 0 Then
        mudtprops.FechaCorte = vbNullString
        mobjValid.RuleBroken "FechaCorte", True
    
    ElseIf IsDate(Value) Then
        mudtprops.FechaCorte = CDate(Value)
        mobjValid.RuleBroken "FechaCorte", False
    
    Else
        mobjValid.RuleBroken "FechaCorte", True
    
    End If
    
    mflgDirty = True
    
End Property

Public Property Get FechaCorte() As Variant

    With mudtprops
      FechaCorte = IIf(.FechaCorte = 0, vbNullString, .FechaCorte)
    End With
    
End Property

Public Property Get OrdenCortada() As Boolean

    OrdenCortada = mobjOrdenCorteItems.OrdenCortada
    
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

Public Property Get Nombre() As String

    Nombre = Trim(mudtprops.Nombre)

End Property

Public Property Get OrdenCorteItems() As OrdenCorteItems

  Set OrdenCorteItems = mobjOrdenCorteItems

End Property

Public Sub Actualizar()

    If Not OrdenCortada Then mudtprops.FechaCorte = Date
    mobjOrdenCorteItems.Actualizar

End Sub

Public Sub Desactualizar()

    mobjOrdenCorteItems.Desactualizar

End Sub

Private Function GetState() As String

  Dim udtData As OrdenCorteData

  With mudtprops
    .IsDeleted = mflgDeleted
    .IsNew = mflgNew
    .IsDirty = mflgDirty
  End With

  LSet udtData = mudtprops
  GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

  Dim udtData As OrdenCorteData
  
  udtData.Buffer = Buffer
  LSet mudtprops = udtData

End Sub

Private Sub CargarListaArticulos()

    Set mobjArticulos = New TextList
    mobjArticulos.Load "Articulos", 0, mudtprops.TemporadaID
    'mudtProps.ArticuloID = mobjArticulos.Key(mobjArticulos.IndexItem(1))
  
End Sub

