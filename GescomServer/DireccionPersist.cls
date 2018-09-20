VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "DireccionPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Option Explicit
  
    Implements ObjectControl

    Private mobjContext As ObjectContext
  
  
  Private mudtProps As DireccionProps

Private Sub SetState(Buffer As String)

    Dim udtData As DireccionData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Private Function GetState() As String

    Dim udtData As DireccionData
  
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Public Function Fetch(ByVal DireccionID As Long) As String
Dim rsDireccion As Recordset
Dim strSQL As String
  

  On Error GoTo errh
    strSQL = "SELECT * " & _
        "FROM Direcciones " & _
        "WHERE DireccionID=" & DireccionID
    Set rsDireccion = New Recordset
    rsDireccion.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsDireccion
        mudtProps.DireccionID = .Fields("DireccionID")
        mudtProps.Calle = .Fields("Calle") & vbNullString
        mudtProps.Poblacion = .Fields("Poblacion") & vbNullString
        mudtProps.CodigoPostal = .Fields("CodigoPostal") & vbNullString
        mudtProps.Provincia = .Fields("Provincia") & vbNullString
        mudtProps.Pais = .Fields("Pais") & vbNullString
        mudtProps.Telefono1 = .Fields("Telefono1") & vbNullString
        mudtProps.Telefono2 = .Fields("Telefono2") & vbNullString
        mudtProps.Telefono3 = .Fields("Telefono3") & vbNullString
        mudtProps.Fax = .Fields("Fax") & vbNullString
        mudtProps.EMAIL = .Fields("EMAIL") & vbNullString
        rsDireccion.Close
    End With

    Set rsDireccion = Nothing
    Fetch = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Function Save(ByVal Buffer As String) As String

    Dim rsDireccion As Recordset
    Dim strSQL As String
  

  On Error GoTo errh
    SetState Buffer
    strSQL = "SELECT * FROM Direcciones WHERE DireccionID=" & mudtProps.DireccionID
    Set rsDireccion = New Recordset
    rsDireccion.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsDireccion.AddNew
  
    With rsDireccion
        .Fields("Calle") = mudtProps.Calle
        .Fields("Poblacion") = mudtProps.Poblacion
        .Fields("CodigoPostal") = mudtProps.CodigoPostal
        .Fields("Provincia") = mudtProps.Provincia
        .Fields("Pais") = mudtProps.Pais
        .Fields("Telefono1") = mudtProps.Telefono1
        .Fields("Telefono2") = mudtProps.Telefono2
        .Fields("Telefono3") = mudtProps.Telefono3
        .Fields("Fax") = mudtProps.Fax
        .Fields("EMAIL") = mudtProps.EMAIL
        
        .Update
    
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.DireccionID = .Fields("DireccionID")
        End If
    
        .Close
    End With

    Set rsDireccion = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal DireccionID As Long)
    Dim cnGescom As Connection


  On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Direcciones WHERE DireccionID=" & DireccionID

    cnGescom.Close
    Set cnGescom = Nothing
    mobjContext.SetComplete

    Exit Sub
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Public Function SaveChild(Buffer As String) As String


  On Error GoTo errh
    SetState Buffer

    With mudtProps
        If .IsDeleted Then
            DeleteObject .DireccionID
            SaveChild = vbNullString

        ElseIf .IsDirty Or .IsNew Then
            SaveChild = Save(GetState)

        Else
            SaveChild = Buffer
            mobjContext.SetComplete

        End If

    End With

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Private Function ObjectControl_CanBePooled() As Boolean

  ObjectControl_CanBePooled = False
  
End Function

Private Sub ObjectControl_Activate()

  Set mobjContext = GetObjectContext

End Sub

Private Sub ObjectControl_Deactivate()

  Set mobjContext = Nothing

End Sub


