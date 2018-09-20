VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ParametroAplicacionPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
    
Private mudtProps As ParametroAplicacionProps

Private Function GetState() As String

  Dim udtData As ParametroAplicacionData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

  Dim udtData As ParametroAplicacionData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String

  Dim rsParametroAplicacion As Recordset
  Dim strSQL As String

    On Error GoTo errh
  SetState Buffer

  Set rsParametroAplicacion = New Recordset
  strSQL = "SELECT * FROM ParametrosAplicacion WHERE ParametroAplicacionID='" & _
    mudtProps.ParametroAplicacionID & "'"
  rsParametroAplicacion.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
  If mudtProps.IsNew Then rsParametroAplicacion.AddNew

  With rsParametroAplicacion
    .Fields("Nombre") = mudtProps.Nombre
    .Fields("Valor") = mudtProps.Valor
    .Fields("Sistema") = mudtProps.Sistema
    .Fields("TipoParametro") = mudtProps.TipoParametro

    .Update

    If mudtProps.IsNew Then
       .Bookmark = .Bookmark
       mudtProps.ParametroAplicacionID = .Fields("ParametroAplicacionID")
    End If
    .Close
  End With

  Set rsParametroAplicacion = Nothing
  Save = GetState
      mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal ParametroAplicacionID As String)
    Dim cnGescom As Connection


  On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

  cnGescom.Execute "DELETE FROM ParametrosAplicacion WHERE ParametroAplicacionID='" & _
    ParametroAplicacionID & "'"
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

Public Function Fetch(ByVal ParametroAplicacionID As String) As String

  Dim rsParametroAplicacion As Recordset
  Dim strSQL As String

  On Error GoTo errh
  strSQL = "SELECT * FROM ParametrosAplicacion WHERE ParametroAplicacionID='" & _
    ParametroAplicacionID & "'"
  Set rsParametroAplicacion = New Recordset
  rsParametroAplicacion.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
  With rsParametroAplicacion
    mudtProps.ParametroAplicacionID = .Fields("ParametroAplicacionID") & ""
    mudtProps.Nombre = .Fields("Nombre")
    mudtProps.Valor = .Fields("Valor")
    mudtProps.Sistema = .Fields("Sistema")
    mudtProps.TipoParametro = .Fields("TipoParametro")

    .Close
  End With

  Set rsParametroAplicacion = Nothing
  Fetch = GetState
    mobjContext.SetComplete

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

