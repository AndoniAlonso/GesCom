VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CoreConsultaDWPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtprops As CoreConsultaDWProps

Private Function GetState() As String

  Dim udtData As CoreConsultaDWData

  LSet udtData = mudtprops
  GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

  Dim udtData As CoreConsultaDWData

  udtData.Buffer = Buffer
  LSet mudtprops = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String

  Dim rsConsulta As Recordset
  Dim strSQL As String

  SetState Buffer

  Set rsConsulta = New Recordset
  strSQL = "SELECT * FROM tbCoreConsultaDW WHERE ConsultaID='" & _
    mudtprops.ConsultaID & "'"
  rsConsulta.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
  If mudtprops.IsNew Then rsConsulta.AddNew

  With rsConsulta
    .Fields("Descripcion") = mudtprops.Descripcion
    .Fields("TiempoRefresco") = mudtprops.TiempoRefresco
    .Fields("DatePartRefresco") = mudtprops.DatePartRefresco
    .Fields("VistaOrigen") = mudtprops.VistaOrigen
    .Fields("TablaDestino") = mudtprops.TablaDestino

    .Update

    If mudtprops.IsNew Then
       .Bookmark = .Bookmark
       mudtprops.ConsultaID = .Fields("ConsultaID")
    End If
    .Close
  End With

  Set rsConsulta = Nothing
  Save = GetState

End Function

Public Sub DeleteObject(ByVal ConsultaID As String)
    Dim cnGescom As Connection

    On Error GoTo errh
    Set cnGescom = New Connection

    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
  
    cnGescom.Execute "DELETE FROM tbCoreConsultaDW WHERE ConsultaID='" & _
      ConsultaID & "'"
    
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

Public Function Fetch(ByVal ConsultaID As String) As String

  Dim rsConsulta As Recordset
  Dim strSQL As String

  strSQL = "SELECT * FROM tbCoreConsultaDW WHERE ConsultaID='" & _
    ConsultaID & "'"
  Set rsConsulta = New Recordset
  rsConsulta.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

  With rsConsulta
    mudtprops.ConsultaID = .Fields("ConsultaID") & ""
    mudtprops.Descripcion = .Fields("Descripcion") & ""
    mudtprops.TiempoRefresco = .Fields("TiempoRefresco")
    mudtprops.DatePartRefresco = .Fields("DatePartRefresco") & ""
    mudtprops.VistaOrigen = .Fields("VistaOrigen") & ""
    mudtprops.TablaDestino = .Fields("TablaDestino") & ""

    .Close
  End With

  Set rsConsulta = Nothing
  Fetch = GetState

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


Public Sub InsertTable(ByVal strOrigen As String, ByVal strDestino As String, ByVal strWhere As String)
    Dim cnGescom As Connection

    On Error GoTo errh
    Set cnGescom = New Connection

    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    cnGescom.Execute "DELETE FROM " & strDestino
    
    cnGescom.Execute "INSERT INTO " & strDestino & " SELECT * FROM " & strOrigen
    
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


