VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "EstrModelosPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    
Implements ObjectControl

Private mobjContext As ObjectContext
  
Public Function Fetch(ByVal ModeloID As Long) As String
    Dim strSQL As String
    Dim rsEstrModelo As Recordset
    Dim objBuffer As Buffer
    Dim objPersist As EstrModeloPersist
    Dim udtProps As EstrModeloProps

    Set objBuffer = New Buffer

    On Error GoTo errh
    strSQL = "SELECT EstrModeloID FROM EstrModelos WHERE ModeloID=" & ModeloID
    Set rsEstrModelo = New Recordset
    rsEstrModelo.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.EstrModeloPersist")

    With objPersist
        objBuffer.Initialize Len(udtProps), 20

        Do While Not rsEstrModelo.EOF
            objBuffer.Add .Fetch(rsEstrModelo("EstrModeloID"))
            rsEstrModelo.MoveNext
        Loop

    End With

    Set objPersist = Nothing
    rsEstrModelo.Close
    Set rsEstrModelo = Nothing
    Fetch = objBuffer.GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Function Save(ByVal ModeloID As Long, ByRef Buffer As String) As String
    Dim objPersist As EstrModeloPersist
    Dim objBuffer As Buffer
    Dim objUpdatedEstrModelos As Buffer
    Dim lngIndex As Long
    Dim strBuffer As String

    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    Set objUpdatedEstrModelos = New Buffer
    objUpdatedEstrModelos.Initialize objBuffer.Length, objBuffer.Count
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.EstrModeloPersist")

    For lngIndex = 1 To objBuffer.Count
        With objPersist
            strBuffer = .SaveChild(ModeloID, objBuffer.Item(lngIndex))
            If Len(strBuffer) > 0 Then _
                objUpdatedEstrModelos.Add strBuffer
        End With

    Next

    Set objPersist = Nothing
    Save = objUpdatedEstrModelos.GetState
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

