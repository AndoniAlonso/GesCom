VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "PickerItemsPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjBuffer As Buffer

Public Function Fetch(ByVal strTabla As String, ByVal Propietario As Long, _
                      ByVal Empresa As Long, ByVal Temporada As Long) As String
    Dim rsPickerItem As Recordset
    Dim strSQL As String
    Dim strWhere As String
    Dim strAND As String
    Dim strOrder As String
    Dim udtData As PickerItemData
    Dim udtProps As PickerItemProps
    
    strWhere = vbNullString
    strOrder = " ORDER BY DocumentoID "
    strAND = vbNullString
    If Propietario Then
        strWhere = " PropietarioID = " & Propietario
        strAND = " AND "
    End If
    
    If Empresa Then
        strWhere = strWhere & strAND & " EmpresaID = " & Empresa
        strAND = " AND "
    End If
    
    If Temporada Then
        strWhere = strWhere & strAND & " TemporadaID = " & Temporada
    End If
    
    If Len(strWhere) Then strWhere = " WHERE " & strWhere
    
    strSQL = "SELECT * FROM " & strTabla & strWhere & strOrder
    
    Set rsPickerItem = New Recordset
    rsPickerItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    Set mobjBuffer = New Buffer
    mobjBuffer.Initialize Len(udtData.Buffer), 20
    
    Do While Not rsPickerItem.EOF
      With udtProps
        .DocumentoID = rsPickerItem("DocumentoID")
        .Nombre = rsPickerItem("Nombre") & vbNullString
        .Numero = IIf(IsNull(rsPickerItem("Numero")), 0, rsPickerItem("Numero"))
        .Descripcion = rsPickerItem("Descripcion") & vbNullString
        .Cantidad = rsPickerItem("Cantidad")
        .Fecha = IIf(IsNull(rsPickerItem("Fecha")), 0, rsPickerItem("Fecha"))
      
      End With
      LSet udtData = udtProps
      mobjBuffer.Add udtData.Buffer
      rsPickerItem.MoveNext
    Loop
    
    rsPickerItem.Close
    Set rsPickerItem = Nothing
    Fetch = mobjBuffer.GetState
  
End Function
