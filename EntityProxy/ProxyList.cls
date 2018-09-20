VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProxyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolList As Collection
Private mcolKeys As Collection
Private mstrCampoClave As String
Private mstrCampoDescripcion As String

Private Sub Class_Initialize()

    Set mcolList = New Collection
    Set mcolKeys = New Collection
    mstrCampoClave = vbNullString
    mstrCampoDescripcion = vbNullString
    
End Sub

Private Sub Class_Terminate()

    Set mcolList = Nothing
    Set mcolKeys = Nothing
    
End Sub

Public Function Item(ByVal Text As String) As String
Attribute Item.VB_UserMemId = 0

    Item = mcolList.Item(Text)

End Function

Public Function IndexItem(ByVal Index As Variant) As String

    IndexItem = mcolList.Item(Index)

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolList.[_NewEnum]

End Function

Public Function Key(ByVal Text As String) As String

    Key = mcolKeys.Item(Text)

End Function

Public Property Get CampoClave() As String
    CampoClave = mstrCampoClave
End Property

Public Property Get CampoDescripcion() As String
    CampoDescripcion = mstrCampoDescripcion
End Property

Friend Sub Load(ListName As String, CampoClave As String, CampoDescripcion As String, _
                ClausulaWhere As String, Proyecto As String, ServidorPersist As String, Optional IsAlphanumericList As Boolean = False)
Dim objRecordList As RecordList
Dim rsRecordList As ADOR.Recordset
    
    Set objRecordList = New RecordList
    Set rsRecordList = objRecordList.Load("Select " & CampoClave & ", " & CampoDescripcion & " from " & ListName, ClausulaWhere, Proyecto, ServidorPersist)
    mstrCampoDescripcion = CampoDescripcion
    mstrCampoClave = CampoClave
    rsRecordList.Sort = mstrCampoDescripcion
    mcolList.Add Format$("(Seleccionar uno)"), Format$(0)
    mcolKeys.Add Format$(0), Format$("(Seleccionar uno)")
    While Not rsRecordList.EOF
        If Not IsAlphanumericList Then
            mcolList.Add Trim(Format$(rsRecordList.Fields(1).Value)), Format$(rsRecordList.Fields(0).Value)
            mcolKeys.Add Format$(rsRecordList.Fields(0).Value), Trim(Format$(rsRecordList.Fields(1).Value))
        Else
            mcolList.Add Trim(rsRecordList.Fields(1).Value), rsRecordList.Fields(0).Value
            mcolKeys.Add rsRecordList.Fields(0).Value, Trim(rsRecordList.Fields(1).Value)
        End If
'        ListAdd mcolList, rsRecordList.Fields(1).Value, rsRecordList.Fields(0).Value, False
'        ListAdd mcolKeys, rsRecordList.Fields(0).Value, rsRecordList.Fields(1).Value, True
        rsRecordList.MoveNext
    Wend
    rsRecordList.Close
    Set rsRecordList = Nothing
    Set objRecordList = Nothing
End Sub
    
'Private Sub ListAdd(ByRef colList As Collection, ByVal Value1 As Variant, ByVal Value2 As Variant, IsKeyList As Boolean)
'
'    On Error GoTo ErrorManager
'
'    If Not IsKeyList Then
'        colList.Add Trim(Format$(Value1)), Format$(Value2)
'    Else
'        colList.Add Format$(Value1), Trim(Format$(Value2))
'    End If
'
'
'    Exit Sub
'ErrorManager:
'    If Err.Number = 457 Then
'        If Not IsKeyList Then
'            colList.Add Trim(Value1), Value2
'        Else
'            colList.Add Value1, Value2
'        End If
'
'        Resume Next
'    End If
'
'End Sub
'
