VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Etiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolEtiquetas As Collection
Private mstrFileName As String
Private mflgEditing As Boolean
Private mobjTallaje As Tallaje

Private Sub Class_Initialize()

    Set mcolEtiquetas = New Collection
    mstrFileName = vbNullString

End Sub

Private Sub Class_Terminate()

    Set mcolEtiquetas = Nothing
    
    Set mobjTallaje = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As Etiqueta
Attribute Item.VB_UserMemId = 0

    Set Item = mcolEtiquetas(Index)

End Function

Public Function Count() As Long

    Count = mcolEtiquetas.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolEtiquetas.[_NewEnum]

End Function

Public Sub BeginEdit()
    Dim objEtiqueta As Etiqueta
    
    If mflgEditing Then Err.Raise 445
    
    For Each objEtiqueta In mcolEtiquetas
      objEtiqueta.ChildBeginEdit
    Next
    
    mflgEditing = True
    
End Sub

Public Sub CancelEdit()
    Dim objEtiqueta As Etiqueta
    Dim lngIndex As Long
    
    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    For lngIndex = mcolEtiquetas.Count To 1 Step -1
      Set objEtiqueta = mcolEtiquetas.Item(lngIndex)
      With objEtiqueta
        If .IsNew Then _
           mcolEtiquetas.Remove lngIndex
        .ChildCancelEdit
      End With
    
    Next lngIndex
    
End Sub

Public Sub ApplyEdit()
    Dim objEtiqueta As Etiqueta

    If Not mflgEditing Then Err.Raise 445
    ' code to delete the object's data goes here
    
    For Each objEtiqueta In mcolEtiquetas
        objEtiqueta.ChildApplyEdit
    Next
    
    mflgEditing = False
    
End Sub

Public Sub Delete()

    Dim objEtiqueta As Etiqueta
    
    If Not mflgEditing Then Err.Raise 445
    ' code to delete the object's data goes here
    
    For Each objEtiqueta In mcolEtiquetas
        With objEtiqueta
            .BeginEdit
            .Delete
            .ApplyEdit
        End With
    Next
    
End Sub

Public Function Add() As Etiqueta

    Dim objEtiqueta As Etiqueta
    
    If Not mflgEditing Then Err.Raise 445
    
    Set objEtiqueta = New Etiqueta
    objEtiqueta.SetAsChild Me
    objEtiqueta.ChildBeginEdit
    Set Add = objEtiqueta
    Set objEtiqueta = Nothing
    
End Function

Friend Sub AddEtiqueta(Child As Etiqueta)

    mcolEtiquetas.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445
    
    With mcolEtiquetas(Index)
      .BeginEdit
      .Delete
      .ApplyEdit
    End With

End Sub

Public Property Get Cantidad() As Double
    Dim objEtiqueta As Etiqueta
    
    Cantidad = 0
    For Each objEtiqueta In mcolEtiquetas
        Cantidad = Cantidad + objEtiqueta.Cantidad
    Next
    
End Property

Friend Property Get IsDirty() As Boolean
    Dim objItem As Etiqueta
    
    IsDirty = False
    
    For Each objItem In mcolEtiquetas
        If objItem.IsDirty Or objItem.IsNew Then
            IsDirty = True
            Exit For
        End If
    Next
    
End Property

Public Property Let FileName(Value As String)

    If mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) = 0 Then _
        Err.Raise vbObjectError + 1001, "El nombre del fichero no puede ser vacío"

    mstrFileName = Trim(Value)

End Property

Public Sub WriteSequentialFile()
    Dim FileNum As Integer 'Windows File Number Reference
    Dim objItem As Etiqueta
    Dim strLineToWrite As String
    Dim i As Integer
    Dim j As Integer
    Dim nCantidad As Integer

    FileNum = FreeFile ' Get a free file number
    
    Open mstrFileName For Output As FileNum
    For Each objItem In mcolEtiquetas
        For i = 1 To 11
            nCantidad = objItem.CantidadTalla(i)
            For j = 1 To nCantidad
                With objItem
                strLineToWrite = IIf(Val(.PorcComposicion1), Trim(.PorcComposicion1) & "% ", vbNullString) & Trim(.Composicion1) & " ," & _
                    IIf(Val(.PorcComposicion2), Trim(.PorcComposicion2) & "% ", vbNullString) & Trim(.Composicion2) & " ," & _
                    IIf(Val(.PorcComposicion3), Trim(.PorcComposicion3) & "% ", vbNullString) & Trim(.Composicion3) & " ," & _
                    ObtenerTalla(i, objItem.TallajeID) & "," & _
                    Trim(.NombreColor) & " ," & _
                    Trim(.NombreModelo) & " ," & _
                    Trim(.NombreSerie) & " ," & _
                    CStr(34 + 2 * i) & CStr(.ArticuloColorID) & "," & _
                    Trim(Replace(FormatoMoneda(objItem.PrecioVentaPublico, "EUR", False), ",", ".")) & "," & _
                    Trim(.NombrePrenda) & " ," & _
                    Trim(.CodigoProveedor)
                End With

                Print #FileNum, strLineToWrite
            Next 'j
        Next 'i
    Next
    
    Close FileNum

End Sub

Private Function ObtenerTalla(i As Integer, TallajeID As Long) As String

    If mobjTallaje Is Nothing Then Set mobjTallaje = New Tallaje
    
    If mobjTallaje.TallajeID <> TallajeID Then
        Set mobjTallaje = Nothing
        Set mobjTallaje = New Tallaje
        mobjTallaje.Load TallajeID
    End If
    
    Select Case i
    Case 1
        ObtenerTalla = mobjTallaje.DescripcionT36
    Case 2
        ObtenerTalla = mobjTallaje.DescripcionT38
    Case 3
        ObtenerTalla = mobjTallaje.DescripcionT40
    Case 4
        ObtenerTalla = mobjTallaje.DescripcionT42
    Case 5
        ObtenerTalla = mobjTallaje.DescripcionT44
    Case 6
        ObtenerTalla = mobjTallaje.DescripcionT46
    Case 7
        ObtenerTalla = mobjTallaje.DescripcionT48
    Case 8
        ObtenerTalla = mobjTallaje.DescripcionT50
    Case 9
        ObtenerTalla = mobjTallaje.DescripcionT52
    Case 10
        ObtenerTalla = mobjTallaje.DescripcionT54
    Case 11
        ObtenerTalla = mobjTallaje.DescripcionT56
    Case Else
        Err.Raise vbObjectError + 1001, "Etiquetas[ObtenerTalla]", "Número de talla incorrecto"
    End Select
    
    ObtenerTalla = Trim(ObtenerTalla)

End Function
