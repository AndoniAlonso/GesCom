VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ArticuloColorList 
   Caption         =   "Lista de Artículos - Color"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "ArticuloColorList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwItems 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbHerramientas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "ArticuloColorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmArticuloColor As ArticuloColorEdit
Private objArticuloColor As ArticuloColor
Private mobjBusqueda As Consulta
Public SentenciaSQL As String

Public Sub ComponentStatus(rsStatus As ADOR.Recordset)
   
    Set mrsRecordList = rsStatus
    Call RefreshListView

End Sub

Private Sub RefreshListView()
    Dim itmList As ListItem

    While Not mrsRecordList.EOF
        Set itmList = _
            lvwItems.ListItems.Add(Key:= _
            Format$(mrsRecordList("ArticuloColorID")) & " K")

        With itmList
            .Text = Trim(mrsRecordList("Nombre"))
            .SubItems(1) = Trim(mrsRecordList("NombreColor"))
            .SubItems(2) = FormatoCantidad(mrsRecordList("StockActualT36") + _
                                           mrsRecordList("StockActualT38") + _
                                           mrsRecordList("StockActualT40") + _
                                           mrsRecordList("StockActualT42") + _
                                           mrsRecordList("StockActualT44") + _
                                           mrsRecordList("StockActualT46") + _
                                           mrsRecordList("StockActualT48") + _
                                           mrsRecordList("StockActualT50") + _
                                           mrsRecordList("StockActualT52") + _
                                           mrsRecordList("StockActualT54") + _
                                           mrsRecordList("StockActualT56"), True)
            .SubItems(3) = FormatoCantidad(mrsRecordList("StockPendienteT36") + _
                                           mrsRecordList("StockPendienteT38") + _
                                           mrsRecordList("StockPendienteT40") + _
                                           mrsRecordList("StockPendienteT42") + _
                                           mrsRecordList("StockPendienteT44") + _
                                           mrsRecordList("StockPendienteT46") + _
                                           mrsRecordList("StockPendienteT48") + _
                                           mrsRecordList("StockPendienteT50") + _
                                           mrsRecordList("StockPendienteT52") + _
                                           mrsRecordList("StockPendienteT54") + _
                                           mrsRecordList("StockPendienteT56"), True)
            
            .Icon = GescomMain.mglIconosGrandes.ListImages("ArticuloColor").Key
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("ArticuloColor").Key
        End With

        mrsRecordList.MoveNext
    Wend

End Sub

Private Sub Form_Load()

    Me.Move 0, 0
    With lvwItems
        .ColumnHeaders.Add , , "Artículo - Color", ColumnSize(20)
        .ColumnHeaders.Add , , "Nombre Color", ColumnSize(10)
        .ColumnHeaders.Add , , "Stock Actual Total", ColumnSize(10), vbRightJustify
        .ColumnHeaders.Add , , "Stock Pendiente Total", ColumnSize(10), vbRightJustify
        
        .Icons = GescomMain.mglIconosGrandes
        .SmallIcons = GescomMain.mglIconosPequeños
    End With
    
    LoadImages Me.tlbHerramientas
    
    Set mobjBusqueda = New Consulta
     
    mlngColumn = 1
    
End Sub

''''???? para liberar memoria ?
Private Sub Form_Unload(Cancel As Integer)
    mrsRecordList.Close
    Set mrsRecordList = Nothing
End Sub

Private Sub lvwItems_DblClick()
    
    Call EditSelected
    
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        EditSelected
    ElseIf KeyCode = 46 Then
        DeleteSelected
    End If
        
End Sub

Private Sub lvwItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        Me.PopupMenu GescomMain.mnuListView
        lvwItems.Enabled = False
        lvwItems.Enabled = True
    End If
    
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    
    ListView_ColumnClick lvwItems, ColumnHeader
    mlngColumn = ColumnHeader.Index
       
End Sub

Public Sub DeleteSelected()

    On Error GoTo ErrorManager
   
    Dim Respuesta As VbMsgBoxResult
    Dim i As Integer

    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos borrar
    Respuesta = MostrarMensaje(MSG_DELETE)
    
    If Respuesta = vbYes Then
        For i = lvwItems.ListItems.Count To 1 Step -1
            If lvwItems.ListItems(i).Selected = True Then
                mlngID = Val(lvwItems.ListItems(i).Key)
                If mlngID > 0 Then
                    Set objArticuloColor = New ArticuloColor
                    objArticuloColor.Load mlngID, GescomMain.objParametro.Moneda
                    objArticuloColor.BeginEdit
                    objArticuloColor.Delete
                    objArticuloColor.ApplyEdit
                    Set objArticuloColor = Nothing
                    lvwItems.ListItems.Remove (i)
                End If
            End If
        Next i
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub UpdateListView(Optional strWhere As String)
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    lvwItems.ListItems.Clear

    Set objRecordList = New RecordList
    '''???? para liberar memoria
    mrsRecordList.Close
    Set mrsRecordList = Nothing
    Set mrsRecordList = objRecordList.Load("Select * from ArticuloColores", _
                            "TemporadaID = " & GescomMain.objParametro.TemporadaActualID & _
                            IIf(strWhere = vbNullString, vbNullString, " AND " & strWhere))
    Set objRecordList = Nothing
    
    Call RefreshListView
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub EditSelected()
    
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos abrirlos todos
    ' si el número es mayor que 5
    If NumeroSeleccionados(lvwItems) >= 5 Then
        Respuesta = MostrarMensaje(MSG_OPEN)
        
        If Respuesta = vbYes Then
            EditItems
        End If
    Else
        EditItems
    End If

    Exit Sub
    
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub EditItems()

    Dim i As Integer

    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Selected = True Then
            mlngID = Val(lvwItems.ListItems(i).Key)
            If mlngID > 0 Then
                Set frmArticuloColor = New ArticuloColorEdit
                Set objArticuloColor = New ArticuloColor
                objArticuloColor.Load mlngID, GescomMain.objParametro.Moneda
                frmArticuloColor.Component objArticuloColor
                frmArticuloColor.Show
                Set objArticuloColor = Nothing
            End If
        End If
    Next i
    
End Sub

Public Sub NewObject()

    Set frmArticuloColor = New ArticuloColorEdit
    Set objArticuloColor = New ArticuloColor
    frmArticuloColor.Component objArticuloColor
    frmArticuloColor.Show
    Set objArticuloColor = Nothing

End Sub

Public Sub SetListViewStyle(View As Integer)
   
    lvwItems.View = View
   
End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = True
    
End Function

Private Sub tlbHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case Is = "Nuevo"
            NewObject
        Case Is = "Abrir"
            EditSelected
        Case Is = "Imprimir"
            Imprimir
        Case Is = "Eliminar"
            DeleteSelected
        Case Is = "Actualizar"
            UpdateListView SentenciaSQL
        Case Is = "Buscar"
            ResultSearch
        Case Is = "IconosGrandes"
            SetListViewStyle (lvwIcon)
        Case Is = "IconosPequeños"
            SetListViewStyle (lvwSmallIcon)
        Case Is = "Lista"
            SetListViewStyle (lvwList)
        Case Is = "Detalle"
            SetListViewStyle (lvwReport)
        Case Is = "QuickSearch"
            QuickSearch
        Case Is = "ExportToExcel"
            ExportRecordList mrsRecordList
        Case Is = "Cerrar"
            Unload Me
    End Select
    
End Sub

Private Sub Form_Resize()

    ListView_Resize lvwItems, Me

End Sub

Public Sub QuickSearch()
    
    ListviewQuickSearch lvwItems, mlngColumn

End Sub

Public Sub ResultSearch()

    Dim frmBusqueda As ConsultaEdit
   
    Set frmBusqueda = New ConsultaEdit
  
    mobjBusqueda.ConsultaCampos "ArticuloColores"
    frmBusqueda.Component mobjBusqueda
    frmBusqueda.Show vbModal
    
    If frmBusqueda.mflgAplicarFiltro Then
        Set mobjBusqueda = frmBusqueda.Consulta
        SentenciaSQL = frmBusqueda.SentenciaSQL
        UpdateListView (SentenciaSQL)
    ElseIf frmBusqueda.lvwConsultaItems.ListItems.Count = 0 Then
        SentenciaSQL = vbNullString
    End If
    
    Unload frmBusqueda

End Sub

Public Sub Imprimir()
    Dim objItem As ListItem
    Dim objPrintClass As PrintClass
    Dim frmPrintOptions As frmPrint
    
    On Error GoTo ErrorManager
    
    Set frmPrintOptions = New frmPrint
    frmPrintOptions.Flags = ShowCopies_po + ShowPrinter_po
    frmPrintOptions.Copies = 1
    frmPrintOptions.Show vbModal
    ' salir de la opcion si no pulsa "imprimir"
    If Not frmPrintOptions.PrintDoc Then
        Unload frmPrintOptions
        Set frmPrintOptions = Nothing
        Exit Sub
    End If
        
    Set objPrintClass = New PrintClass
    objPrintClass.PrinterNumber = frmPrintOptions.PrinterNumber
    objPrintClass.Copies = frmPrintOptions.Copies
    
    objPrintClass.Titulo = "Articulos de la temporada " & GescomMain.objParametro.TemporadaActual
    
    objPrintClass.Columnas = lvwItems.ColumnHeaders
    
    For Each objItem In lvwItems.ListItems
        objPrintClass.Item = objItem
    Next
    objPrintClass.EndDoc

    Unload frmPrintOptions
    Set frmPrintOptions = Nothing
    Set objPrintClass = Nothing
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    
End Sub

