VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "gridex20_b.ocx"
Object = "{E51AF47C-BB3B-48B6-A74A-7DA1722D2C68}#2.1#0"; "EntityProxy.ocx"
Begin VB.Form PedidoVentaList 
   Caption         =   "Lista de Pedidos de Venta"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PedidoVentaList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFiltro 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   10815
      Begin EntityProxy.ctlEntityProxy epCliente 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
      End
      Begin EntityProxy.ctlEntityProxy epTemporada 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
      End
      Begin EntityProxy.ctlEntityProxy epEmpresa 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
      End
   End
   Begin MSComctlLib.Toolbar tlbHerramientas 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
   End
   Begin GridEX20.GridEX jgrdItems 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7223
      Version         =   "2.0"
      AllowRowSizing  =   -1  'True
      AutomaticSort   =   -1  'True
      ScrollToolTips  =   -1  'True
      RecordNavigatorString=   "Registro:|de"
      ShowToolTips    =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ShowEmptyFields =   0   'False
      CalendarTodayText=   "Hoy"
      CalendarNoneText=   "Ninguno"
      MultiSelect     =   -1  'True
      HeaderStyle     =   2
      UseEvenOddColor =   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      GroupByBoxInfoText=   "Arrastre una columna aquí para agrupar."
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      HeaderFontSize  =   7.5
      FontSize        =   7.5
      ColumnHeaderHeight=   285
      IntProp1        =   0
      FormatStylesCount=   6
      FormatStyle(1)  =   "PedidoVentaList.frx":1042
      FormatStyle(2)  =   "PedidoVentaList.frx":116A
      FormatStyle(3)  =   "PedidoVentaList.frx":121A
      FormatStyle(4)  =   "PedidoVentaList.frx":12CE
      FormatStyle(5)  =   "PedidoVentaList.frx":13A6
      FormatStyle(6)  =   "PedidoVentaList.frx":145E
      ImageCount      =   0
      PrinterProperties=   "PedidoVentaList.frx":153E
   End
End
Attribute VB_Name = "PedidoVentaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmPedidoVenta As PedidoVentaEdit
Private objPedidoVenta As PedidoVenta
Private mobjBusqueda As Consulta
Public SentenciaSQL As String
Private strLayout As String

Public Sub ComponentStatus(rsStatus As ADOR.Recordset)
   
    Set mrsRecordList = rsStatus
    Call RefreshListView

End Sub

Private Sub RefreshListView()
    Dim jcoltemp As JSColumn
    
    
    Set jgrdItems.ADORecordset = mrsRecordList
    ' Ocultamos TODAS las columnas
    For Each jcoltemp In jgrdItems.Columns
        jcoltemp.Visible = False
    Next
        
    With jgrdItems
        ' Propiedades de la 1ª columna, campo clave.
        .Columns("PedidoVentaID").ColumnType = jgexIcon
        .Columns("PedidoVentaID").DefaultIcon = jgrdItems.GridImages(1).Index
        .Columns("PedidoVentaID").Caption = vbNullString
        .Columns("PedidoVentaID").Visible = True
        .Columns("PedidoVentaID").ColPosition = 1
        .Columns("PedidoVentaID").Width = 330
        
        FormatoJColumn .Columns("Numero"), 2, "Número"
        FormatoJColumn .Columns("Fecha"), 3, "Fecha"
        FormatoJColumn .Columns("FechaTopeServicio"), 4, "Tope Servicio", , ColumnSize(12)
        FormatoJColumn .Columns("NombreCliente"), 5, "Cliente"
        .Columns("NombreCliente").ButtonStyle = jgexButtonEllipsis
        FormatoJColumn .Columns("TotalBrutoEUR"), 6, "Total Bruto", True
        FormatoJColumn .Columns("TotalPedido"), 7, "Pedido", True, , enFormatoCantidad
        FormatoJColumn .Columns("TotalServido"), 8, "Servido", True, , enFormatoCantidad
        FormatoJColumn .Columns("PorcentajeServicio"), 9, "% Servicio", False, ColumnSize(5), enFormatoPorcentaje
        FormatoJColumn .Columns("NombreRepresentante"), 10, "Representante", False, ColumnSize(10)
        FormatoJColumn .Columns("Observaciones"), 11, "Observaciones", False, ColumnSize(20)
        
'        jgrdItems_RowFormat jgrdItems.GetRowData(jgrdItems.Row)
         
        jgrdItems.GroupFooterStyle = jgexTotalsGroupFooter
        
        Dim jscolAgrupar As JSColumn
        Set jscolAgrupar = .Columns.Add(, jgexText, jgexEditNone)
        jscolAgrupar.AggregateFunction = jgexCount
        jscolAgrupar.Width = 0
        jscolAgrupar.Caption = "(Agrupar todos)"
        jscolAgrupar.GroupEmptyStringCaption = "(Agrupar todos)"

        .Groups.Add jscolAgrupar.Index, jgexSortAscending
    
    End With

    If strLayout <> vbNullString Then jgrdItems.LoadLayoutString strLayout

End Sub

Private Sub Form_Load()
    Dim objButton As Button

    Me.Move 0, 0
    Me.WindowState = vbMaximized
    
    jgrdItems.View = jgexTable
    
    LoadImages Me.tlbHerramientas
    
    ' Añadimos los botones especificos de esta opción:
    ' - Tallaje máximo y mínimo de pedidos.
    ' - Imprimir los pedidos.
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "TallajePedido", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("OLAPQuery").Key)
    objButton.ToolTipText = "Tallaje máximo y mínimo por modelo"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Documento", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("PrintDocument").Key)
    objButton.ToolTipText = "Imprimir los pedidos seleccionados"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "PreciosArticulos", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Remesa").Key)
    objButton.ToolTipText = "Actualizar los precios de venta en los pedidos seleccionados"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "NecesidadesMaterial", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Material").Key)
    objButton.ToolTipText = "Cálculo de necesidades de material"
    
    Set mobjBusqueda = New Consulta
    
    mlngColumn = 1
    
    ' Criterios de seleccion, filtros
    epEmpresa.Initialize 1, "Empresas", "EmpresaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.objParametro.EmpresaActual, 0
    epEmpresa.LoadControl "Empresa"
    
    epTemporada.Initialize 1, "Temporadas", "TemporadaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.objParametro.TemporadaActual, 0
    epTemporada.LoadControl "Temporada"
    
    epCliente.Initialize 1, "Clientes", "ClienteID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epCliente.LoadControl "Cliente"
    
    jgrdItems.ImageHeight = 16
    jgrdItems.ImageWidth = 16
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("PedidoVenta").Picture
    
    strLayout = vbNullString
    
End Sub

''''???? para liberar memoria ?
Private Sub Form_Unload(Cancel As Integer)
    mrsRecordList.Close
    Set mrsRecordList = Nothing
End Sub

Private Sub jgrdItems_DblClick()
    
    Call EditSelected

End Sub

Private Sub jgrdItems_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    
    mlngColumn = Column.Index
    'txtQuickSearch.ToolTipText = "Búsqueda rápida en " & Column.Caption

End Sub



Private Sub jgrdItems_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        EditSelected
    ElseIf KeyCode = 46 Then
        DeleteSelected
    ElseIf KeyCode = 116 Then   ' Tecla F5
        UpdateListView
    ElseIf KeyCode = 114 Then   ' Tecla F3
        QuickSearch
    End If
        
End Sub

Private Sub jgrdItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        Me.PopupMenu GescomMain.mnuListView
        jgrdItems.Enabled = False
        jgrdItems.Enabled = True
    End If
    
End Sub

Public Sub DeleteSelected()
    Dim Respuesta As VbMsgBoxResult
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    On Error GoTo ErrorManager
    
    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos borrar
    Respuesta = MostrarMensaje(MSG_DELETE)
    
    If Respuesta = vbYes Then
        ' Como este proceso puede ser lento muestro el puntero de reloj de arena
        Screen.MousePointer = vbHourglass
        For Each simTemp In jgrdItems.SelectedItems
            If simTemp.RowType = jgexRowTypeRecord Then
                Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
                mlngID = RowData(1)
                If mlngID > 0 Then
                    Set objPedidoVenta = New PedidoVenta
                    objPedidoVenta.Load mlngID, GescomMain.objParametro.Moneda
                    objPedidoVenta.BeginEdit GescomMain.objParametro.Moneda
                    objPedidoVenta.Delete
                    objPedidoVenta.ApplyEdit
                    Set objPedidoVenta = Nothing
                End If
            End If
        Next
        UpdateListView SentenciaSQL
        Screen.MousePointer = vbDefault
    End If
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Public Sub UpdateListView(Optional strWhere As String)
    Dim objRecordList As RecordList
    Dim strClausulaWhere As String

    On Error GoTo ErrorManager
    
    strClausulaWhere = epTemporada.ClausulaWhere

    If epEmpresa.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epEmpresa.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epEmpresa.ClausulaWhere
        End If
    End If

    If epCliente.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epCliente.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epCliente.ClausulaWhere
        End If
    End If
        
    If strWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = strWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & strWhere
        End If
    End If

    Set objRecordList = New RecordList
    '''???? para liberar memoria
    mrsRecordList.Close
    Set mrsRecordList = Nothing
    Set mrsRecordList = objRecordList.Load("Select * from vPedidosVenta", strClausulaWhere)
    Set objRecordList = Nothing
    
    strLayout = jgrdItems.LayoutString
    Call RefreshListView
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub EditSelected()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos abrirlos todos
    ' si el número es mayor que 5
    If jgrdItems.SelectedItems.Count >= 5 Then
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

Public Sub EditItems()
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
    For Each simTemp In jgrdItems.SelectedItems
        If simTemp.RowType = jgexRowTypeRecord Then
            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
            mlngID = RowData(1)
            If mlngID > 0 Then
                Set frmPedidoVenta = New PedidoVentaEdit
                Set objPedidoVenta = New PedidoVenta
                objPedidoVenta.Load mlngID, GescomMain.objParametro.Moneda
                frmPedidoVenta.Component objPedidoVenta
                frmPedidoVenta.Show
                Set objPedidoVenta = Nothing
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub NewObject()

    Set frmPedidoVenta = New PedidoVentaEdit
    Set objPedidoVenta = New PedidoVenta
    frmPedidoVenta.Component objPedidoVenta
    frmPedidoVenta.Show
    Set objPedidoVenta = Nothing

End Sub

Public Sub SetListViewStyle(View As Integer)
   If View = jgexCard Then jgrdItems.View = jgexCard
   If View = 3 Then jgrdItems.View = jgexTable
   
   
End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
    IsList = True
   
End Function

Private Sub jgrdItems_RowFormat(RowBuffer As GridEX20.JSRowData)
    
    If RowBuffer.RowType = jgexRowTypeConstants.jgexRowTypeGroupFooter Then
        Dim lngServido  As Long
        Dim lngPedido  As Long
        Dim dblPorcentajeServicio As Double
        
        lngServido = RowBuffer.GetSubTotal(jgrdItems.Columns("TotalServido").Index, jgexSum)
        lngPedido = RowBuffer.GetSubTotal(jgrdItems.Columns("TotalPedido").Index, jgexSum)
        Debug.Assert lngPedido <> 0
        
        dblPorcentajeServicio = lngServido * 100 / lngPedido
        RowBuffer.DisplayValue(jgrdItems.Columns("PorcentajeServicio").Index) = Format(dblPorcentajeServicio, "###,##0.00\%")
    End If
 
End Sub

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
        Case Is = "GroupBy"
            frmGroupBy.GroupGrid jgrdItems
        Case Is = "ShowFields"
            frmShowfields.ShowFields jgrdItems
        Case Is = "Sort"
            frmSort.SortGrid jgrdItems
        Case Is = "Cerrar"
            Unload Me
        Case Is = "ExportToExcel"
            ExportRecordList mrsRecordList
        Case Is = "TallajePedido"
            TallajePedido
        Case Is = "Documento"
            PrintSelected
        Case Is = "PreciosArticulos"
            PreciosArticulos
        Case Is = "NecesidadesMaterial"
            NecesidadesMaterial
    End Select
        
End Sub

Private Sub Form_Resize()

    GridEX_Resize jgrdItems, Me, frmFiltro

End Sub

Public Sub QuickSearch()
    
    JanusQuickSearch jgrdItems, mlngColumn

End Sub

Public Sub ResultSearch()
    Dim frmBusqueda As ConsultaEdit
   
    Set frmBusqueda = New ConsultaEdit
  
    mobjBusqueda.ConsultaCampos "vPedidosVenta"
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

Private Sub jgrdItems_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)

    jgrdItems.PrinterProperties.FooterString(jgexHFRight) = "Página " & PageNumber & vbCrLf & " de " & nPages
    
End Sub

' Tallaje mínimo y máximo por modelo a partir de pedidos.
Public Sub TallajePedido()
    Dim frmList As OLAPQueryList
    Dim objOLAPQuery As OLAPQuery
    Dim strWhere As String

    Set frmList = New OLAPQueryList
    Set objOLAPQuery = New OLAPQuery
    strWhere = "TemporadaID = " & GescomMain.objParametro.TemporadaActualID
     
    objOLAPQuery.Load QRY_TallajePedido, strWhere
    With frmList
        .Component objOLAPQuery
        .Show vbModal
  
    End With

    Set objOLAPQuery = Nothing

End Sub

Public Sub Imprimir()
    On Error GoTo ErrorManager
    
    With jgrdItems.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = False
        .TranslateColors = True
        .HeaderString(jgexHFCenter) = "Listado de pedidos de venta"
        .FooterString(jgexHFLeft) = Now
        .BottomMargin = 720
        .TopMargin = 720
        .LeftMargin = 720
        .RightMargin = 720
        .TransparentBackground = True
        'Right footer is set in the BeforePrintPage to indicate page number
    End With
    Load frmPreview
    frmPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    jgrdItems.PrintPreview frmPreview.grPrev, False  '(Check4.value = vbChecked)
    frmPreview.Show vbModal
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    
End Sub

Public Sub PrintSelected()
    Dim Respuesta As VbMsgBoxResult
    Dim objPrintPedido As PrintPedido
    Dim frmPrintOptions As frmPrint
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    On Error GoTo ErrorManager
   
    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos imprimir los documentos
    Respuesta = MostrarMensaje(MSG_DOCUMENTO)
    
    If Respuesta = vbYes Then
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
        
        For Each simTemp In jgrdItems.SelectedItems
            If simTemp.RowType = jgexRowTypeRecord Then
                Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
                mlngID = RowData(1)
                If mlngID > 0 Then
                    Set objPedidoVenta = New PedidoVenta
                    Set objPrintPedido = New PrintPedido
                    objPedidoVenta.Load mlngID, GescomMain.objParametro.Moneda
                    
                    objPrintPedido.PrinterNumber = frmPrintOptions.PrinterNumber
                    objPrintPedido.Copies = frmPrintOptions.Copies
                    objPrintPedido.Component objPedidoVenta
                    
                    objPrintPedido.PrintObject
                    
                    Set objPrintPedido = Nothing
                    Set objPedidoVenta = Nothing
                End If
            End If
        Next
        Unload frmPrintOptions
        Set frmPrintOptions = Nothing
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub PreciosArticulos()
    Dim Respuesta As VbMsgBoxResult
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    On Error GoTo ErrorManager
   
    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos actualizar los precios de venta de los articulos
    Respuesta = MostrarMensaje(MSG_ACTUALIZAR_PRECIOS)
    
    Screen.MousePointer = vbHourglass
    
    If Respuesta = vbYes Then
            For Each simTemp In jgrdItems.SelectedItems
                If simTemp.RowType = jgexRowTypeRecord Then
                    Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
                    mlngID = RowData(1)
                If mlngID > 0 Then
                    Set objPedidoVenta = New PedidoVenta
                    objPedidoVenta.Load mlngID, GescomMain.objParametro.Moneda
                    objPedidoVenta.BeginEdit GescomMain.objParametro.Moneda
                    objPedidoVenta.ActualizarPreciosVenta
                    objPedidoVenta.ApplyEdit
                    
                    Set objPedidoVenta = Nothing
                End If
            End If
        Next
    End If
    
    Screen.MousePointer = vbDefault
    UpdateListView
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub NecesidadesMaterial()
    Dim frmNecesidadesMaterial As NecesidadesMaterialEdit
    Dim objNecesidadesMaterial As NecesidadesMaterial
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    For Each simTemp In jgrdItems.SelectedItems
        If simTemp.RowType = jgexRowTypeRecord Then
            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
            mlngID = RowData(1)
            If mlngID > 0 Then
                Set frmNecesidadesMaterial = New NecesidadesMaterialEdit
                Set objNecesidadesMaterial = New NecesidadesMaterial
                objNecesidadesMaterial.TemporadaID = GescomMain.objParametro.TemporadaActualID
                frmNecesidadesMaterial.Component objNecesidadesMaterial
                frmNecesidadesMaterial.Show vbModal
                Set objNecesidadesMaterial = Nothing
            End If
            Exit For
        End If
    Next

End Sub

Private Sub jgrdItems_ColButtonClick(ByVal ColIndex As Integer)
Dim objCliente As Cliente
Dim frmClienteEdit As ClienteEdit
Dim rdCurrent As JSRowData
Dim lngClienteID As Long

    If ColIndex = jgrdItems.Columns("NombreCliente").Index Then
        Set objCliente = New Cliente
        Set rdCurrent = jgrdItems.GetRowData(jgrdItems.Row)
        lngClienteID = rdCurrent.Value(jgrdItems.Columns("ClienteID").Index)
        objCliente.Load lngClienteID
        
        Set frmClienteEdit = New ClienteEdit
        frmClienteEdit.Component objCliente
        frmClienteEdit.Show
        
        Set objCliente = Nothing
        Set frmClienteEdit = Nothing
    End If
    
End Sub


