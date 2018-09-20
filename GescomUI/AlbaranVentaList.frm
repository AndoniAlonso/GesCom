VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "gridex20_b.ocx"
Object = "{E51AF47C-BB3B-48B6-A74A-7DA1722D2C68}#3.0#0"; "EntityProxy.ocx"
Begin VB.Form AlbaranVentaList 
   Caption         =   "Lista de Albaranes de Venta"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AlbaranVentaList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   11055
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
      Width           =   11055
      _ExtentX        =   19500
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
      FormatStyle(1)  =   "AlbaranVentaList.frx":1042
      FormatStyle(2)  =   "AlbaranVentaList.frx":116A
      FormatStyle(3)  =   "AlbaranVentaList.frx":121A
      FormatStyle(4)  =   "AlbaranVentaList.frx":12CE
      FormatStyle(5)  =   "AlbaranVentaList.frx":13A6
      FormatStyle(6)  =   "AlbaranVentaList.frx":145E
      ImageCount      =   0
      PrinterProperties=   "AlbaranVentaList.frx":153E
   End
End
Attribute VB_Name = "AlbaranVentaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmAlbaranVenta As AlbaranVentaEdit
Private frmAlbaranVentaAutomatico As AlbaranAutomaticoEdit
Private objAlbaranVenta As AlbaranVenta
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
        .Columns("AlbaranVentaID").ColumnType = jgexIcon
        .Columns("AlbaranVentaID").DefaultIcon = jgrdItems.GridImages(1).Index
        .Columns("AlbaranVentaID").Caption = vbNullString
        .Columns("AlbaranVentaID").Visible = True
        .Columns("AlbaranVentaID").ColPosition = 1
        .Columns("AlbaranVentaID").Width = 330
        
        FormatoJColumn .Columns("Numero"), 2, "Número"
        FormatoJColumn .Columns("Fecha"), 3, "Fecha"
        FormatoJColumn .Columns("NombreCliente"), 4, "Cliente"
        .Columns("NombreCliente").ButtonStyle = jgexButtonEllipsis
        FormatoJColumn .Columns("TotalBrutoEUR"), 5, "Total Bruto", True, , enFormatoImporte
        FormatoJColumn .Columns("Cantidad"), 6, "Cantidad", True, , enFormatoCantidad
        FormatoJColumn .Columns("Facturado"), 7, "Facturado"
        FormatoJColumn .Columns("NombreRepresentante"), 8, "Representante", , ColumnSize(12)
        FormatoJColumn .Columns("Observaciones"), 9, "Observaciones", False, ColumnSize(8)
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
    ' - Imprimir los albaranes.
    ' - Crear la factura desde el albarán.
    ' - Crear la factura desde el albarán según el porcentaje A-B del cliente.

    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Documento", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("PrintDocument").Key)
    objButton.ToolTipText = "Imprimir los albaranes seleccionados"
    
'    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
'    Set objButton = Me.tlbHerramientas.Buttons.Add(, "FacturarAlbaran", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("FacturaVenta").Key)
'    objButton.ToolTipText = "Facturar los albaranes seleccionados"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "FacturarAlbaranAB", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("FacturaVenta").Key)
    objButton.ToolTipText = "Facturar los albaranes según el porcentaje A-B"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "AlbaranAutomatico", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("BarCode").Key)
    objButton.ToolTipText = "Confeccionar el albarán capturando el código de barras"
    
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
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("AlbaranVenta").Picture
    
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
        For Each simTemp In jgrdItems.SelectedItems
            If simTemp.RowType = jgexRowTypeRecord Then
                Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
                mlngID = RowData(1)
                If mlngID > 0 Then
                    Set objAlbaranVenta = New AlbaranVenta
                    objAlbaranVenta.Load mlngID
                    objAlbaranVenta.BeginEdit
                    objAlbaranVenta.Delete
                    objAlbaranVenta.ApplyEdit
                    Set objAlbaranVenta = Nothing
                End If
            End If
        Next
        UpdateListView SentenciaSQL
    End If
    Exit Sub

ErrorManager:
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
    Set mrsRecordList = objRecordList.Load("Select * from vAlbaranesVenta", strClausulaWhere)
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
                Set frmAlbaranVenta = New AlbaranVentaEdit
                Set objAlbaranVenta = New AlbaranVenta
                objAlbaranVenta.Load mlngID
                frmAlbaranVenta.Component objAlbaranVenta
                frmAlbaranVenta.Show
                Set objAlbaranVenta = Nothing
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub


ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub NewObject()

    Set frmAlbaranVenta = New AlbaranVentaEdit
    Set objAlbaranVenta = New AlbaranVenta
    frmAlbaranVenta.Component objAlbaranVenta
    frmAlbaranVenta.Show
    Set objAlbaranVenta = Nothing

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
        Case Is = "Documento"
            PrintSelected
'        Case Is = "FacturarAlbaran"
'            FacturarSelected
        Case Is = "FacturarAlbaranAB"
            FacturarSelectedAB
        Case Is = "AlbaranAutomatico"
            AlbaranAutomatico
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
  
    mobjBusqueda.ConsultaCampos "vAlbaranesVenta"
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
    On Error GoTo ErrorManager
    
    With jgrdItems.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = False
        .TranslateColors = True
        .HeaderString(jgexHFCenter) = "Listado de albaranes de venta"
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

Private Sub jgrdItems_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)

    jgrdItems.PrinterProperties.FooterString(jgexHFRight) = "Página " & PageNumber & vbCrLf & " de " & nPages
    
End Sub

'Public Sub FacturarSelected()
'    Dim Respuesta As VbMsgBoxResult
'
'    On Error GoTo ErrorManager
'
'    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
'
'    ' aquí hay que avisar de si realmente queremos abrirlos todos
'    ' si el número es mayor que 5
'    Respuesta = MostrarMensaje(MSG_FACTURARALBARAN)
'
'
'    If Respuesta = vbYes Then
'        FacturarItems
'    End If
'
'
'    Exit Sub
'
'ErrorManager:
'    ManageErrors (Me.Caption)
'End Sub
'
'Public Sub FacturarItems()
'    Dim objFacturaVenta As FacturaVenta
'    Dim objAlbaranVenta As AlbaranVenta
'    Dim frmFacturaVenta As FacturaVentaEdit
'    Dim objAlbaranVentaItem As AlbaranVentaItem
'    Dim objFacturaVentaItem As FacturaVentaItem
'    Dim simTemp As JSSelectedItem
'    Dim RowData As JSRowData
'
'    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
'    Screen.MousePointer = vbHourglass
'
'    For Each simTemp In jgrdItems.SelectedItems
'        If simTemp.RowType = jgexRowTypeRecord Then
'            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
'            mlngID = RowData(1)
'            If mlngID > 0 Then
'                Set objAlbaranVenta = New AlbaranVenta
'                objAlbaranVenta.Load mlngID
'                If objAlbaranVenta.AlbaranVentaItems.Facturado Then _
'                    Err.Raise vbObjectError + 1001, "Albarán " & _
'                                objAlbaranVenta.Numero & " ya facturado, no se genera factura"
'
'                Set objFacturaVenta = New FacturaVenta
'                With objFacturaVenta
'                    .BeginEdit
'                    .EmpresaID = objAlbaranVenta.EmpresaID
'                    .TemporadaID = objAlbaranVenta.TemporadaID
'                    .Cliente = objAlbaranVenta.Cliente
'                    .Transportista = objAlbaranVenta.Transportista
'                    .Representante = objAlbaranVenta.Representante
'                    .FormaPago = objAlbaranVenta.FormaPago
'                    .NuestraReferencia = objAlbaranVenta.NuestraReferencia
'                    .SuReferencia = objAlbaranVenta.SuReferencia
'                    .Observaciones = objAlbaranVenta.Observaciones
'                    .Embalajes = objAlbaranVenta.Embalajes
'                    .Portes = objAlbaranVenta.Portes
'                    .DatoComercial.ChildBeginEdit
'                    .DatoComercial.Descuento = objAlbaranVenta.DatoComercial.Descuento
'                    .DatoComercial.RecargoEquivalencia = objAlbaranVenta.DatoComercial.RecargoEquivalencia
'                    .DatoComercial.IVA = objAlbaranVenta.DatoComercial.IVA
'                    .DatoComercial.ChildApplyEdit
'                    .Fecha = objAlbaranVenta.Fecha
'                    .Numero = objAlbaranVenta.Numero
'                    For Each objAlbaranVentaItem In objAlbaranVenta.AlbaranVentaItems
'                        If Not objAlbaranVentaItem.Facturado Then
'                            Set objFacturaVentaItem = .FacturaVentaItems.Add
'                            objFacturaVentaItem.BeginEdit
'                            objFacturaVentaItem.FacturaDesdeAlbaran objAlbaranVentaItem.AlbaranVentaItemID
'                            objFacturaVentaItem.ApplyEdit
'                            Set objFacturaVentaItem = Nothing
'                        End If
'
'                    Next
'                    .CalcularBruto
'                    .CrearCobros
'                    .ApplyEdit
'
'                End With
'
'                Set frmFacturaVenta = New FacturaVentaEdit
'                frmFacturaVenta.Component objFacturaVenta
'                frmFacturaVenta.Show
'
'                Set objAlbaranVenta = Nothing
'                Set objFacturaVenta = Nothing
'            End If
'        End If
'    Next
'
'    Screen.MousePointer = vbDefault
'    Exit Sub
'
'ErrorManager:
'    Screen.MousePointer = vbDefault
'    ManageErrors (Me.Caption)
'End Sub

Public Sub PrintSelected()
Dim Respuesta As VbMsgBoxResult
Dim objPrintAlbaran As PrintAlbaran
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
                    Set objAlbaranVenta = New AlbaranVenta
                    Set objPrintAlbaran = New PrintAlbaran
                    
                    objAlbaranVenta.Load mlngID
                    
                    objPrintAlbaran.PrinterNumber = frmPrintOptions.PrinterNumber
                    objPrintAlbaran.Copies = frmPrintOptions.Copies
                    objPrintAlbaran.Component objAlbaranVenta
                    
                    objPrintAlbaran.PrintObject
                    
                    Set objPrintAlbaran = Nothing
                    Set objAlbaranVenta = Nothing
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

Public Sub FacturarSelectedAB()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos abrirlos todos
    ' si el número es mayor que 5
    Respuesta = MostrarMensaje(MSG_FACTURARALBARAN)

    If Respuesta = vbYes Then
        FacturarItemsAB
    End If
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub FacturarItemsAB()
'    Dim objFacturaVenta As FacturaVenta
    Dim objAlbaranVenta As AlbaranVenta
    Dim frmFacturaVentaA As FacturaVentaEdit
    Dim frmFacturaVentaB As FacturaVentaEdit
'    Dim objAlbaranVentaItem As AlbaranVentaItem
'    Dim objFacturaVentaItem As FacturaVentaItem
    Dim colListaAlbaranes As Collection
    Dim lngFacturaVentaIDA As Long
    Dim lngFacturaVentaIDB As Long
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    Dim objCliente As Cliente
    Dim PorcFacturacionAB As Double
    Dim frmPorcFacturacion As frmPorcFacturacionAB
    
    
  
    Set colListaAlbaranes = New Collection
    
    For Each simTemp In jgrdItems.SelectedItems
        If simTemp.RowType = jgexRowTypeRecord Then
            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
            mlngID = RowData(1)
            If mlngID > 0 Then
                colListaAlbaranes.Add mlngID
            End If
        End If
    Next
    
    If colListaAlbaranes.Count > 0 Then
        ' Tratar el cliente del primer albaran, obtener su porcentaje de facturación y permitir cambiarlo
        Set objAlbaranVenta = New AlbaranVenta
        objAlbaranVenta.Load colListaAlbaranes(1)

        Set objCliente = New Cliente
        objCliente.Load objAlbaranVenta.ClienteID
        
        Set frmPorcFacturacion = New frmPorcFacturacionAB
        frmPorcFacturacion.PorcFacturacionAB = objCliente.PorcFacturacionAB
        frmPorcFacturacion.Cliente = objCliente.Nombre
        frmPorcFacturacion.Show vbModal
        If Not frmPorcFacturacion.Ok Then
            Unload frmPorcFacturacion
            Set frmPorcFacturacion = Nothing
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            PorcFacturacionAB = frmPorcFacturacion.PorcFacturacionAB
            Unload frmPorcFacturacion
            Set frmPorcFacturacion = Nothing
        End If
        
        Set objCliente = Nothing
        Set objAlbaranVenta = Nothing
        
        ' Como este proceso puede ser lento muestro el puntero de reloj de arena
        Screen.MousePointer = vbHourglass
    
        Set objAlbaranVenta = New AlbaranVenta
        objAlbaranVenta.FacturarAlbaranAB colListaAlbaranes, lngFacturaVentaIDA, lngFacturaVentaIDB, PorcFacturacionAB
        Set objAlbaranVenta = Nothing
        
        Dim objFacturaVentaA As FacturaVenta
        Dim objFacturaVentaB As FacturaVenta
        Set objFacturaVentaB = New FacturaVenta
        If lngFacturaVentaIDA Then
            Set objFacturaVentaA = New FacturaVenta
            objFacturaVentaA.Load lngFacturaVentaIDA
            Set frmFacturaVentaA = New FacturaVentaEdit
            frmFacturaVentaA.Component objFacturaVentaA
            frmFacturaVentaA.Show
            Set objFacturaVentaA = Nothing
        End If
    
        If lngFacturaVentaIDB Then
            Set objFacturaVentaB = New FacturaVenta
            objFacturaVentaB.Load lngFacturaVentaIDB
            Set frmFacturaVentaB = New FacturaVentaEdit
            frmFacturaVentaB.Component objFacturaVentaB
            frmFacturaVentaB.Show
            Set objFacturaVentaB = Nothing
        End If
        
    End If

    UpdateListView
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub AlbaranAutomatico()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos abrirlos todos
    ' si el número es mayor que 5
    If jgrdItems.SelectedItems.Count >= 5 Then
        Respuesta = MostrarMensaje(MSG_OPEN)
        
        If Respuesta = vbYes Then
            AbrirAutomaticos
        End If
    Else
        AbrirAutomaticos
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub
    
Private Sub AbrirAutomaticos()
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    For Each simTemp In jgrdItems.SelectedItems
        If simTemp.RowType = jgexRowTypeRecord Then
            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
            mlngID = RowData(1)
            If mlngID > 0 Then
                Set frmAlbaranVentaAutomatico = New AlbaranAutomaticoEdit
                Set objAlbaranVenta = New AlbaranVenta
                objAlbaranVenta.Load mlngID
                frmAlbaranVentaAutomatico.Component objAlbaranVenta
                frmAlbaranVentaAutomatico.Show
                Set objAlbaranVenta = Nothing
            End If
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


