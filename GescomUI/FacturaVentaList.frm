VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "gridex20_b.ocx"
Object = "{E51AF47C-BB3B-48B6-A74A-7DA1722D2C68}#3.0#0"; "EntityProxy.ocx"
Begin VB.Form FacturaVentaList 
   Caption         =   "Lista de Facturas de Venta"
   ClientHeight    =   7560
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
   Icon            =   "FacturaVentaList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
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
         TabIndex        =   6
         Top             =   840
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
      End
      Begin EntityProxy.ctlEntityProxy epTemporada 
         Height          =   495
         Left            =   120
         TabIndex        =   5
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
      Begin EntityProxy.ctlEntityProxy epAnio 
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   120
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
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
      FormatStyle(1)  =   "FacturaVentaList.frx":1042
      FormatStyle(2)  =   "FacturaVentaList.frx":116A
      FormatStyle(3)  =   "FacturaVentaList.frx":121A
      FormatStyle(4)  =   "FacturaVentaList.frx":12CE
      FormatStyle(5)  =   "FacturaVentaList.frx":13A6
      FormatStyle(6)  =   "FacturaVentaList.frx":145E
      ImageCount      =   0
      PrinterProperties=   "FacturaVentaList.frx":153E
   End
End
Attribute VB_Name = "FacturaVentaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmFacturaVenta As FacturaVentaEdit
Private objFacturaVenta As FacturaVenta
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
        .Columns("FacturaVentaID").ColumnType = jgexIcon
        .Columns("FacturaVentaID").DefaultIcon = jgrdItems.GridImages(1).Index
        .Columns("FacturaVentaID").Caption = vbNullString
        .Columns("FacturaVentaID").Visible = True
        .Columns("FacturaVentaID").ColPosition = 1
        .Columns("FacturaVentaID").Width = 330
        
        FormatoJColumn .Columns("Numero"), 2, "Número"
        FormatoJColumn .Columns("Fecha"), 3, "Fecha"
        FormatoJColumn .Columns("NombreRepresentante"), 4, "Representante", , ColumnSize(12)
        FormatoJColumn .Columns("NombreCliente"), 5, "Cliente"
        .Columns("NombreCliente").ButtonStyle = jgexButtonEllipsis
        FormatoJColumn .Columns("BrutoEUR"), 6, "Total Bruto", True ', vbRightJustify
        FormatoJColumn .Columns("DescuentoEUR"), 7, "Descuento", True ', vbRightJustify
        FormatoJColumn .Columns("BaseImponibleEUR"), 8, "Base Imponible", True ', vbRightJustify
        FormatoJColumn .Columns("IVAEUR"), 9, "IVA", True ', vbRightJustify
        FormatoJColumn .Columns("RecargoEUR"), 10, "Recargo", True ', vbRightJustify
        FormatoJColumn .Columns("NetoEUR"), 11, "TOTAL", True ', vbRightJustify
        FormatoJColumn .Columns("SituacionContable"), 12, "Situación", False
        FormatoJColumn .Columns("Observaciones"), 13, "Observaciones", False, ColumnSize(8)
                        
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
    ' - Imprimir las facturas.
    ' - Contabilizar las facturas.
    ' - Comisiones.
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Documento", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("PrintDocument").Key)
    objButton.ToolTipText = "Imprimir las facturas seleccionadas"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Contabilizar", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Contabilidad").Key)
    objButton.ToolTipText = "Contabilizar las facturas seleccionadas"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Comisiones", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Representante").Key)
    objButton.ToolTipText = "Comisiones a representantes"
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "DESContabilizar", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Contabilidad").Key)
    objButton.ToolTipText = "DESContabilizar las facturas seleccionadas"
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Declaracion347", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Proveedor").Key)
    objButton.ToolTipText = "347-Relación de facturas > 3005,06€"
    
    Set mobjBusqueda = New Consulta
    
    mlngColumn = 1
    
    ' Criterios de seleccion, filtros
    epEmpresa.Initialize 1, "Empresas", "EmpresaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.objParametro.EmpresaActual, 0
    epEmpresa.LoadControl "Empresa"
    
    epTemporada.Initialize 1, "Temporadas", "TemporadaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epTemporada.LoadControl "Temporada"
    
    epCliente.Initialize 1, "Clientes", "ClienteID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epCliente.LoadControl "Cliente"
    
    epAnio.Initialize 1, "Anios", "Anio", "Anio", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, Year(Date), 0
    epAnio.LoadControl "Año"
    
    jgrdItems.ImageHeight = 16
    jgrdItems.ImageWidth = 16
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("FacturaVenta").Picture
    
    strLayout = vbNullString
    
End Sub

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
                    Set objFacturaVenta = New FacturaVenta
                    objFacturaVenta.Load mlngID
                    objFacturaVenta.BeginEdit
                    If objFacturaVenta.HayFacturaComplementaria Then
                       ' aquí hay que avisar de que se borrará la factura complementaria
                        Respuesta = MostrarMensaje(MSG_DELETEFACTURA)
                        If Respuesta = vbYes Then
                            objFacturaVenta.Delete
                            objFacturaVenta.ApplyEdit
                        Else
                            objFacturaVenta.CancelEdit
                        End If
                    Else
                        objFacturaVenta.Delete
                        objFacturaVenta.ApplyEdit
                    End If
                    Set objFacturaVenta = Nothing
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

   If epAnio.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epAnio.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epAnio.ClausulaWhere
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
    Set mrsRecordList = objRecordList.Load("Select * from vFacturasVenta", strClausulaWhere)
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
                Set frmFacturaVenta = New FacturaVentaEdit
                Set objFacturaVenta = New FacturaVenta
                objFacturaVenta.Load mlngID
                frmFacturaVenta.Component objFacturaVenta
                frmFacturaVenta.Show
                Set objFacturaVenta = Nothing
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub


ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub NewObject()

    Set frmFacturaVenta = New FacturaVentaEdit
    Set objFacturaVenta = New FacturaVenta
    frmFacturaVenta.Component objFacturaVenta
    frmFacturaVenta.Show
    Set objFacturaVenta = Nothing

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
        Case Is = "Contabilizar"
            Contabilizar
        Case Is = "Comisiones"
            Comisiones
        Case Is = "DESContabilizar"
            DESContabilizar
        Case Is = "Declaracion347"
            Declaracion347
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
  
    mobjBusqueda.ConsultaCampos "vFacturasVenta"
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
        .HeaderString(jgexHFCenter) = "Listado de facturas de venta"
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

Private Sub Contabilizar()
    Dim Respuesta As VbMsgBoxResult
    Dim flgForzar As Boolean
    Dim flgPreguntar As Boolean
    Dim flgAbortar As Boolean
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    On Error GoTo ErrorManager
   
    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos contabilizar
    Respuesta = MostrarMensaje(MSG_CONTABILIZAR)
    flgPreguntar = True
    flgForzar = False
    flgAbortar = False
    
    Screen.MousePointer = vbHourglass
    If Respuesta = vbYes Then
        For Each simTemp In jgrdItems.SelectedItems
            If simTemp.RowType = jgexRowTypeRecord Then
                Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
                mlngID = RowData(1)
                If mlngID > 0 Then
                    ContabilizaItems mlngID, flgAbortar, flgPreguntar, flgForzar
                End If
                ' Abortamos si se ha pedido al contabilizar
                If flgAbortar Then Exit For
            End If
        Next
    End If
    Screen.MousePointer = vbDefault

    ' aquí hay que avisar de que la contabilidad ha ido OK
    Respuesta = MostrarMensaje(MSG_CONTABILIZAR_OK)
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub ContabilizaItems(mlngID As Long, ByRef flgAbortar As Boolean, ByRef flgPreguntar As Boolean, ByRef flgForzar As Boolean)
    Dim Respuesta As VbMsgBoxResult
    
    Set objFacturaVenta = New FacturaVenta
    objFacturaVenta.Load mlngID
    If objFacturaVenta.Neto = 0 Then Exit Sub
    ' Si ya está contabilizado hay que:
    ' - preguntar si re-contabilizar todo
    ' - no recontabilizar si se dice que no (por defecto)
    ' - recontabilizar si se dice que si.
    ' - abortar si se pide
    If objFacturaVenta.Contabilizado Then
        ' ¿Hay que preguntar que hacer?
        If flgPreguntar Then
            flgPreguntar = False
            Respuesta = MostrarMensaje(MSG_VOLVER_A_CONTABILIZAR)
            Select Case Respuesta
            Case vbNo
                flgForzar = False
                Exit Sub
            Case vbYes
                flgForzar = True
            Case vbCancel
                flgAbortar = True
                Exit Sub
            End Select
        Else
            If Not flgForzar Then
                Exit Sub
            End If
        End If
    End If
        
    objFacturaVenta.Contabilizar flgForzar
    
    Set objFacturaVenta = Nothing

End Sub

Public Sub PrintSelected()
    Dim Respuesta As VbMsgBoxResult
    Dim objPrintFactura As PrintFactura
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
        frmPrintOptions.Copies = 2
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
                    Set objFacturaVenta = New FacturaVenta
                    Set objPrintFactura = New PrintFactura

                    objFacturaVenta.Load mlngID

                    objPrintFactura.PrinterNumber = frmPrintOptions.PrinterNumber
                    objPrintFactura.Copies = frmPrintOptions.Copies
                    objPrintFactura.Component objFacturaVenta

                    objPrintFactura.PrintObject

                    Set objPrintFactura = Nothing
                    Set objFacturaVenta = Nothing
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

Private Sub DESContabilizar()
    Dim Respuesta As VbMsgBoxResult
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    On Error GoTo ErrorManager
   
    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos descontabilizar
    Respuesta = MostrarMensaje(MSG_DESCONTABILIZAR)
    
    If Respuesta = vbYes Then
        Screen.MousePointer = vbHourglass
        For Each simTemp In jgrdItems.SelectedItems
            If simTemp.RowType = jgexRowTypeRecord Then
                Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
                mlngID = RowData(1)
                If mlngID > 0 Then
                    DESContabilizaItems mlngID
                End If
            End If
        Next
        Screen.MousePointer = vbDefault
        ' aquí hay que avisar de que la contabilidad ha ido OK
        Respuesta = MostrarMensaje(MSG_PROCESO_OK)
    End If
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub DESContabilizaItems(mlngID As Long)
    
    Set objFacturaVenta = New FacturaVenta
    objFacturaVenta.Load mlngID
    objFacturaVenta.BeginEdit
    objFacturaVenta.DESContabilizar
    objFacturaVenta.ApplyEdit
    Set objFacturaVenta = Nothing

End Sub

Private Sub Comisiones()
    Dim frmComisiones As ComisionesEdit
    Dim objComisiones As Comisiones
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    For Each simTemp In jgrdItems.SelectedItems
        If simTemp.RowType = jgexRowTypeRecord Then
            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
            mlngID = RowData(1)
            If mlngID > 0 Then
                Set frmComisiones = New ComisionesEdit
                Set objComisiones = New Comisiones
                'objComisiones.TemporadaID = GescomMain.objParametro.TemporadaActualID
                objComisiones.EmpresaID = GescomMain.objParametro.EmpresaActualID
                frmComisiones.Component objComisiones
                frmComisiones.Show vbModal
                Set objComisiones = Nothing
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

Private Sub Declaracion347()
Dim strAnio As String
Dim lngEmpresa As Long
Dim strWhere As String
Dim rlDeclaracion347 As RecordList
Dim rsDeclaracion347 As ADOR.Recordset
    
    On Error GoTo ErrorManager
    
    ' Nos aseguramos de que se haya seleccionado un año para obtener la declaración 347.
    strAnio = epAnio.SelectedValue
    If strAnio = vbNullString Then
        MsgBox "No se ha seleccionado un año para la declaración 347", vbOKOnly, "Declaración 347"
        Exit Sub
    End If
    
    lngEmpresa = epEmpresa.Selectedkey
    If lngEmpresa = 0 Then
        MsgBox "No se ha seleccionado una empresa para la declaración 347", vbOKOnly, "Declaración 347"
        Exit Sub
    End If
    
    Set rlDeclaracion347 = New RecordList
    strWhere = "AnioDeclaracion = '" & strAnio & "' AND EmpresaID=" & lngEmpresa
    Set rsDeclaracion347 = rlDeclaracion347.Load("SELECT * FROM vrptDeclaracion347", strWhere)
    Set rlDeclaracion347 = Nothing
    ExportRSToExcel rsDeclaracion347
    Set rsDeclaracion347 = Nothing
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub
