VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "gridex20_b.ocx"
Object = "{E51AF47C-BB3B-48B6-A74A-7DA1722D2C68}#3.0#0"; "EntityProxy.ocx"
Begin VB.Form FacturaCompraList 
   Caption         =   "Lista de Facturas de Compra"
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
   Icon            =   "FacturaCompraList.frx":0000
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
      Begin EntityProxy.ctlEntityProxy epTemporada 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epProveedor 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epAnio 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
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
      ButtonWidth     =   614
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
      FormatStyle(1)  =   "FacturaCompraList.frx":1042
      FormatStyle(2)  =   "FacturaCompraList.frx":116A
      FormatStyle(3)  =   "FacturaCompraList.frx":121A
      FormatStyle(4)  =   "FacturaCompraList.frx":12CE
      FormatStyle(5)  =   "FacturaCompraList.frx":13A6
      FormatStyle(6)  =   "FacturaCompraList.frx":145E
      ImageCount      =   0
      PrinterProperties=   "FacturaCompraList.frx":153E
   End
End
Attribute VB_Name = "FacturaCompraList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmFacturaCompra As FacturaCompraEdit
Private objFacturaCompra As FacturaCompra
Private mobjBusqueda As Consulta
Public SentenciaSQL As String
Private strLayout As String

Public Sub ComponentStatus(rsStatus As ADOR.Recordset)
   
    Set mrsRecordList = rsStatus
    Call RefreshListView

End Sub

Private Sub RefreshListView()
    Dim dblBrutoTotal As Double
    Dim dblDescuentoTotal As Double
    Dim dblIVATotal As Double
    Dim dblGastosTotal As Double
    Dim dblNetoTotal As Double
    Dim jcoltemp As JSColumn
    
    dblBrutoTotal = 0
    dblDescuentoTotal = 0
    dblIVATotal = 0
    dblGastosTotal = 0
    dblNetoTotal = 0

    Set jgrdItems.ADORecordset = mrsRecordList
    ' Ocultamos TODAS las columnas
    For Each jcoltemp In jgrdItems.Columns
        jcoltemp.Visible = False
    Next
    With jgrdItems
        ' Propiedades de la 1ª columna, campo clave.
        .Columns("FacturaCompraID").ColumnType = jgexIcon
        .Columns("FacturaCompraID").DefaultIcon = jgrdItems.GridImages(1).Index
        .Columns("FacturaCompraID").Caption = vbNullString
        .Columns("FacturaCompraID").Visible = True
        .Columns("FacturaCompraID").ColPosition = 1
        .Columns("FacturaCompraID").Width = 330
        
        FormatoJColumn .Columns("CuentaContable"), 2, "Cta.Contable", , ColumnSize(7)
        FormatoJColumn .Columns("CodigoFactura"), 3, "Número", , ColumnSize(6)
        FormatoJColumn .Columns("Fecha"), 4, "Fecha", , ColumnSize(7)
        FormatoJColumn .Columns("FechaContable"), 5, "Fec.Contable", , ColumnSize(7)
        FormatoJColumn .Columns("NombreProveedor"), 6, "Proveedor", , ColumnSize(20)
        .Columns("NombreProveedor").ButtonStyle = jgexButtonEllipsis
        FormatoJColumn .Columns("NetoEUR"), 7, "Neto", True ', vbRightJustify
        FormatoJColumn .Columns("BrutoEUR"), 8, "Bruto", True ', vbRightJustify
        FormatoJColumn .Columns("GastosEUR"), 9, "Gastos", True ', vbRightJustify
        FormatoJColumn .Columns("DescuentoEUR"), 10, "Descuento", True ', vbRightJustify
        FormatoJColumn .Columns("BaseImponibleEUR"), 11, "Base Imponible", True ', vbRightJustify
        FormatoJColumn .Columns("IVAEUR"), 12, "IVA", True ', vbRightJustify
        FormatoJColumn .Columns("SituacionContable"), 13, "Situación"
        FormatoJColumn .Columns("Observaciones"), 14, "Observaciones"
        FormatoJColumn .Columns("NombreBanco"), 15, "Banco"
        FormatoJColumn .Columns("NombreTransportista"), 16, "Transportista"
                        
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
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Contabilizar", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Contabilidad").Key)
    objButton.ToolTipText = "Contabilizar las facturas seleccionadas"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "DESContabilizar", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Contabilidad").Key)
    objButton.ToolTipText = "DESContabilizar las facturas seleccionadas"
    
    Set mobjBusqueda = New Consulta
    
    mlngColumn = 1
    
    ' Criterios de seleccion, filtros
    epTemporada.Initialize 1, "Temporadas", "TemporadaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0 ', GescomMain.objParametro.TemporadaActual
    epTemporada.LoadControl "Temporada"
    
    epProveedor.Initialize 1, "Proveedores", "ProveedorID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epProveedor.LoadControl "Proveedor"
    
    epAnio.Initialize 1, "Anios", "Anio", "Anio", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, Year(Date), 0
    epAnio.LoadControl "Año"
    
    jgrdItems.ImageHeight = 16
    jgrdItems.ImageWidth = 16
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("FacturaCompra").Picture
    
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
                    Set objFacturaCompra = New FacturaCompra
                    objFacturaCompra.Load mlngID, GescomMain.objParametro.Moneda
                    objFacturaCompra.BeginEdit GescomMain.objParametro.Moneda
                    objFacturaCompra.Delete
                    objFacturaCompra.ApplyEdit
                    Set objFacturaCompra = Nothing
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
    
   If epProveedor.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epProveedor.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epProveedor.ClausulaWhere
        End If
    End If

   If epAnio.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epAnio.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epAnio.ClausulaWhere
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
    Set mrsRecordList = objRecordList.Load("Select * from vFacturasCompra", strClausulaWhere)
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

    'If lvwItems.SelectedItem Is Nothing Then Exit Sub
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
                Set frmFacturaCompra = New FacturaCompraEdit
                Set objFacturaCompra = New FacturaCompra
                objFacturaCompra.Load mlngID, GescomMain.objParametro.Moneda
                frmFacturaCompra.Component objFacturaCompra
                frmFacturaCompra.Show
                Set objFacturaCompra = Nothing
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub


ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub NewObject()

    Set frmFacturaCompra = New FacturaCompraEdit
    Set objFacturaCompra = New FacturaCompra
    frmFacturaCompra.Component objFacturaCompra
    frmFacturaCompra.Show
    Set objFacturaCompra = Nothing

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
        Case Is = "Contabilizar"
            Contabilizar
        Case Is = "DESContabilizar"
            DESContabilizar
    End Select
        
End Sub

Private Sub Form_Resize()

    'ListView_Resize lvwItems, Me, frmFiltro
    GridEX_Resize jgrdItems, Me, frmFiltro

End Sub

Public Sub QuickSearch()
    
    JanusQuickSearch jgrdItems, mlngColumn

End Sub

Public Sub ResultSearch()
    Dim frmBusqueda As ConsultaEdit
   
    Set frmBusqueda = New ConsultaEdit
  
    mobjBusqueda.ConsultaCampos "vFacturasCompra"
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
        .HeaderString(jgexHFCenter) = "Listado de facturas de compra"
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
    
    Set objFacturaCompra = New FacturaCompra
    objFacturaCompra.Load mlngID, GescomMain.objParametro.Moneda
    If objFacturaCompra.Neto = 0 Then Exit Sub
    ' Si ya está contabilizado hay que:
    ' - preguntar si re-contabilizar todo
    ' - no recontabilizar si se dice que no (por defecto)
    ' - recontabilizar si se dice que si.
    ' - abortar si se pide
    If objFacturaCompra.Contabilizado Then
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
        
    objFacturaCompra.Contabilizar flgForzar
    
    Set objFacturaCompra = Nothing

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
    
    Set objFacturaCompra = New FacturaCompra
    objFacturaCompra.Load mlngID, GescomMain.objParametro.Moneda
    objFacturaCompra.BeginEdit "EUR"
    objFacturaCompra.DESContabilizar
    objFacturaCompra.ApplyEdit
    Set objFacturaCompra = Nothing

End Sub

Private Sub jgrdItems_ColButtonClick(ByVal ColIndex As Integer)
Dim objProveedor As Proveedor
Dim frmProveedorEdit As ProveedorEdit
Dim rdCurrent As JSRowData
Dim lngProveedorID As Long

    If ColIndex = jgrdItems.Columns("NombreProveedor").Index Then
        Set objProveedor = New Proveedor
        Set rdCurrent = jgrdItems.GetRowData(jgrdItems.Row)
        lngProveedorID = rdCurrent.Value(jgrdItems.Columns("ProveedorID").Index)
        objProveedor.Load lngProveedorID
        
        Set frmProveedorEdit = New ProveedorEdit
        frmProveedorEdit.Component objProveedor
        frmProveedorEdit.Show
        
        Set objProveedor = Nothing
        Set frmProveedorEdit = Nothing
    End If
    
End Sub
