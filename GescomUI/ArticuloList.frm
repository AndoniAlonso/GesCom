VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "gridex20_b.ocx"
Object = "{E51AF47C-BB3B-48B6-A74A-7DA1722D2C68}#3.0#0"; "EntityProxy.ocx"
Begin VB.Form ArticuloList 
   Caption         =   "Lista de Artículos"
   ClientHeight    =   6945
   ClientLeft      =   3885
   ClientTop       =   3435
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ArticuloList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFiltro 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   10815
      Begin EntityProxy.ctlEntityProxy epTemporada 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3870
         _ExtentX        =   7858
         _ExtentY        =   873
      End
      Begin EntityProxy.ctlEntityProxy epProveedor 
         Height          =   495
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   3870
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
      Width           =   10695
      _ExtentX        =   18865
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
      ColumnsCount    =   1
      Column(1)       =   "ArticuloList.frx":0E42
      FormatStylesCount=   7
      FormatStyle(1)  =   "ArticuloList.frx":0FEA
      FormatStyle(2)  =   "ArticuloList.frx":1112
      FormatStyle(3)  =   "ArticuloList.frx":11C2
      FormatStyle(4)  =   "ArticuloList.frx":1276
      FormatStyle(5)  =   "ArticuloList.frx":134E
      FormatStyle(6)  =   "ArticuloList.frx":1406
      FormatStyle(7)  =   "ArticuloList.frx":14E6
      ImageCount      =   0
      PrinterProperties=   "ArticuloList.frx":1506
   End
End
Attribute VB_Name = "ArticuloList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmArticulo As ArticuloEdit
Private objArticulo As Articulo
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
        .Columns("ArticuloID").ColumnType = jgexIcon
        .Columns("ArticuloID").DefaultIcon = jgrdItems.GridImages(1).Index
        .Columns("ArticuloID").Caption = vbNullString
        .Columns("ArticuloID").Visible = True
        .Columns("ArticuloID").ColPosition = 1
        .Columns("ArticuloID").Width = 330

        FormatoJColumn .Columns("Nombre"), 2, "Artículo", , ColumnSize(10)
        FormatoJColumn .Columns("NombrePrenda"), 3, "Prenda", , ColumnSize(10)
        FormatoJColumn .Columns("NombreModelo"), 4, "Modelo", , ColumnSize(10)
        FormatoJColumn .Columns("NombreSerie"), 5, "Serie", , ColumnSize(10)
        FormatoJColumn .Columns("StockActual"), 6, "Stock", , ColumnSize(8)
        FormatoJColumn .Columns("StockPendiente"), 7, "Pendiente", , ColumnSize(8)
        FormatoJColumn .Columns("PrecioCosteEUR"), 8, "Precio coste", , ColumnSize(10)
        FormatoJColumn .Columns("PrecioVentaEUR"), 9, "Precio Venta", , ColumnSize(10)
        FormatoJColumn .Columns("PrecioCompraEUR"), 10, "Precio Compra", , ColumnSize(10)
        FormatoJColumn .Columns("PrecioVentaPublico"), 11, "PVP", , ColumnSize(10)
        FormatoJColumn .Columns("NombreTallaje"), 12, "Tallaje", , ColumnSize(15)
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
    ' - Recalcular los precios de coste.
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Recalcular", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Recalcular").Key)
    objButton.ToolTipText = "Recalcular los precios de coste de los artículos seleccionados"
   
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "RecalcularVenta", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Remesa").Key)
    objButton.ToolTipText = "Recalcular los precios de venta de los artículos seleccionados"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "RecalcularPVP", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("PVP").Key)
    objButton.ToolTipText = "Recalcular los PVP de los artículos seleccionados"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "FichaArticulo", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("PrintDocument").Key)
    objButton.ToolTipText = "Imprimir ficha de artículos"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "DetalleArticulo", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Articulo").Key)
    objButton.ToolTipText = "Ficha de artículos"
   
    Set mobjBusqueda = New Consulta
     
    mlngColumn = 1

    ' Criterios de seleccion, filtros

    epTemporada.Initialize 1, "Temporadas", "TemporadaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.objParametro.TemporadaActual, 0
    epTemporada.LoadControl "Temporada"
    
    epProveedor.Initialize 1, "Proveedores", "ProveedorID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epProveedor.LoadControl "Proveedor"
    
    jgrdItems.ImageHeight = 16
    jgrdItems.ImageWidth = 16
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("Articulo").Picture
    
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
                    Set objArticulo = New Articulo
                    objArticulo.Load mlngID
                    objArticulo.BeginEdit
                    objArticulo.Delete
                    objArticulo.ApplyEdit
                    Set objArticulo = Nothing
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

    If epProveedor.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epProveedor.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epProveedor.ClausulaWhere
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
    Set mrsRecordList = objRecordList.Load("Select * from vArticulos", strClausulaWhere)
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
                Set frmArticulo = New ArticuloEdit
                Set objArticulo = New Articulo
                objArticulo.Load mlngID
                frmArticulo.Component objArticulo
                frmArticulo.Show
                Set objArticulo = Nothing
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub NewObject()

    Set frmArticulo = New ArticuloEdit
    Set objArticulo = New Articulo
    frmArticulo.Component objArticulo
    frmArticulo.Show
    Set objArticulo = Nothing

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
        Case Is = "Recalcular"
            Recalcular
        Case Is = "RecalcularVenta"
            RecalcularVenta
        Case Is = "RecalcularPVP"
            RecalcularPVP
        Case Is = "FichaArticulo"
            FichaArticulo
        Case Is = "DetalleArticulo"
            DetalleArticulo
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
  
    mobjBusqueda.ConsultaCampos "vArticulos"
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


Public Sub Imprimir()
    On Error GoTo ErrorManager
    
    With jgrdItems.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = False
        .TranslateColors = True
        .HeaderString(jgexHFCenter) = "Listado de Series"
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

Public Sub Recalcular()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    Respuesta = MostrarMensaje(MSG_RECALCULAR_ARTICULO)
    
    If Respuesta = vbYes Then
        RecalcularItems
    End If

    Exit Sub
    
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub RecalcularItems()
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
                Set objArticulo = New Articulo
                objArticulo.Load mlngID
                objArticulo.BeginEdit
                objArticulo.PrecioCoste = objArticulo.CalcularPrecioCoste
                objArticulo.ApplyEdit
                
                Set objArticulo = Nothing
            End If
        End If
    Next
    
    UpdateListView SentenciaSQL
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub RecalcularVenta()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    Respuesta = MostrarMensaje(MSG_RECALCULAR_VENTA)
    
    If Respuesta = vbYes Then
        RecalcularVentaItems
    End If

    Exit Sub
    
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub RecalcularVentaItems()
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
                Set objArticulo = New Articulo
                objArticulo.Load mlngID
                objArticulo.BeginEdit
                objArticulo.PrecioVenta = objArticulo.CalcularPrecioVenta
                'objArticulo.PrecioVentaPublico = objArticulo.CalcularPrecioVentaPublico
                objArticulo.ApplyEdit
                
                Set objArticulo = Nothing
            End If
        End If
    Next
    
    UpdateListView SentenciaSQL
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub RecalcularPVP()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    Respuesta = MostrarMensaje(MSG_RECALCULAR_PVP)
    
    If Respuesta = vbYes Then
        RecalcularPVPItems
    End If

    Exit Sub
    
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub RecalcularPVPItems()
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
                Set objArticulo = New Articulo
                objArticulo.Load mlngID
                objArticulo.BeginEdit
                objArticulo.PrecioVentaPublico = objArticulo.CalcularPrecioVentaPublico
                objArticulo.ApplyEdit
                
                Set objArticulo = Nothing
            End If
        End If
    Next
    
    UpdateListView SentenciaSQL
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub FichaArticulo()
Dim simTemp As JSSelectedItem
Dim RowData As JSRowData
Dim Respuesta As VbMsgBoxResult
Dim objPrintFichaArticulo As PrintFichaArticulo
Dim frmPrintOptions As frmPrint
    
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
                    Set objArticulo = New Articulo
                    Set objPrintFichaArticulo = New PrintFichaArticulo
                    objArticulo.Load mlngID
                    
                    objPrintFichaArticulo.PrinterNumber = frmPrintOptions.PrinterNumber
                    objPrintFichaArticulo.Copies = frmPrintOptions.Copies
                    objPrintFichaArticulo.Component objArticulo
                    
                    objPrintFichaArticulo.PrintObject
                    
                    Set objPrintFichaArticulo = Nothing
                    Set objArticulo = Nothing
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

Public Sub DetalleArticulo()
'    Dim frmConsultaArticulo As Consulta_ArtículoColor
'
'    On Error GoTo ErrorManager
'
'    Set frmConsultaArticulo = New Consulta_ArtículoColor
'    frmConsultaArticulo.Show
'
'    Exit Sub
'
'ErrorManager:
'    ManageErrors (Me.Caption)
End Sub


