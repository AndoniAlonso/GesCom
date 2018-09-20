VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form ProveedorList 
   Caption         =   "Lista de Proveedores"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ProveedorList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar tlbHerramientas 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8130
      _ExtentX        =   14340
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
      GroupByBoxInfoText=   "Arrastre una columna aqu� para agrupar."
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      HeaderFontSize  =   7.5
      FontSize        =   7.5
      ColumnHeaderHeight=   285
      IntProp1        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "ProveedorList.frx":0CCA
      FormatStyle(2)  =   "ProveedorList.frx":0DF2
      FormatStyle(3)  =   "ProveedorList.frx":0EA2
      FormatStyle(4)  =   "ProveedorList.frx":0F56
      FormatStyle(5)  =   "ProveedorList.frx":102E
      FormatStyle(6)  =   "ProveedorList.frx":10E6
      FormatStyle(7)  =   "ProveedorList.frx":11C6
      ImageCount      =   0
      PrinterProperties=   "ProveedorList.frx":11E6
   End
End
Attribute VB_Name = "ProveedorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmProveedor As ProveedorEdit
Private objProveedor As Proveedor
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
        ' Propiedades de la 1� columna, campo clave.
        .Columns("ProveedorID").ColumnType = jgexIcon
        .Columns("ProveedorID").DefaultIcon = jgrdItems.GridImages(1).Index
        .Columns("ProveedorID").Caption = vbNullString
        .Columns("ProveedorID").Visible = True
        .Columns("ProveedorID").ColPosition = 1
        .Columns("ProveedorID").Width = 330

        FormatoJColumn .Columns("Nombre"), 2, "Proveedor", , ColumnSize(20)
        FormatoJColumn .Columns("DNINIF"), 3, "DNI/NIF", , ColumnSize(12)
        FormatoJColumn .Columns("Calle"), 4, "Direcci�n", , ColumnSize(20)
        FormatoJColumn .Columns("Poblacion"), 5, "Poblaci�n", , ColumnSize(10)
        FormatoJColumn .Columns("Provincia"), 6, "Provincia", , ColumnSize(10)
        FormatoJColumn .Columns("Telefono1"), 7, "Tel�fono", , ColumnSize(8)
        FormatoJColumn .Columns("CuentaContable"), 8, "Cuenta contable", , ColumnSize(8)
        FormatoJColumn .Columns("NombreAbreviado"), 9, "Medio de pago", , ColumnSize(3)
        FormatoJColumn .Columns("DescFormaPago"), 10, "Forma de pago", , ColumnSize(8)
        FormatoJColumn .Columns("NombreEntidad"), 11, "Banco", , ColumnSize(8)
        FormatoJColumn .Columns("NombreTransportista"), 12, "Transportista", , ColumnSize(10)
        FormatoJColumn .Columns("Codigo"), 13, "C�d.", , ColumnSize(3)
    End With
    
    If strLayout <> vbNullString Then jgrdItems.LoadLayoutString strLayout

End Sub

Private Sub Form_Load()
Dim objButton As Button

    Me.Move 0, 0
    Me.WindowState = vbMaximized
    
    jgrdItems.View = jgexTable

    LoadImages Me.tlbHerramientas

    ' A�adimos los botones especificos de esta opci�n:
    ' - Exportar los datos a contawin.
    ' - Exportar los datos a CSV.
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Contawin", , tbrDefault, GescomMain.mglIconosPeque�os.ListImages("Contawin").Key)
    objButton.ToolTipText = "Exportar datos de clientes a Contawin"
            
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "ExportToCSV", , tbrDefault, GescomMain.mglIconosPeque�os.ListImages("Proveedor").Key)
    objButton.ToolTipText = "Exportar datos de proveedores a fichero CSV"
            
    Set mobjBusqueda = New Consulta
    
    mlngColumn = 1
    

    jgrdItems.ImageHeight = 16
    jgrdItems.ImageWidth = 16
    jgrdItems.GridImages.Add GescomMain.mglIconosPeque�os.ListImages("Proveedor").Picture
    
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
    'txtQuickSearch.ToolTipText = "B�squeda r�pida en " & Column.Caption

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
    
    ' aqu� hay que avisar de si realmente queremos borrar
    Respuesta = MostrarMensaje(MSG_DELETE)
    
    If Respuesta = vbYes Then
        ' Como este proceso puede ser lento muestro el puntero de reloj de arena
        Screen.MousePointer = vbHourglass
        For Each simTemp In jgrdItems.SelectedItems
            If simTemp.RowType = jgexRowTypeRecord Then
                Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
                mlngID = RowData(1)
                If mlngID > 0 Then
                    Set objProveedor = New Proveedor
                    objProveedor.Load mlngID
                    objProveedor.BeginEdit
                    objProveedor.Delete
                    objProveedor.ApplyEdit
                    Set objProveedor = Nothing
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
'    Dim strClausulaWhere As String
    
    On Error GoTo ErrorManager



    Set objRecordList = New RecordList
    '''???? para liberar memoria
    mrsRecordList.Close
    Set mrsRecordList = Nothing
    Set mrsRecordList = objRecordList.Load("Select * from vProveedores", strWhere)
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
    
    ' aqu� hay que avisar de si realmente queremos abrirlos todos
    ' si el n�mero es mayor que 5
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
                Set frmProveedor = New ProveedorEdit
                Set objProveedor = New Proveedor
                objProveedor.Load mlngID
                frmProveedor.Component objProveedor
                frmProveedor.Show
                Set objProveedor = Nothing
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub NewObject()

    Set frmProveedor = New ProveedorEdit
    Set objProveedor = New Proveedor
    frmProveedor.Component objProveedor
    frmProveedor.Show
    Set objProveedor = Nothing

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
        Case Is = "IconosPeque�os"
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
        Case Is = "Contawin"
            Contawin
        Case Is = "ExportToCSV"
            ExportToCSV
    End Select
    
End Sub

Private Sub Form_Resize()

    GridEX_Resize jgrdItems, Me ', frmFiltro

End Sub

Public Sub QuickSearch()
    
    JanusQuickSearch jgrdItems, mlngColumn

End Sub

Public Sub ResultSearch()
    Dim frmBusqueda As ConsultaEdit
   
    Set frmBusqueda = New ConsultaEdit
  
    mobjBusqueda.ConsultaCampos "Proveedores"
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

    jgrdItems.PrinterProperties.FooterString(jgexHFRight) = "P�gina " & PageNumber & vbCrLf & " de " & nPages
    
End Sub


Public Sub Imprimir()
    On Error GoTo ErrorManager
    
    With jgrdItems.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = False
        .TranslateColors = True
        .HeaderString(jgexHFCenter) = "Listado de Proveedores"
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


Public Sub Contawin()
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
'    Dim i As Integer
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorManager
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    For Each simTemp In jgrdItems.SelectedItems
        If simTemp.RowType = jgexRowTypeRecord Then
            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
            mlngID = RowData(1)
            If mlngID > 0 Then
                Set objProveedor = New Proveedor
                objProveedor.Load mlngID
                objProveedor.ExportarContawin GescomMain.objParametro.ServidorContawin, GescomMain.objParametro.EmpresaActualID
                Set objProveedor = Nothing
            End If
        End If
    Next
    Screen.MousePointer = vbDefault

    ' aqu� hay que avisar de que la contabilidad ha ido OK
    Respuesta = MostrarMensaje(MSG_PROCESO_OK)
    
    UpdateListView
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub


Public Sub ExportToCSV()

    MsgBox jgrdItems.ADORecordset.GetString(adClipString, -1, ",", vbCrLf, "(NULL)")
    
End Sub
