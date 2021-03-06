VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form PrendaList 
   Caption         =   "Lista de Prendas"
   ClientHeight    =   5175
   ClientLeft      =   3885
   ClientTop       =   3435
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PrendaList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar tlbHerramientas 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
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
      Top             =   960
      Width           =   7815
      _ExtentX        =   13785
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
      FormatStylesCount=   6
      FormatStyle(1)  =   "PrendaList.frx":1042
      FormatStyle(2)  =   "PrendaList.frx":116A
      FormatStyle(3)  =   "PrendaList.frx":121A
      FormatStyle(4)  =   "PrendaList.frx":12CE
      FormatStyle(5)  =   "PrendaList.frx":13A6
      FormatStyle(6)  =   "PrendaList.frx":145E
      ImageCount      =   0
      PrinterProperties=   "PrendaList.frx":153E
   End
End
Attribute VB_Name = "PrendaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmprenda As PrendaEdit
Private objprenda As Prenda
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
        .Columns("PrendaID").ColumnType = jgexIcon
        .Columns("PrendaID").DefaultIcon = jgrdItems.GridImages(1).Index
        .Columns("PrendaID").Caption = vbNullString
        .Columns("PrendaID").Visible = True
        .Columns("PrendaID").ColPosition = 1
        .Columns("PrendaID").Width = 330

        FormatoJColumn .Columns("Nombre"), 2, "Nombre"
        FormatoJColumn .Columns("Codigo"), 3, "C�digo"
        
   End With
    
   If strLayout <> vbNullString Then jgrdItems.LoadLayoutString strLayout
    
End Sub

Private Sub Form_Load()
'Dim objButton As Button
    
    Me.Move 0, 0
    Me.WindowState = vbMaximized
        
    jgrdItems.View = jgexTable
    
    LoadImages Me.tlbHerramientas
    
    
    Set mobjBusqueda = New Consulta
    
    mlngColumn = 1
    

    jgrdItems.ImageHeight = 16
    jgrdItems.ImageWidth = 16
    jgrdItems.GridImages.Add GescomMain.mglIconosPeque�os.ListImages("Prenda").Picture
    
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
                    Set objprenda = New Prenda
                    objprenda.Load mlngID, GescomMain.objParametro.Moneda
                    objprenda.BeginEdit GescomMain.objParametro.Moneda
                    objprenda.Delete
                    objprenda.ApplyEdit
                    Set objprenda = Nothing
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
    Set mrsRecordList = objRecordList.Load("Select * from Prendas", strWhere)
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
            Set frmprenda = New PrendaEdit
            Set objprenda = New Prenda
            objprenda.Load mlngID, GescomMain.objParametro.Moneda
            frmprenda.Component objprenda
            frmprenda.Show
            Set objprenda = Nothing
        End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub NewObject()

    Set frmprenda = New PrendaEdit
    Set objprenda = New Prenda
    frmprenda.Component objprenda
    frmprenda.Show
    Set objprenda = Nothing

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
  
    mobjBusqueda.ConsultaCampos "Prendas"
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
        .HeaderString(jgexHFCenter) = "Listado de Prendas"
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

    

    
