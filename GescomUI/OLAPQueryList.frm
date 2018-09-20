VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OLAPQueryList 
   Caption         =   "OLAP - On Line analisis processing"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "OLAPQueryList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwItems 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
Attribute VB_Name = "OLAPQueryList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjOLAPQuery As OLAPQuery
Private mlngColumn As Integer

Private mstrWhere As String
Private mintQuery As Integer

Public Sub Component(objComponent As OLAPQuery)
    Dim strCabecera As Variant

    Set mobjOLAPQuery = objComponent
    
    For Each strCabecera In mobjOLAPQuery.ReportColumns
        lvwItems.ColumnHeaders.Add , , strCabecera, 1000
    Next
    
    Me.Caption = mobjOLAPQuery.Titulo
    
    mstrWhere = mobjOLAPQuery.ClausulaWhere
    mintQuery = mobjOLAPQuery.TipoConsulta
    
    Call RefreshListView

End Sub

Private Sub RefreshListView()
    Dim objItem As Collection
    Dim itmList As ListItem
    Dim lngIndex As Long
    Dim lngIndexJ As Long

    For lngIndex = 1 To mobjOLAPQuery.Count
        With objItem
            Set objItem = mobjOLAPQuery.Item(lngIndex)
            Set itmList = lvwItems.ListItems.Add

            With itmList
                .Text = Trim(objItem(1))
                For lngIndexJ = 2 To objItem.Count
                    .SubItems(lngIndexJ - 1) = Trim(objItem(lngIndexJ))
    '                .Icon = GescomMain.mglIconosGrandes.ListImages("Cliente").Key
    '                .SmallIcon = GescomMain.mglIconosPequeños.ListImages("Cliente").Key
                Next
            End With

        End With
    Next
    LV_AutoSizeColumn lvwItems

End Sub

Private Sub Form_Load()
    'lvwItems.Icons = GescomMain.mglIconosGrandes
    'lvwItems.SmallIcons = GescomMain.mglIconosPequeños
    
    Me.Move 0, 0
    LoadImages Me.tlbHerramientas
    
    mlngColumn = 1
    
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    
    ListView_ColumnClick lvwItems, ColumnHeader
    mlngColumn = ColumnHeader.Index
       
End Sub

Public Sub UpdateListView()
    
    On Error GoTo ErrorManager

    lvwItems.ListItems.Clear

    Set mobjOLAPQuery = Nothing
    Set mobjOLAPQuery = New OLAPQuery
    mobjOLAPQuery.Load mintQuery, mstrWhere
    Call RefreshListView
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
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
        Case Is = "Imprimir"
            Imprimir
        Case Is = "Actualizar"
            UpdateListView
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
        Case Is = "Cerrar"
            Unload Me
        'Case Is = "ExportToExcel"
        '    ExportRecordList mrsRecordList
    End Select
    
End Sub

Private Sub Form_Resize()

    ListView_Resize lvwItems, Me

End Sub

Public Sub QuickSearch()
    
    ListviewQuickSearch lvwItems, mlngColumn

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
    
    objPrintClass.Titulo = "Previsión de necesidades de etiquetas de la temporada " & GescomMain.objParametro.TemporadaActual
    
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

