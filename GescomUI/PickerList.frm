VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PickerList 
   Caption         =   "Documentos"
   ClientHeight    =   6255
   ClientLeft      =   3885
   ClientTop       =   3435
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PickerList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSelectedItems 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox txtTotalItems 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Aña&dir"
      Height          =   375
      Left            =   9000
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "El&iminar"
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de Documentos"
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.TextBox txtQuickSearch 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   2775
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4895
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwSelectedItems 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4260
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
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CheckBox chkCabecera 
      Caption         =   "I&ncorporar datos de cabecera"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Total seleccionados"
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Total artículos"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   6000
      Width           =   1095
   End
End
Attribute VB_Name = "PickerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrTitulo As String
Private mbolCabecera As Boolean
Private mobjSelectedItems As PickerItems
Private mlngID As Long
Private mlngColumn As Integer
Private mlngNuevaBusqueda As Long
Private mdblCantidadItems As Double
Private mdblCantidadSelected As Double

'Private objPicker As PickerItemDisplay

Public Property Let Titulo(Value As String)
   
    mstrTitulo = Trim(Value)
   
End Property

Public Property Get SelectedItems() As PickerItems
   
    Set SelectedItems = mobjSelectedItems
   
End Property

Public Sub LoadData(ByVal strTabla As String, ByVal Propietario As Long, _
    ByVal Empresa As Long, ByVal Temporada As Long)

    Dim objPickerItems As PickerItems
    Dim objItem As PickerItemDisplay
    Dim itmList As ListItem
    Dim lngIndex As Long
    
  
    Set objPickerItems = New PickerItems
  
    objPickerItems.Load strTabla, Propietario, Empresa, Temporada

    For lngIndex = 1 To objPickerItems.Count
        With objItem
            Set objItem = objPickerItems.Item(lngIndex)
            Set itmList = _
                lvwItems.ListItems.Add(Key:= _
                Format$(objItem.DocumentoID) & " K")
    
            With itmList
                .Text = objItem.Numero
                .SubItems(1) = objItem.Nombre
                .SubItems(2) = objItem.Descripcion
                .SubItems(3) = objItem.Cantidad
                .SubItems(4) = Format(objItem.Fecha, "yyyy/mm/dd")
                
                '.Icon = GescomMain.mglIconosGrandes.ListImages("Documento").Key
                .SmallIcon = GescomMain.mglIconosPequeños.ListImages("Documento").Key
            End With

            mdblCantidadItems = mdblCantidadItems + objItem.Cantidad
        End With
    Next
    
    txtTotalItems = mdblCantidadItems
  
    Set objPickerItems = Nothing
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    Me.Hide
    
End Sub

Private Sub cmdSeleccionar_Click()
    'Dim objDisplay As PickerItemDisplay
    Dim i As Integer
    Dim itmList As ListItem

    If lvwItems.SelectedItem Is Nothing Then Exit Sub

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    Set mobjSelectedItems = New PickerItems

    ' Añadimos en el listview de seleccionados.
    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Selected = True Then
            mlngID = Val(lvwItems.ListItems(i).Key)
            If mlngID > 0 Then
                Set itmList = lvwSelectedItems.ListItems.Add(Key:=lvwItems.ListItems(i).Key)
                With itmList
                    .Text = lvwItems.ListItems(i).Text
                    .SubItems(1) = lvwItems.ListItems(i).SubItems(1)
                    .SubItems(2) = lvwItems.ListItems(i).SubItems(2)
                    .SubItems(3) = lvwItems.ListItems(i).SubItems(3)
                    .SubItems(4) = lvwItems.ListItems(i).SubItems(4)
                    .SmallIcon = GescomMain.mglIconosPequeños.ListImages("Documento").Key
                End With
                mdblCantidadSelected = mdblCantidadSelected + lvwItems.ListItems(i).SubItems(3)
            End If
        End If
    Next i
        
    ' Eliminamos para evitar multiples selecciones
    For i = lvwItems.ListItems.Count To 1 Step -1
        If lvwItems.ListItems(i).Selected = True Then
            mdblCantidadItems = mdblCantidadItems - lvwItems.ListItems(i).SubItems(3)
            lvwItems.ListItems.Remove (i)
        End If
    Next i
    
    Screen.MousePointer = vbDefault
    lvwItems.Refresh
    lvwSelectedItems.Refresh
    
    txtTotalItems.Text = mdblCantidadItems
    txtSelectedItems.Text = mdblCantidadSelected
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub lvwItems_DblClick()
  
    Call cmdSeleccionar_Click
    
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        Call cmdSeleccionar_Click
        txtQuickSearch.SetFocus
    ElseIf KeyCode = 27 Then
        txtQuickSearch.SetFocus
    End If

End Sub

Private Sub cmdEliminar_Click()
    'Dim objDisplay As PickerItemDisplay
    Dim i As Integer
    Dim itmList As ListItem

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    Set mobjSelectedItems = New PickerItems

    If lvwSelectedItems.SelectedItem Is Nothing Then Exit Sub

    ' Añadimos en el listview de seleccionados.
    For i = 1 To lvwSelectedItems.ListItems.Count
        If lvwSelectedItems.ListItems(i).Selected = True Then
            mlngID = Val(lvwSelectedItems.ListItems(i).Key)
            If mlngID > 0 Then
                Set itmList = lvwItems.ListItems.Add(Key:=lvwSelectedItems.ListItems(i).Key)
                With itmList
                    .Text = lvwSelectedItems.ListItems(i).Text
                    .SubItems(1) = lvwSelectedItems.ListItems(i).SubItems(1)
                    .SubItems(2) = lvwSelectedItems.ListItems(i).SubItems(2)
                    .SubItems(3) = lvwSelectedItems.ListItems(i).SubItems(3)
                    .SubItems(4) = lvwSelectedItems.ListItems(i).SubItems(4)
                    .SmallIcon = GescomMain.mglIconosPequeños.ListImages("Documento").Key
                End With
                mdblCantidadItems = mdblCantidadItems + lvwSelectedItems.ListItems(i).SubItems(3)
            End If
        End If
    Next i
        
    ' Eliminamos para evitar multiples selecciones
    For i = lvwSelectedItems.ListItems.Count To 1 Step -1
        If lvwSelectedItems.ListItems(i).Selected = True Then
            mdblCantidadSelected = mdblCantidadSelected - lvwSelectedItems.ListItems(i).SubItems(3)
            lvwSelectedItems.ListItems.Remove (i)
        End If
    Next i
    
    Screen.MousePointer = vbDefault
    lvwItems.Refresh
    lvwSelectedItems.Refresh
    
    txtTotalItems.Text = mdblCantidadItems
    txtSelectedItems.Text = mdblCantidadSelected
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdOK_Click()
    Dim objDisplay As PickerItemDisplay
    Dim i As Integer

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    Set mobjSelectedItems = New PickerItems

    For i = 1 To lvwSelectedItems.ListItems.Count
        mlngID = Val(lvwSelectedItems.ListItems(i).Key)
        If mlngID > 0 Then
            Set objDisplay = New PickerItemDisplay
            objDisplay.DocumentoID = Val(lvwSelectedItems.ListItems(i).Key)
            objDisplay.Numero = lvwSelectedItems.ListItems(i).Text
            objDisplay.Nombre = lvwSelectedItems.ListItems(i).SubItems(1)
            objDisplay.Descripcion = lvwSelectedItems.ListItems(i).SubItems(2)
            objDisplay.Cantidad = lvwSelectedItems.ListItems(i).SubItems(3)
            objDisplay.Fecha = lvwSelectedItems.ListItems(i).SubItems(4)

            mobjSelectedItems.AddPickerItemDisplay objDisplay
            Set objDisplay = Nothing
        End If
    Next i

    Me.Hide
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub chkCabecera_Click()
   
    mbolCabecera = chkCabecera.Value
    
End Sub

Private Sub Form_Load()
    
    DisableX Me
    
    lvwItems.ColumnHeaders.Add , , "Documento", ColumnSize(8)
    lvwItems.ColumnHeaders.Add , , "Nombre", ColumnSize(10)
    lvwItems.ColumnHeaders.Add , , "Descripción", ColumnSize(28)
    lvwItems.ColumnHeaders.Add , , "Cantidad", ColumnSize(7), vbRightJustify
    lvwItems.ColumnHeaders.Add , , "Fecha", ColumnSize(8)
    
    'lvwItems.Icons = GescomMain.mglIconosGrandes
    lvwItems.SmallIcons = GescomMain.mglIconosPequeños

    lvwSelectedItems.ColumnHeaders.Add , , "Documento", ColumnSize(8)
    lvwSelectedItems.ColumnHeaders.Add , , "Nombre", ColumnSize(10)
    lvwSelectedItems.ColumnHeaders.Add , , "Descripción", ColumnSize(28)
    lvwSelectedItems.ColumnHeaders.Add , , "Cantidad", ColumnSize(7), vbRightJustify
    lvwSelectedItems.ColumnHeaders.Add , , "Fecha", ColumnSize(8)
    
    'lvwSelectedItems.Icons = GescomMain.mglIconosGrandes
    lvwSelectedItems.SmallIcons = GescomMain.mglIconosPequeños

    mlngNuevaBusqueda = 0
    mlngColumn = 2
    'txtQuickSearch.SetFocus
    
    mdblCantidadItems = 0
    mdblCantidadSelected = 0

End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    
    ListView_ColumnClick lvwItems, ColumnHeader
    mlngColumn = ColumnHeader.Index
       
End Sub

Public Sub SetListViewStyle(View As Integer)
    
    lvwItems.View = View
   
End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
    IsList = False
    
End Function

Private Sub lvwSelectedItems_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 Then
        Call cmdEliminar_Click
    End If

End Sub

Private Sub txtQuickSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        mlngNuevaBusqueda = 1
        QuickSearch
    End If

End Sub

Private Sub txtQuickSearch_GotFocus()

    SelTextBox txtQuickSearch
        
End Sub


Public Sub QuickSearch()
    Dim strDato As String
    Dim i As Integer

' Búsqueda rápida en los campos del listview
    For i = 1 To lvwItems.ListItems.Count
        lvwItems.ListItems(i).Selected = False
    Next i

    If mlngNuevaBusqueda > lvwItems.ListItems.Count Then
        mlngNuevaBusqueda = 1
    End If

    For i = mlngNuevaBusqueda To lvwItems.ListItems.Count
        If mlngColumn = 1 Then
            strDato = lvwItems.ListItems(i).Text
        Else
            strDato = lvwItems.ListItems(i).SubItems(mlngColumn - 1)
        End If
        mlngNuevaBusqueda = i + 1
        If InStr(1, strDato, txtQuickSearch, vbTextCompare) Then
            lvwItems.ListItems(i).EnsureVisible
            Set lvwItems.SelectedItem = lvwItems.ListItems(i)
            lvwItems.SetFocus
            Exit Sub
        End If
    Next i


End Sub

