VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ConsultaList 
   Caption         =   "Lista de Consultas"
   ClientHeight    =   5175
   ClientLeft      =   3885
   ClientTop       =   3435
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ConsultaList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwItems 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8070
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
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbHerramientas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "ConsultaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjConsultas As Consultas
Private mlngID As Long
Private mlngColumn As Integer

Private frmConsulta As ConsultaEdit
Private mobjConsulta As Consulta

Public Sub Component(objComponent As Consultas)

    Set mobjConsultas = objComponent
    Call RefreshListView

End Sub

Public Function ConsultaSeleccionada() As Consulta

    Set ConsultaSeleccionada = mobjConsulta

End Function

Private Sub RefreshListView()
    Dim objItem As ConsultaDisplay
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    For lngIndex = 1 To mobjConsultas.Count
        With objItem
            Set objItem = mobjConsultas.Item(lngIndex)
            Set itmList = _
                lvwItems.ListItems.Add(Key:= _
                Format$(objItem.ConsultaID) & " K")

            With itmList
                .Text = objItem.Nombre
        
                .Icon = GescomMain.mglIconosGrandes.ListImages("Buscar").Key
                .SmallIcon = GescomMain.mglIconosPequeños.ListImages("Buscar").Key
            End With
    
        End With
    Next

End Sub

Private Sub Form_Load()
    
    Me.Move 0, 0
    lvwItems.ColumnHeaders.Add , , "Nombre", ColumnSize(50)

    lvwItems.Icons = GescomMain.mglIconosGrandes
    lvwItems.SmallIcons = GescomMain.mglIconosPequeños
    
    LoadImages Me.tlbHerramientas
    
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
    Dim Respuesta As VbMsgBoxResult
    Dim i As Integer

    On Error GoTo ErrorManager

    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    ' aquí avisamos de si realmente queremos borrar
    Respuesta = MostrarMensaje(MSG_DELETE)
    
    If Respuesta = vbYes Then
        For i = lvwItems.ListItems.Count To 1 Step -1
            If lvwItems.ListItems(i).Selected = True Then
                mlngID = Val(lvwItems.ListItems(i).Key)
                If mlngID > 0 Then
                    Set mobjConsulta = New Consulta
                    mobjConsulta.Load mlngID
                    mobjConsulta.BeginEdit
                    mobjConsulta.Delete
                    mobjConsulta.ApplyEdit
                    Set mobjConsulta = Nothing
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

    On Error GoTo ErrorManager

    lvwItems.ListItems.Clear

    Set mobjConsultas = Nothing
    Set mobjConsultas = New Consultas
    mobjConsultas.Load strWhere
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
            SelectFirstItem
        End If
    Else
        SelectFirstItem
    End If

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub SelectFirstItem()

    Dim i As Integer
    
    ' Seleccionamos el primer item seleccionado, y lo dejamos en la variable
       
    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Selected = True Then
            mlngID = Val(lvwItems.ListItems(i).Key)
            If mlngID > 0 Then
                Set frmConsulta = New ConsultaEdit
                Set mobjConsulta = New Consulta
                mobjConsulta.Load mlngID
                Me.Hide
                
                Exit For
            End If
        End If
    Next i
    
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
        Case Is = "Abrir"
            EditSelected
        Case Is = "Eliminar"
            DeleteSelected
        Case Is = "IconosGrandes"
            SetListViewStyle (lvwIcon)
        Case Is = "IconosPequeños"
            SetListViewStyle (lvwSmallIcon)
        Case Is = "Lista"
            SetListViewStyle (lvwList)
        Case Is = "Detalle"
            SetListViewStyle (lvwReport)
        Case Is = "Cerrar"
            Me.Hide
    End Select
    
End Sub

