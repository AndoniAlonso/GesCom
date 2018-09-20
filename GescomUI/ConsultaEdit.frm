VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ConsultaEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultas"
   ClientHeight    =   3975
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ConsultaEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdConsultaSave 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsultaList 
      Caption         =   "Abrir c&onsulta"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Condiciones de la consulta"
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   6480
         TabIndex        =   4
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   5400
         TabIndex        =   3
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   2640
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwConsultaItems 
         Height          =   2175
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
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
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "ConsultaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjConsulta As Consulta
Attribute mobjConsulta.VB_VarHelpID = -1

Private mstrWhere As String
Public mflgAplicarFiltro As Boolean

''Definición de la API bloquear las teclas
'Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
'Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer

''Definicion de las constantes
'Private Const MF_BYPOSITION = &H400

Public Sub Component(ConsultaObject As Consulta)

    Set mobjConsulta = ConsultaObject

End Sub

Public Function Consulta() As Consulta

    Set Consulta = mobjConsulta

End Function

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager
    
    mobjConsulta.ApplyEdit
    mobjConsulta.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdConsultaSave_Click()
    Dim strConsultaNombre As String
    
    On Error GoTo ErrorManager

    strConsultaNombre = InputBox("Nombre:", "Introducir nombre de la consulta", mobjConsulta.Nombre)

    If Len(strConsultaNombre) = 0 Then Exit Sub
    
    mobjConsulta.Nombre = strConsultaNombre
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjConsulta.ApplyEdit
  
    mobjConsulta.SaveConsulta
    
    mobjConsulta.BeginEdit
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
    Exit Sub
End Sub

Private Sub cmdCancel_Click()

    On Error GoTo ErrorManager

    mobjConsulta.CancelEdit
    mflgAplicarFiltro = False
    Me.Hide
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjConsulta.ApplyEdit
  
    mstrWhere = mobjConsulta.ClausulaWhere
    mflgAplicarFiltro = True
    Me.Hide

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()
    'Dim MenuSistema As Integer
    'Dim Res As Integer
    
    DisableX Me
    
    mflgLoading = True
    With mobjConsulta
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Consulta [(nueva)]"
    
        Else
            Caption = "Consulta"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        'txtNombre = .Nombre
    
        .BeginEdit
    End With
    
    lvwConsultaItems.SmallIcons = GescomMain.mglIconosPequeños
    
    lvwConsultaItems.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwConsultaItems.ColumnHeaders.Add , , "Condicion", ColumnSize(30)
    LoadConsultaItems
      
    mflgLoading = False
    
    'desactivo la x
    'MenuSistema% = GetSystemMenu(hWnd, 0)
    'Res% = RemoveMenu(MenuSistema%, 6, MF_BYPOSITION)

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
'    cmdCancel_Click
    
End Sub

Private Sub lvwConsultaItems_DblClick()
    
    Call cmdEdit_Click
    
End Sub

Private Sub mobjConsulta_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function

' a partir de aqui -----> child

Private Sub cmdAdd_Click()
    Dim frmConsultaItem As ConsultaItemEdit
  
    On Error GoTo ErrorManager
    Set frmConsultaItem = New ConsultaItemEdit
    frmConsultaItem.Component mobjConsulta.ConsultaItems.Add
    frmConsultaItem.Show vbModal
    LoadConsultaItems
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdEdit_Click()

    Dim frmConsultaItem As ConsultaItemEdit
  
    On Error GoTo ErrorManager
    
    If lvwConsultaItems.SelectedItem Is Nothing Then Exit Sub
    
    Set frmConsultaItem = New ConsultaItemEdit
    frmConsultaItem.Component _
        mobjConsulta.ConsultaItems(Val(lvwConsultaItems.SelectedItem.Key))
    frmConsultaItem.Show vbModal
    LoadConsultaItems
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdRemove_Click()

    On Error GoTo ErrorManager
        
    If lvwConsultaItems.SelectedItem Is Nothing Then Exit Sub
    mobjConsulta.ConsultaItems.Remove Val(lvwConsultaItems.SelectedItem.Key)
    LoadConsultaItems
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadConsultaItems()

    Dim objConsultaItem As ConsultaItem
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwConsultaItems.ListItems.Clear
    For lngIndex = 1 To mobjConsulta.ConsultaItems.Count
        Set itmList = lvwConsultaItems.ListItems.Add _
            (Key:=Format$(lngIndex) & "K")
        Set objConsultaItem = mobjConsulta.ConsultaItems(lngIndex)

        With itmList
            'If objConsultaItem.IsNew Then
            '    .Text = "(new)"

            'Else
            '    .Text = objConsultaItem.ConsultaItemID

            'End If
            
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
            
            If objConsultaItem.IsDeleted Then
                lvwConsultaItems.ListItems.Remove (Format$(lngIndex) & "K")
            Else
                .SubItems(1) = objConsultaItem.ClausulaWhere
            End If
        End With

    Next
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

Public Property Get SentenciaSQL() As String

    SentenciaSQL = mstrWhere
    
End Property

Private Sub cmdConsultaList_Click()
    Dim frmList As ConsultaList
    Dim objConsultas As Consultas
    
    On Error GoTo ErrorManager

    Set frmList = New ConsultaList
    Set objConsultas = New Consultas
    objConsultas.Load mobjConsulta.Objeto
    With frmList
        .Component objConsultas
        .Show vbModal
        
        If Not (.ConsultaSeleccionada Is Nothing) Then
            mobjConsulta.CancelEdit
            Set mobjConsulta = frmList.ConsultaSeleccionada
            mobjConsulta.BeginEdit
            LoadConsultaItems
        End If
  
        Unload frmList
    End With

    Set objConsultas = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)

End Sub


