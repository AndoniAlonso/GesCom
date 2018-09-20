VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EtiquetasEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etiquetas de articulos"
   ClientHeight    =   5895
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EtiquetasEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Etiquetas a imprimir"
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9855
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   8640
         TabIndex        =   4
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   7560
         TabIndex        =   3
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Height          =   375
         Left            =   6480
         TabIndex        =   2
         Top             =   4320
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwEtiquetas 
         Height          =   3855
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   6800
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9000
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
End
Attribute VB_Name = "EtiquetasEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mobjEtiquetas As Etiquetas
Attribute mobjEtiquetas.VB_VarHelpID = -1

Public Sub Component(EtiquetasObject As Etiquetas)

    Set mobjEtiquetas = EtiquetasObject

End Sub

Private Sub cmdCancel_Click()

    On Error GoTo ErrorManager

    mobjEtiquetas.CancelEdit
  
    Unload Me
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjEtiquetas.ApplyEdit
    
    If ShowDialogSave("Fichero de exportación de etiquetas", _
                       ".TXT", "EtiComp.TXT", "Texto (*.TXT)") = vbOK Then
        mobjEtiquetas.FileName = GescomMain.dlgFileSave.FileName
        mobjEtiquetas.WriteSequentialFile
        Unload Me
    Else
        mobjEtiquetas.BeginEdit
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjEtiquetas
        
        Caption = "Etiquetas de artículos"

        .BeginEdit
        
    End With
    
    lvwEtiquetas.SmallIcons = GescomMain.mglIconosPequeños
    
    lvwEtiquetas.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwEtiquetas.ColumnHeaders.Add , , "Artículo - Color", ColumnSize(15)
    lvwEtiquetas.ColumnHeaders.Add , , "36", ColumnSize(4)
    lvwEtiquetas.ColumnHeaders.Add , , "38", ColumnSize(4)
    lvwEtiquetas.ColumnHeaders.Add , , "40", ColumnSize(4)
    lvwEtiquetas.ColumnHeaders.Add , , "42", ColumnSize(4)
    lvwEtiquetas.ColumnHeaders.Add , , "44", ColumnSize(4)
    lvwEtiquetas.ColumnHeaders.Add , , "46", ColumnSize(4)
    lvwEtiquetas.ColumnHeaders.Add , , "48", ColumnSize(4)
    lvwEtiquetas.ColumnHeaders.Add , , "50", ColumnSize(4)
    lvwEtiquetas.ColumnHeaders.Add , , "52", ColumnSize(4)
    lvwEtiquetas.ColumnHeaders.Add , , "54", ColumnSize(4)
    lvwEtiquetas.ColumnHeaders.Add , , "56", ColumnSize(4)
    LoadEtiquetas
  
    mflgLoading = False

End Sub

Private Sub lvwEtiquetas_DblClick()
  
    Call cmdEdit_Click
    
End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
    IsList = False
    
End Function

' a partir de aqui -----> child

Private Sub cmdAdd_Click()
  
    Dim frmEtiqueta As EtiquetaEdit
  
    On Error GoTo ErrorManager
    Set frmEtiqueta = New EtiquetaEdit
    frmEtiqueta.Component mobjEtiquetas.Add
    frmEtiqueta.Show vbModal
    LoadEtiquetas
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdEdit_Click()

    Dim frmEtiqueta As EtiquetaEdit
  
    On Error GoTo ErrorManager
    
    If lvwEtiquetas.SelectedItem Is Nothing Then Exit Sub

    Set frmEtiqueta = New EtiquetaEdit
    frmEtiqueta.Component _
        mobjEtiquetas(Val(lvwEtiquetas.SelectedItem.Key))
    frmEtiqueta.Show vbModal
    LoadEtiquetas
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdRemove_Click()

    On Error GoTo ErrorManager
        
    mobjEtiquetas.Remove Val(lvwEtiquetas.SelectedItem.Key)
    LoadEtiquetas
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadEtiquetas()

    Dim objEtiqueta As Etiqueta
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwEtiquetas.ListItems.Clear
    For lngIndex = 1 To mobjEtiquetas.Count
        Set itmList = lvwEtiquetas.ListItems.Add _
            (Key:=Format$(lngIndex) & "K")
        Set objEtiqueta = mobjEtiquetas(lngIndex)

        With itmList
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("Etiqueta").Key
            
            If objEtiqueta.IsDeleted Then .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
            .SubItems(1) = Trim(objEtiqueta.ArticuloColor)
            .SubItems(2) = objEtiqueta.CantidadT36
            .SubItems(3) = objEtiqueta.CantidadT38
            .SubItems(4) = objEtiqueta.CantidadT40
            .SubItems(5) = objEtiqueta.CantidadT42
            .SubItems(6) = objEtiqueta.CantidadT44
            .SubItems(7) = objEtiqueta.CantidadT46
            .SubItems(8) = objEtiqueta.CantidadT48
            .SubItems(9) = objEtiqueta.CantidadT50
            .SubItems(10) = objEtiqueta.CantidadT52
            .SubItems(11) = objEtiqueta.CantidadT54
            .SubItems(12) = objEtiqueta.CantidadT56
        End With

    Next
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

