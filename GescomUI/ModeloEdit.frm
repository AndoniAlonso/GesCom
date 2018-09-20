VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ModeloEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modelos"
   ClientHeight    =   7935
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ModeloEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Modelo"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   700
         Width           =   5415
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   340
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   5880
      TabIndex        =   23
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Propiedades"
      Height          =   6015
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   6735
      Begin VB.TextBox txtBeneficioPVP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4200
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtCantidadTela 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4200
         TabIndex        =   13
         Top             =   705
         Width           =   1335
      End
      Begin VB.TextBox txtCorte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   705
         Width           =   1335
      End
      Begin VB.TextBox txtTaller 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   1065
         Width           =   1335
      End
      Begin VB.TextBox txtBeneficio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   340
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Líneas del Modelo"
         Height          =   4455
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   6255
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Aña&dir"
            Height          =   375
            Left            =   2880
            TabIndex        =   18
            Top             =   3960
            Width           =   975
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Editar"
            Height          =   375
            Left            =   3960
            TabIndex        =   19
            Top             =   3960
            Width           =   975
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "El&iminar"
            Height          =   375
            Left            =   5040
            TabIndex        =   20
            Top             =   3960
            Width           =   975
         End
         Begin MSComctlLib.ListView lvwEstrModelos 
            Height          =   3615
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   6376
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "% Beneficio PVP"
         Height          =   195
         Left            =   2880
         TabIndex        =   8
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "% Beneficio"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Tela"
         Height          =   195
         Left            =   3000
         TabIndex        =   12
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Precio Corte"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Precio Taller"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   870
      End
   End
End
Attribute VB_Name = "ModeloEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjModelo As Modelo
Attribute mobjModelo.VB_VarHelpID = -1

Public Sub Component(ModeloObject As Modelo)

    Set mobjModelo = ModeloObject

End Sub

Private Sub cmdApply_Click()
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorManager

    Respuesta = MostrarMensaje(MSG_MODIF_ARTICULO)
    
    mobjModelo.ApplyEdit
    mobjModelo.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    Dim Respuesta As VbMsgBoxResult
    
    If mobjModelo.IsDirty And Not mobjModelo.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjModelo.CancelEdit
            Unload Me
        End If
    Else
        mobjModelo.CancelEdit
        Unload Me
    End If

End Sub

 Private Sub cmdOK_Click()
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorManager

    Respuesta = MostrarMensaje(MSG_MODIF_ARTICULO)
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjModelo.ApplyEdit
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjModelo
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Modelo [(nuevo)]"
    
        Else
            Caption = "Modelo [" & .Nombre & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNombre = .Nombre
        txtCodigo = .Codigo
        txtBeneficio = .Beneficio
        txtBeneficioPVP = .BeneficioPVP
        txtCantidadTela = .CantidadTela
        txtCorte = .Corte
        txtTaller = .Taller
    
        .BeginEdit
    
        If .IsNew Then .TemporadaID = GescomMain.objParametro.TemporadaActualID
    
    End With
    
    lvwEstrModelos.SmallIcons = GescomMain.mglIconosPequeños
    
    lvwEstrModelos.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwEstrModelos.ColumnHeaders.Add , , "Material", ColumnSize(10)
    lvwEstrModelos.ColumnHeaders.Add , , "Cantidad", ColumnSize(6), vbRightJustify
    lvwEstrModelos.ColumnHeaders.Add , , "Precio", ColumnSize(8), vbRightJustify
    lvwEstrModelos.ColumnHeaders.Add , , "Observaciones", ColumnSize(20)
    LoadEstrModelos
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub lvwEstrModelos_DblClick()
    
    Call cmdEdit_Click
    
End Sub

Private Sub mobjModelo_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtBeneficio_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtBeneficio
        
End Sub

Private Sub txtBeneficioPVP_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtBeneficioPVP
        
End Sub

Private Sub txtCantidadTela_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidadTela
        
End Sub

Private Sub txtCodigo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCodigo
        
End Sub

Private Sub txtCorte_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCorte
        
End Sub

Private Sub txtNombre_Change()

    If Not mflgLoading Then _
        TextChange txtNombre, mobjModelo, "Nombre"

End Sub

Private Sub txtNombre_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNombre
        
End Sub

Private Sub txtNombre_LostFocus()

    txtNombre = TextLostFocus(txtNombre, mobjModelo, "Nombre")

End Sub

Private Sub txtCodigo_Change()

    If Not mflgLoading Then _
        TextChange txtCodigo, mobjModelo, "Codigo"

End Sub

Private Sub txtCodigo_LostFocus()

    txtCodigo = TextLostFocus(txtCodigo, mobjModelo, "Codigo")

End Sub

Private Sub txtBeneficio_Change()

    If Not mflgLoading Then _
        TextChange txtBeneficio, mobjModelo, "Beneficio"

End Sub

Private Sub txtBeneficio_LostFocus()

    txtBeneficio = TextLostFocus(txtBeneficio, mobjModelo, "Beneficio")

End Sub

Private Sub txtBeneficioPVP_Change()

    If Not mflgLoading Then _
        TextChange txtBeneficioPVP, mobjModelo, "BeneficioPVP"

End Sub

Private Sub txtBeneficioPVP_LostFocus()

    txtBeneficioPVP = TextLostFocus(txtBeneficioPVP, mobjModelo, "BeneficioPVP")

End Sub

Private Sub txtCantidadTela_Change()

    If Not mflgLoading Then _
        TextChange txtCantidadTela, mobjModelo, "CantidadTela"

End Sub

Private Sub txtCantidadTela_LostFocus()

    txtCantidadTela = TextLostFocus(txtCantidadTela, mobjModelo, "CantidadTela")

End Sub

Private Sub txtCorte_Change()

    If Not mflgLoading Then _
        TextChange txtCorte, mobjModelo, "Corte"

End Sub

Private Sub txtCorte_LostFocus()

    txtCorte = TextLostFocus(txtCorte, mobjModelo, "Corte")

End Sub

Private Sub txtTaller_Change()

    If Not mflgLoading Then _
        TextChange txtTaller, mobjModelo, "Taller"

End Sub

Private Sub txtTaller_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtTaller
        
End Sub

Private Sub txtTaller_LostFocus()

    txtTaller = TextLostFocus(txtTaller, mobjModelo, "Taller")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function

' a partir de aqui -----> child

Private Sub cmdAdd_Click()
    
    Dim frmEstrModelo As EstrModeloEdit
  
    On Error GoTo ErrorManager
    Set frmEstrModelo = New EstrModeloEdit
    frmEstrModelo.Component mobjModelo.EstrModelos.Add
    frmEstrModelo.Show vbModal
    LoadEstrModelos
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdEdit_Click()

    Dim frmEstrModelo As EstrModeloEdit
    
    If lvwEstrModelos.SelectedItem Is Nothing Then Exit Sub
    
    On Error GoTo ErrorManager
    Set frmEstrModelo = New EstrModeloEdit
    frmEstrModelo.Component _
        mobjModelo.EstrModelos(Val(lvwEstrModelos.SelectedItem.Key))
    frmEstrModelo.Show vbModal
    LoadEstrModelos
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdRemove_Click()

    On Error GoTo ErrorManager
        
    If lvwEstrModelos.SelectedItem Is Nothing Then Exit Sub
    mobjModelo.EstrModelos.Remove Val(lvwEstrModelos.SelectedItem.Key)
    LoadEstrModelos
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadEstrModelos()

    Dim objEstrModelo As EstrModelo
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwEstrModelos.ListItems.Clear
    For lngIndex = 1 To mobjModelo.EstrModelos.Count
        Set itmList = lvwEstrModelos.ListItems.Add _
            (Key:=Format$(lngIndex) & "K")
        Set objEstrModelo = mobjModelo.EstrModelos(lngIndex)

        With itmList
            'If objEstrModelo.IsNew Then
            '    .Text = "(new)"
        
            'Else
            '    .Text = objEstrModelo.EstrModeloID
        
            'End If
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
            
            If objEstrModelo.IsDeleted Then .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
            .SubItems(1) = Trim(objEstrModelo.Material)
            .SubItems(2) = objEstrModelo.Cantidad
            .SubItems(3) = FormatoMoneda(objEstrModelo.Precio, GescomMain.objParametro.Moneda)
            .SubItems(4) = objEstrModelo.Observaciones
        End With

    Next
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub

End Sub
