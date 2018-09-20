VERSION 5.00
Begin VB.Form SerieEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Series"
   ClientHeight    =   3030
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SerieEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Datos de la Serie"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   700
         Width           =   4575
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
      Height          =   435
      Left            =   5040
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   3840
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   435
      Left            =   2640
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Propiedades"
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   5895
      Begin VB.ComboBox cboMaterial 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Text            =   "cboMaterial"
         Top             =   340
         Width           =   4575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Material"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   570
      End
   End
End
Attribute VB_Name = "SerieEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintMaterialSelStart As Integer

Private WithEvents mobjSerie As Serie
Attribute mobjSerie.VB_VarHelpID = -1

Public Sub Component(SerieObject As Serie)

    Set mobjSerie = SerieObject

End Sub

Private Sub cmdApply_Click()
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorManager

    Respuesta = MostrarMensaje(MSG_MODIF_ARTICULO)
    
    mobjSerie.ApplyEdit
    mobjSerie.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()
    Dim Respuesta As VbMsgBoxResult
    
    If mobjSerie.IsDirty And Not mobjSerie.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjSerie.CancelEdit
            Unload Me
        End If
    Else
        mobjSerie.CancelEdit
        Unload Me
    End If
    
End Sub

Private Sub cmdOK_Click()
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorManager

    Respuesta = MostrarMensaje(MSG_MODIF_ARTICULO)
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjSerie.ApplyEdit
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
    With mobjSerie
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Serie [(nueva)]"

        Else
            Caption = "Serie [" & .Nombre & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNombre = .Nombre
        txtCodigo = .Codigo
        
        LoadCombo cboMaterial, .Materiales
        cboMaterial.Text = .Material
    
        .BeginEdit
    
        If .IsNew Then .TemporadaID = GescomMain.objParametro.TemporadaActualID
    
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjSerie_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtCodigo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCodigo
        
End Sub

Private Sub txtNombre_Change()

    If Not mflgLoading Then _
        TextChange txtNombre, mobjSerie, "Nombre"

End Sub

Private Sub txtNombre_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNombre
        
End Sub

Private Sub txtNombre_LostFocus()

    txtNombre = TextLostFocus(txtNombre, mobjSerie, "Nombre")

End Sub

Private Sub txtCodigo_Change()

    If Not mflgLoading Then _
        TextChange txtCodigo, mobjSerie, "Codigo"

End Sub

Private Sub txtCodigo_LostFocus()

    txtCodigo = TextLostFocus(txtCodigo, mobjSerie, "Codigo")

End Sub

Private Sub cboMaterial_Click()

    If mflgLoading Then Exit Sub
    mobjSerie.Material = cboMaterial.Text

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function

Private Sub cboMaterial_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintMaterialSelStart = cboMaterial.SelStart
End Sub

Private Sub cboMaterial_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintMaterialSelStart, cboMaterial
    
End Sub

