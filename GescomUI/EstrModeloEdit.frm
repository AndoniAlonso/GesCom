VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EstrModeloEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Línea del Modelo"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "EstrModeloEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Datos del EstrModelo"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1060
         Width           =   960
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   1420
         Width           =   1695
      End
      Begin VB.TextBox txtPrecioCoste 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   700
         Width           =   1455
      End
      Begin VB.ComboBox cboMaterial 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Text            =   "cboMaterial"
         Top             =   320
         Width           =   3135
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   1780
         Width           =   4095
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2521
         TabIndex        =   7
         Top             =   1060
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtCantidad"
         BuddyDispid     =   196610
         OrigLeft        =   2775
         OrigTop         =   1060
         OrigRight       =   3015
         OrigBottom      =   1345
         Max             =   99999
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Precio"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precio de Coste"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Material"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "EstrModeloEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintMaterialSelStart As Integer

Private WithEvents mobjEstrModelo As EstrModelo
Attribute mobjEstrModelo.VB_VarHelpID = -1

Public Sub Component(EstrModeloObject As EstrModelo)

    Set mobjEstrModelo = EstrModeloObject

End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjEstrModelo.ApplyEdit
    mobjEstrModelo.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    mobjEstrModelo.CancelEdit
    Unload Me

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjEstrModelo.ApplyEdit
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
    With mobjEstrModelo
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Línea del Modelo [(nueva)]"

        Else
            Caption = "Línea del Modelo [" & .Material & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtCantidad = .Cantidad
        txtPrecioCoste = .PrecioCoste
        txtPrecio = .Precio
        txtObservaciones = .Observaciones
        
        LoadCombo cboMaterial, .Materiales
        cboMaterial.Text = .Material

        .BeginEdit
    
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjEstrModelo_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub cboMaterial_Click()

    On Error GoTo ErrorManager

    If mflgLoading Then Exit Sub
    mobjEstrModelo.Material = cboMaterial.Text
    
    txtPrecioCoste = mobjEstrModelo.PrecioCoste
    txtPrecio = mobjEstrModelo.Precio

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub txtCantidad_Change()

    If Not mflgLoading Then
        TextChange txtCantidad, mobjEstrModelo, "Cantidad"
        txtPrecio = mobjEstrModelo.Precio
    End If

End Sub

Private Sub txtCantidad_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidad
        
End Sub

Private Sub txtCantidad_LostFocus()

    txtCantidad = TextLostFocus(txtCantidad, mobjEstrModelo, "Cantidad")

End Sub

Private Sub txtObservaciones_Change()

    If Not mflgLoading Then _
        TextChange txtObservaciones, mobjEstrModelo, "Observaciones"

End Sub

Private Sub txtObservaciones_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtObservaciones
        
End Sub

Private Sub txtObservaciones_LostFocus()

    txtObservaciones = TextLostFocus(txtObservaciones, mobjEstrModelo, "Observaciones")

End Sub

Private Sub txtPrecioCoste_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPrecioCoste

End Sub

Private Sub txtPrecioCoste_Change()

    If Not mflgLoading Then _
        TextChange txtPrecioCoste, mobjEstrModelo, "PrecioCoste"

End Sub

Private Sub txtPrecioCoste_LostFocus()

    txtPrecioCoste = TextLostFocus(txtPrecioCoste, mobjEstrModelo, "PrecioCoste")

End Sub

Private Sub txtPrecio_GotFocus()
    
    If Not mflgLoading Then _
        SelTextBox txtPrecio

End Sub

Private Sub txtPrecio_Change()

    If Not mflgLoading Then _
        TextChange txtPrecio, mobjEstrModelo, "Precio"

End Sub

Private Sub txtPrecio_LostFocus()

    txtPrecio = TextLostFocus(txtPrecio, mobjEstrModelo, "Precio")

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

