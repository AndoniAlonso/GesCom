VERSION 5.00
Begin VB.Form PrendaEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prendas"
   ClientHeight    =   4575
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PrendaEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos de la Prenda"
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.Frame Frame1 
         Caption         =   "Costes de Fabricación"
         Height          =   2295
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   6015
         Begin VB.TextBox txtPlancha 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   7
            Top             =   340
            Width           =   1455
         End
         Begin VB.TextBox txtTransporte 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   9
            Top             =   700
            Width           =   1455
         End
         Begin VB.TextBox txtPercha 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   11
            Top             =   1060
            Width           =   1455
         End
         Begin VB.TextBox txtCarton 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   15
            Top             =   1420
            Width           =   1455
         End
         Begin VB.TextBox txtEtiqueta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   17
            Top             =   1780
            Width           =   1455
         End
         Begin VB.TextBox txtAdministracion 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4320
            TabIndex        =   13
            Top             =   1060
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Plancha"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Transporte"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Percha"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cartón"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Etiquetado"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   1800
            Width           =   780
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "% Administración"
            Height          =   195
            Left            =   2880
            TabIndex        =   12
            Top             =   1080
            Width           =   1245
         End
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   700
         Width           =   3975
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
         Top             =   720
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
      Left            =   5640
      TabIndex        =   20
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   4080
      Width           =   1095
   End
End
Attribute VB_Name = "PrendaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjPrenda As Prenda
Attribute mobjPrenda.VB_VarHelpID = -1

Public Sub Component(PrendaObject As Prenda)

    Set mobjPrenda = PrendaObject

End Sub

Private Sub cmdApply_Click()
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorManager

    Respuesta = MostrarMensaje(MSG_MODIF_ARTICULO)
    
    mobjPrenda.ApplyEdit
    mobjPrenda.BeginEdit GescomMain.objParametro.Moneda
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()
    Dim Respuesta As VbMsgBoxResult
    
    If mobjPrenda.IsDirty And Not mobjPrenda.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjPrenda.CancelEdit
            Unload Me
        End If
    Else
        mobjPrenda.CancelEdit
        Unload Me
    End If
    
End Sub

Private Sub cmdOK_Click()
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorManager

    Respuesta = MostrarMensaje(MSG_MODIF_ARTICULO)
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjPrenda.ApplyEdit
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
    With mobjPrenda
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Prenda [(nueva)]"

        Else
            Caption = "Prenda [" & .Nombre & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNombre = .Nombre
        txtCodigo = .Codigo
        txtPlancha = .Plancha
        txtTransporte = .transporte
        txtPercha = .percha
        txtCarton = .Carton
        txtEtiqueta = .Etiqueta
        txtAdministracion = .Administracion
        
        .BeginEdit GescomMain.objParametro.Moneda
    
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjPrenda_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtAdministracion_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtAdministracion
        
End Sub

Private Sub txtCarton_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCarton
        
End Sub

Private Sub txtCodigo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCodigo
        
End Sub

Private Sub txtEtiqueta_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtEtiqueta
        
End Sub

Private Sub txtNombre_Change()

    If Not mflgLoading Then _
        TextChange txtNombre, mobjPrenda, "Nombre"

End Sub

Private Sub txtNombre_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNombre
        
End Sub

Private Sub txtNombre_LostFocus()

    txtNombre = TextLostFocus(txtNombre, mobjPrenda, "Nombre")

End Sub

Private Sub txtCodigo_Change()

    If Not mflgLoading Then _
        TextChange txtCodigo, mobjPrenda, "Codigo"

End Sub

Private Sub txtCodigo_LostFocus()

    txtCodigo = TextLostFocus(txtCodigo, mobjPrenda, "Codigo")

End Sub

Private Sub txtPercha_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPercha
        
End Sub

Private Sub txtPlancha_Change()

    If Not mflgLoading Then _
        TextChange txtPlancha, mobjPrenda, "Plancha"

End Sub

Private Sub txtPlancha_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPlancha
        
End Sub

Private Sub txtPlancha_LostFocus()

    txtPlancha = TextLostFocus(txtPlancha, mobjPrenda, "Plancha")

End Sub

Private Sub txtTransporte_Change()

    If Not mflgLoading Then _
        TextChange txtTransporte, mobjPrenda, "Transporte"

End Sub

Private Sub txtTransporte_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtTransporte
        
End Sub

Private Sub txtTransporte_LostFocus()

    txtTransporte = TextLostFocus(txtTransporte, mobjPrenda, "Transporte")

End Sub

Private Sub txtPercha_Change()

    If Not mflgLoading Then _
        TextChange txtPercha, mobjPrenda, "Percha"

End Sub

Private Sub txtPercha_LostFocus()

    txtPercha = TextLostFocus(txtPercha, mobjPrenda, "Percha")

End Sub

Private Sub txtCarton_Change()

    If Not mflgLoading Then _
        TextChange txtCarton, mobjPrenda, "Carton"

End Sub

Private Sub txtCarton_LostFocus()

    txtCarton = TextLostFocus(txtCarton, mobjPrenda, "Carton")

End Sub

Private Sub txtEtiqueta_Change()

    If Not mflgLoading Then _
        TextChange txtEtiqueta, mobjPrenda, "Etiqueta"

End Sub

Private Sub txtEtiqueta_LostFocus()

    txtEtiqueta = TextLostFocus(txtEtiqueta, mobjPrenda, "Etiqueta")

End Sub

Private Sub txtAdministracion_Change()

    If Not mflgLoading Then _
        TextChange txtAdministracion, mobjPrenda, "Administracion"

End Sub

Private Sub txtAdministracion_LostFocus()

    txtAdministracion = TextLostFocus(txtAdministracion, mobjPrenda, "Administracion")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function
