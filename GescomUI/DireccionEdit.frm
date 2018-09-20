VERSION 5.00
Begin VB.Form DireccionEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Direcciones"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DireccionEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   4680
      TabIndex        =   21
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Dirección"
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtCalle 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   340
         Width           =   3855
      End
      Begin VB.TextBox txtPoblacion 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   700
         Width           =   3855
      End
      Begin VB.TextBox txtCodigoPostal 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1060
         Width           =   855
      End
      Begin VB.TextBox txtProvincia 
         Height          =   285
         Left            =   3480
         TabIndex        =   8
         Top             =   1060
         Width           =   1815
      End
      Begin VB.TextBox txtPais 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1420
         Width           =   1215
      End
      Begin VB.TextBox txtTelefono1 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1780
         Width           =   1575
      End
      Begin VB.TextBox txtTelefono2 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   2140
         Width           =   1575
      End
      Begin VB.TextBox txtTelefono3 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   2500
         Width           =   1575
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   3720
         TabIndex        =   14
         Top             =   1780
         Width           =   1575
      End
      Begin VB.TextBox txtEMAIL 
         Height          =   285
         Left            =   1440
         TabIndex        =   18
         Top             =   2860
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Calle"
         Height          =   195
         Left            =   255
         TabIndex        =   1
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Población"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código Postal"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "País"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Teléfonos"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         Height          =   195
         Left            =   3240
         TabIndex        =   13
         Top             =   1800
         Width           =   270
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   2880
         Width           =   420
      End
   End
End
Attribute VB_Name = "DireccionEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjDireccion As Direccion
Attribute mobjDireccion.VB_VarHelpID = -1

Public Sub Component(DireccionObject As Direccion)

    Set mobjDireccion = DireccionObject

End Sub

Private Sub cmdApply_Click()

    mobjDireccion.ChildApplyEdit
    mobjDireccion.ChildBeginEdit

End Sub

Private Sub cmdCancel_Click()

    mobjDireccion.ChildCancelEdit
    Unload Me

End Sub

Private Sub cmdOK_Click()

    mobjDireccion.ChildApplyEdit
    Unload Me

End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjDireccion
        EnableOK .IsValid
        
        If .IsNew Then
            Caption = "Dirección [(nueva)]"

        Else
            Caption = "Dirección [" & .Calle & "]"
      
        End If
        
        txtCalle = .Calle
        txtPoblacion = .Poblacion
        txtCodigoPostal = .CodigoPostal
        txtProvincia = .Provincia
        txtPais = .Pais
        txtTelefono1 = .Telefono1
        txtTelefono2 = .Telefono2
        txtTelefono3 = .Telefono3
        txtFax = .Fax
        txtEMAIL = .EMAIL
        .ChildBeginEdit
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjDireccion_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtCalle_Change()

    If Not mflgLoading Then _
        TextChange txtCalle, mobjDireccion, "Calle"

End Sub

Private Sub txtCalle_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCalle
        
End Sub

Private Sub txtCalle_LostFocus()

    TextLostFocus txtCalle, mobjDireccion, "Calle"

End Sub

Private Sub txtCodigoPostal_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCodigoPostal
        
End Sub

Private Sub txtEMAIL_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtEMAIL
        
End Sub

Private Sub txtFax_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtFax
        
End Sub

Private Sub txtPais_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPais
        
End Sub

Private Sub txtPoblacion_Change()

    If Not mflgLoading Then _
        TextChange txtPoblacion, mobjDireccion, "Poblacion"

End Sub

Private Sub txtPoblacion_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPoblacion
        
End Sub

Private Sub txtPoblacion_LostFocus()

    TextLostFocus txtPoblacion, mobjDireccion, "Poblacion"

End Sub

Private Sub txtCodigoPostal_Change()

    If Not mflgLoading Then _
        TextChange txtCodigoPostal, mobjDireccion, "CodigoPostal"

End Sub

Private Sub txtCodigoPostal_LostFocus()

    TextLostFocus txtCodigoPostal, mobjDireccion, "CodigoPostal"

End Sub

Private Sub txtProvincia_Change()

    If Not mflgLoading Then _
        TextChange txtProvincia, mobjDireccion, "Provincia"

End Sub

Private Sub txtProvincia_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtProvincia
        
End Sub

Private Sub txtProvincia_LostFocus()

    TextLostFocus txtProvincia, mobjDireccion, "Provincia"

End Sub

Private Sub txtPais_Change()

    If Not mflgLoading Then _
        TextChange txtPais, mobjDireccion, "Pais"

End Sub

Private Sub txtPais_LostFocus()

    TextLostFocus txtPais, mobjDireccion, "Pais"

End Sub

Private Sub txtTelefono1_Change()

    If Not mflgLoading Then _
        TextChange txtTelefono1, mobjDireccion, "Telefono1"

End Sub

Private Sub txtTelefono1_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtTelefono1
        
End Sub

Private Sub txtTelefono1_LostFocus()

    TextLostFocus txtTelefono1, mobjDireccion, "Telefono1"

End Sub

Private Sub txtTelefono2_Change()

    If Not mflgLoading Then _
        TextChange txtTelefono2, mobjDireccion, "Telefono2"

End Sub

Private Sub txtTelefono2_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtTelefono2
        
End Sub

Private Sub txtTelefono2_LostFocus()

    TextLostFocus txtTelefono2, mobjDireccion, "Telefono2"

End Sub

Private Sub txtTelefono3_Change()

    If Not mflgLoading Then _
        TextChange txtTelefono3, mobjDireccion, "Telefono3"

End Sub

Private Sub txtTelefono3_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtTelefono3
        
End Sub

Private Sub txtTelefono3_LostFocus()

    TextLostFocus txtTelefono3, mobjDireccion, "Telefono3"

End Sub

Private Sub txtFax_Change()

    If Not mflgLoading Then _
        TextChange txtFax, mobjDireccion, "Fax"

End Sub

Private Sub txtFax_LostFocus()

    TextLostFocus txtFax, mobjDireccion, "Fax"

End Sub

Private Sub txtEMAIL_Change()

    If Not mflgLoading Then _
        TextChange txtEMAIL, mobjDireccion, "EMAIL"

End Sub

Private Sub txtEMAIL_LostFocus()

    TextLostFocus txtEMAIL, mobjDireccion, "EMAIL"

End Sub
