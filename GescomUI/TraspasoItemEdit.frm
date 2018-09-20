VERSION 5.00
Begin VB.Form TraspasoItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Línea de traspasos entre almacén/tiendas"
   ClientHeight    =   3105
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TraspasoItemEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de la línea de traspasos de artículos"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.TextBox txtArticuloColor 
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox txtCodigoArticuloColor 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   1920
         Width           =   6135
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cantidades por tallas"
         Height          =   735
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   9135
         Begin VB.TextBox txtCantidad 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   5
            Top             =   315
            Width           =   495
         End
         Begin VB.TextBox txtTalla 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2280
            TabIndex        =   6
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Talla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   8
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Label Label16 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Artículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "TraspasoItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjTraspasoItem As TraspasoItem
Attribute mobjTraspasoItem.VB_VarHelpID = -1


Public Sub Component(TraspasoItemObject As TraspasoItem)

  Set mobjTraspasoItem = TraspasoItemObject

End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager
    
    mobjTraspasoItem.ApplyEdit
    mobjTraspasoItem.BeginEdit
    Exit Sub
    
ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub
End Sub

Private Sub cmdCancel_Click()

    mobjTraspasoItem.CancelEdit
    Unload Me

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjTraspasoItem.ApplyEdit
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
    Exit Sub
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjTraspasoItem
        EnableOK .IsValid
        
        If .IsNew Then
          Caption = "TraspasoItem [(nuevo)]"
        
        Else
          Caption = "TraspasoItem [" '& .ArticuloColor & "]"
        
        End If
        
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtCantidad.Text = .Cantidad
        Select Case True
        Case .CantidadT36 > 0
            txtTalla = "36"
        Case .CantidadT38 > 0
            txtTalla = "38"
        Case .CantidadT40 > 0
            txtTalla = "40"
        Case .CantidadT42 > 0
            txtTalla = "42"
        Case .CantidadT44 > 0
            txtTalla = "44"
        Case .CantidadT46 > 0
            txtTalla = "46"
        Case .CantidadT48 > 0
            txtTalla = "48"
        Case .CantidadT50 > 0
            txtTalla = "50"
        Case .CantidadT52 > 0
            txtTalla = "52"
        Case .CantidadT54 > 0
            txtTalla = "54"
        Case .CantidadT56 > 0
            txtTalla = "56"
        Case Else
            Err.Raise vbObjectError + 1001, "TraspasoItemEdit LoadTraspasoItems", "No hay cantidad en ninguna talla!"
        End Select
        
        'txtDescripcion = .Descripcion
        'txtArticuloColor = .Descripcion
            
        .BeginEdit
        
        If .ArticuloColorID <> 0 Then _
            txtCodigoArticuloColor = .CodigoArticuloColor
    
    End With
    
    mflgLoading = False
    
End Sub

Private Sub EnableOK(flgValid As Boolean)

  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid

End Sub

Private Sub mobjTraspasoItem_Valid(IsValid As Boolean)

  EnableOK IsValid

End Sub

Private Sub txtDescripcion_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtDescripcion

End Sub
Private Sub txtDescripcion_Change()

  If Not mflgLoading Then _
    TextChange txtDescripcion, mobjTraspasoItem, "Descripcion"

End Sub
Private Sub txtDescripcion_LostFocus()

  txtDescripcion = TextLostFocus(txtDescripcion, mobjTraspasoItem, "Descripcion")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   IsList = False
End Function

