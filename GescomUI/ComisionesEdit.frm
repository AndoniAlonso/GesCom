VERSION 5.00
Begin VB.Form ComisionesEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comisiones"
   ClientHeight    =   3810
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ComisionesEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Representantes"
      Height          =   1335
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
      Begin VB.ComboBox cboRepresentante 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Text            =   "cboRepresentante"
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton optRepresentanteTodos 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optRepresentanteUno 
         Caption         =   "&Un representante"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   2160
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   435
      Left            =   960
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fecha"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txtFechaFinal 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtFechaInicial 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Facturas entre"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "y"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   255
      End
   End
End
Attribute VB_Name = "ComisionesEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjComisiones As Comisiones
Attribute mobjComisiones.VB_VarHelpID = -1

Public Sub Component(ComisionesObject As Comisiones)

    Set mobjComisiones = ComisionesObject

End Sub

Private Sub cboRepresentante_Click()
  
    On Error GoTo ErrorManager

    mobjComisiones.Representante = cboRepresentante.Text
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()
    Dim objPrintComisiones As PrintComisiones

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjComisiones.Load

    Set objPrintComisiones = New PrintComisiones
    objPrintComisiones.Component mobjComisiones
    objPrintComisiones.PrintObject
    Set objPrintComisiones = Nothing

    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()

    DisableX Me

    mflgLoading = True

    With mobjComisiones
        EnableOK .IsValid

        Caption = "Comisiones de representantes"

        ' Aquí se vuelcan los campos del objeto al interfaz
        txtFechaInicial = .FechaInicial
        txtFechaFinal = .FechaFinal

        optRepresentanteTodos.Value = True
        cboRepresentante.Enabled = False
        
        LoadCombo cboRepresentante, .Representantes
        cboRepresentante.Text = .Representante

    End With

    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid

End Sub

Private Sub mobjComisiones_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub optRepresentanteTodos_Click()
    
    cboRepresentante.Enabled = False
    mobjComisiones.Representante = "(Seleccionar uno)"
    
End Sub

Private Sub optRepresentanteUno_Click()

    cboRepresentante.Enabled = True

End Sub

Private Sub txtFechaInicial_Change()

'    If Not mflgLoading Then _
'        TextChange txtFechaInicial, mobjComisiones, "FechaInicial"

End Sub

Private Sub txtFechaInicial_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtFechaInicial

End Sub

Private Sub txtFechaInicial_LostFocus()

    TextChange txtFechaInicial, mobjComisiones, "FechaInicial"
    
    txtFechaInicial = TextLostFocus(txtFechaInicial, mobjComisiones, "FechaInicial")
    If mobjComisiones.FechaFinal < mobjComisiones.FechaInicial Then txtFechaFinal = txtFechaInicial

End Sub

Private Sub txtFechaFinal_Change()

 '   If Not mflgLoading Then _
 '       TextChange txtFechaFinal, mobjComisiones, "FechaFinal"

End Sub

Private Sub txtFechaFinal_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtFechaFinal

End Sub

Private Sub txtFechaFinal_LostFocus()

    TextChange txtFechaFinal, mobjComisiones, "FechaFinal"
  
    txtFechaFinal = TextLostFocus(txtFechaFinal, mobjComisiones, "FechaFinal")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean

    IsList = False

End Function

