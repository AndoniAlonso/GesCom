VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ConsultaItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ConsultaItems"
   ClientHeight    =   3015
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   4875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ConsultaItemEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Elemento de Consulta"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin MSComCtl2.DTPicker DTPValor1 
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         Format          =   78184449
         CurrentDate     =   37026
      End
      Begin VB.ComboBox cboOperador 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   820
         Width           =   2055
      End
      Begin VB.TextBox txtValor1 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   1300
         Width           =   2895
      End
      Begin VB.ComboBox cboCampo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   340
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DTPValor2 
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   1680
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         Format          =   78184449
         CurrentDate     =   37026
      End
      Begin VB.TextBox txtValor2 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   1660
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Campo"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Valor"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Operador"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblValor2 
         Caption         =   "Valor"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "ConsultaItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjConsultaItem As ConsultaItem
Attribute mobjConsultaItem.VB_VarHelpID = -1

Public Sub Component(ConsultaItemObject As ConsultaItem)

    Set mobjConsultaItem = ConsultaItemObject

End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjConsultaItem.ApplyEdit
    mobjConsultaItem.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    mobjConsultaItem.CancelEdit
    Unload Me

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjConsultaItem.ApplyEdit
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub DTPValor1_Change()
    
    mobjConsultaItem.Valor1 = DTPValor1.Value

End Sub

Private Sub DTPValor2_Change()
    
    mobjConsultaItem.Valor2 = DTPValor2.Value

End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjConsultaItem
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Elemento de Consulta [(nuevo)]"
    
        Else
            Caption = "Elemento de Consulta"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        LoadCombo cboCampo, .Campos
        cboCampo.Text = .Campo
    
        LoadCombo cboOperador, .Operadores
        cboOperador.Text = .Operador

        ConfiguraCampos

            
        .BeginEdit
    
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjConsultaItem_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub mobjConsultaItem_OperadorTernario(IsTer As Boolean)

    txtValor2.Visible = mobjConsultaItem.DosValores
    lblValor2.Visible = mobjConsultaItem.DosValores

End Sub

Private Sub cboOperador_Click()

    If mflgLoading Then Exit Sub
    mobjConsultaItem.Operador = cboOperador.Text
    ConfiguraCampos
    
End Sub

Private Sub cboCampo_Click()

    If mflgLoading Then Exit Sub
    mobjConsultaItem.Campo = cboCampo.Text
    
    ConfiguraCampos
End Sub

Private Sub txtValor1_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtValor1

End Sub

Private Sub txtValor1_Change()

    If Not mflgLoading Then _
        TextChange txtValor1, mobjConsultaItem, "Valor1"

End Sub

Private Sub txtValor1_LostFocus()

    txtValor1 = TextLostFocus(txtValor1, mobjConsultaItem, "Valor1")

End Sub

Private Sub txtValor2_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtValor2

End Sub

Private Sub txtValor2_Change()

    If Not mflgLoading Then _
        TextChange txtValor2, mobjConsultaItem, "Valor2"

End Sub

Private Sub txtValor2_LostFocus()

    txtValor2 = TextLostFocus(txtValor2, mobjConsultaItem, "Valor2")

End Sub

Private Sub ConfiguraCampos()

        If mobjConsultaItem.ConsultaCampo.IsDate Then
           DTPValor1.Value = mobjConsultaItem.Valor1
           DTPValor2.Value = mobjConsultaItem.Valor2
           txtValor1.Visible = False
           txtValor2.Visible = False
           DTPValor1.Visible = True
           DTPValor2.Visible = mobjConsultaItem.DosValores
        Else
           txtValor1 = mobjConsultaItem.Valor1
           txtValor2 = mobjConsultaItem.Valor2
           DTPValor1.Visible = False
           DTPValor2.Visible = False
           txtValor1.Visible = True
           txtValor2.Visible = mobjConsultaItem.DosValores
        End If
        lblValor2.Visible = mobjConsultaItem.DosValores

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
       
    IsList = False
    
End Function
