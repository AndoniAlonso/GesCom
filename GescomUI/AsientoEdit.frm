VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AsientoEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asientos contables"
   ClientHeight    =   5520
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AsientoEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Líneas del Asiento"
      Height          =   3615
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   6735
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtTotalHaber 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtTotalDebe 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Top             =   2640
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwApuntes 
         Height          =   2175
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   3120
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total haber"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2880
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total debe"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   2655
         Width           =   765
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Asiento"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtEjercicio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtConcepto 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   700
         Width           =   5415
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   340
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ejercicio"
         Height          =   195
         Left            =   2280
         TabIndex        =   3
         Top             =   375
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   750
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   5760
      TabIndex        =   20
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   18
      Top             =   5040
      Width           =   1095
   End
End
Attribute VB_Name = "AsientoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjAsiento As Asiento
Attribute mobjAsiento.VB_VarHelpID = -1

Public Sub Component(AsientoObject As Asiento)

    Set mobjAsiento = AsientoObject

End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjAsiento.ApplyEdit
    mobjAsiento.BeginEdit GescomMain.objParametro.Moneda
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()
    Dim Respuesta As VbMsgBoxResult
    
    If mobjAsiento.IsDirty And Not mobjAsiento.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjAsiento.CancelEdit
            Unload Me
        End If
    Else
        mobjAsiento.CancelEdit
        Unload Me
    End If

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjAsiento.ApplyEdit
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjAsiento
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Asiento [(nuevo)]"
    
        Else
            Caption = "Asiento [" & .Concepto & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtConcepto = .Concepto
        txtNumero = .Numero
        txtEjercicio = .Ejercicio
    
        .BeginEdit GescomMain.objParametro.Moneda
    
        If .IsNew Then
            .TemporadaID = GescomMain.objParametro.TemporadaActualID
            .EmpresaID = GescomMain.objParametro.EmpresaActualID
        End If
    
    End With
    
    lvwApuntes.SmallIcons = GescomMain.mglIconosPequeños
    
    lvwApuntes.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwApuntes.ColumnHeaders.Add , , "Cuenta", ColumnSize(10)
    lvwApuntes.ColumnHeaders.Add , , "Concepto", ColumnSize(15)
    lvwApuntes.ColumnHeaders.Add , , "Importe", ColumnSize(10), vbRightJustify
    lvwApuntes.ColumnHeaders.Add , , "D/H", ColumnSize(5)
    LoadApuntes
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub lvwApuntes_DblClick()
    
    Call cmdEdit_Click
    
End Sub

Private Sub mobjAsiento_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtEjercicio_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtEjercicio
        
End Sub

Private Sub txtNumero_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNumero
        
End Sub

Private Sub txtConcepto_Change()

    If Not mflgLoading Then _
        TextChange txtConcepto, mobjAsiento, "Concepto"

End Sub

Private Sub txtConcepto_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtConcepto
        
End Sub

Private Sub txtConcepto_LostFocus()

    txtConcepto = TextLostFocus(txtConcepto, mobjAsiento, "Concepto")

End Sub

Private Sub txtNumero_Change()

    If Not mflgLoading Then _
        TextChange txtNumero, mobjAsiento, "Numero"

End Sub

Private Sub txtNumero_LostFocus()

    txtNumero = TextLostFocus(txtNumero, mobjAsiento, "Numero")

End Sub

Private Sub txtEjercicio_Change()

    If Not mflgLoading Then _
        TextChange txtEjercicio, mobjAsiento, "Ejercicio"

End Sub

Private Sub txtEjercicio_LostFocus()

    txtEjercicio = TextLostFocus(txtEjercicio, mobjAsiento, "Ejercicio")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function

' a partir de aqui -----> child

Private Sub cmdAdd_Click()
    Dim frmApunte As ApunteEdit
  
    On Error GoTo ErrorManager
    Set frmApunte = New ApunteEdit
    frmApunte.Component mobjAsiento.Apuntes.Add
    frmApunte.Show vbModal
    LoadApuntes
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdEdit_Click()
    Dim frmApunte As ApunteEdit
    
    If lvwApuntes.SelectedItem Is Nothing Then Exit Sub
    
    On Error GoTo ErrorManager
    Set frmApunte = New ApunteEdit
    frmApunte.Component _
        mobjAsiento.Apuntes(Val(lvwApuntes.SelectedItem.Key))
    frmApunte.Show vbModal
    LoadApuntes
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdRemove_Click()

    On Error GoTo ErrorManager
        
    If lvwApuntes.SelectedItem Is Nothing Then Exit Sub
    mobjAsiento.Apuntes.Remove Val(lvwApuntes.SelectedItem.Key)
    LoadApuntes
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadApuntes()

    Dim objApunte As Apunte
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwApuntes.ListItems.Clear
    For lngIndex = 1 To mobjAsiento.Apuntes.Count
        Set itmList = lvwApuntes.ListItems.Add _
            (Key:=Format$(lngIndex) & "K")
        Set objApunte = mobjAsiento.Apuntes(lngIndex)

        With itmList
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
            
            If objApunte.IsDeleted Then .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
            .SubItems(1) = Trim(objApunte.Cuenta)
            .SubItems(2) = objApunte.Descripcion
            .SubItems(3) = FormatoMoneda(objApunte.Importe, GescomMain.objParametro.Moneda)
            .SubItems(4) = objApunte.TipoImporte
        End With

    Next
    
    txtTotalDebe = FormatoMoneda(mobjAsiento.Apuntes.TotalDebe, GescomMain.objParametro.Moneda)
    txtTotalHaber = FormatoMoneda(mobjAsiento.Apuntes.TotalHaber, GescomMain.objParametro.Moneda)
    txtSaldo = FormatoMoneda(mobjAsiento.Apuntes.Saldo, GescomMain.objParametro.Moneda)
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub
