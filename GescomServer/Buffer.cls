VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
  
Private Type BufferProps
    Length As Integer
    EstCount As Long
    MaxCount As Long
    Count As Long
End Type

Private Type BufferData
    Buffer As String * 8
End Type

Private Const BUFFER_START = 9
Private mstrBuffer As String

  Private mudtprops As BufferProps
Private mlngPos As Long

Public Sub Initialize(Length As Integer, EstimatedCount As Long)

    With mudtprops
        .Length = Length
        .EstCount = EstimatedCount
        .MaxCount = EstimatedCount
        .Count = 0
        mstrBuffer = Space$(BUFFER_START + .MaxCount * .Length)
        mlngPos = BUFFER_START
    End With

End Sub

Public Sub Add(Data As String)

    With mudtprops
        If .Count = .MaxCount Then
            mstrBuffer = mstrBuffer & _
                Space$(mudtprops.EstCount / 2 * mudtprops.Length)
            .MaxCount = .MaxCount + mudtprops.EstCount / 2
        End If

        Mid$(mstrBuffer, mlngPos, .Length) = Data
        mlngPos = mlngPos + .Length
        .Count = .Count + 1
    End With

End Sub

Public Function GetState() As String

    Dim udtData As BufferData

    LSet udtData = mudtprops
    Mid$(mstrBuffer, 1, Len(udtData.Buffer)) = udtData.Buffer
    GetState = Left$(mstrBuffer, mlngPos)

End Function

Public Sub SetState(Buffer As String)

    Dim udtData As BufferData
  
    udtData.Buffer = Mid$(Buffer, 1, Len(udtData.Buffer))
    LSet mudtprops = udtData
    mstrBuffer = Buffer

End Sub

Public Property Get Item(Index As Long) As String
Attribute Item.VB_UserMemId = 0

    Item = Mid$(mstrBuffer, BUFFER_START + (Index - 1) * _
        mudtprops.Length, mudtprops.Length)

End Property

Public Function Count() As Long

    Count = mudtprops.Count

End Function

Public Function Length() As Long

    Length = mudtprops.Length

End Function

Public Property Let Item(Index As Long, Buffer As String)

    Mid$(mstrBuffer, BUFFER_START + (Index - 1) * _
        mudtprops.Length, mudtprops.Length) = Buffer

End Property
