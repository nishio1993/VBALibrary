VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Capacity As Integer
Private Data() As Variant
Public Count As Integer

Private Sub Class_Initialize()
    Capacity = 4
    ReDim Data(Capacity)
    Count = 0
End Sub

Private Sub Class_Terminate()

End Sub

Public Sub Init(Capacity As Integer)
    Me.Capacity = Capacity
    ReDim Preserve Data(Me.Capacity)
End Sub

Public Sub Add(element As Variant)
    Count = Count + 1
    If Count > Capacity Then
        Capacity = Capacity * 2
        ReDim Preserve Data(Capacity)
    End If
    
    Data(Count - 1) = element
End Sub

Public Sub RemoveAt(index As Integer)
    Data(i) = Empty
    For i = index To Capacity - 2
        If IsEmpty(Data(i + 1)) Then
            Exit For
        Else
            Data(i) = Data(i + 1)
        End If
    Next i
End Sub

Public Function IndexOf(index As Integer) As Variant
    IndexOf = Data(index)
End Function

Public Function Contains(element As Variant) As Boolean

End Function

Public Sub Sort(isAsc As Boolean)

End Sub

