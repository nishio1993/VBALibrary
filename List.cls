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
    ReDim Data(Capacity - 1)
    Count = 0
End Sub

Private Sub Class_Terminate()

End Sub

Public Sub Init(Capacity As Integer)
    Me.Capacity = Capacity
    ReDim Preserve Data(Me.Capacity)
End Sub

Public Sub Add(value As Variant)
    Count = Count + 1
    If Count = Capacity Then
        Capacity = Capacity * 2
        ReDim Preserve Data(Capacity - 1)
    End If
    Call Allocate(Count - 1, value)
End Sub

Public Sub InsertAt(index As Integer, value As Variant)
    Dim i As Integer
    For i = Count To index Step -1
        If i = index Then
            Call Allocate(i, value)
        Else
            Data(i) = Data(i - 1)
        End If
    Next i
End Sub

Public Sub UpdateAt(index As Integer, value As Variant)
    Call Allocate(index, value)
End Sub

Public Sub Remove(value As Variant)
    Dim i As Integer
    For i = 0 To Count - 1
        If IsObject(value) Then
            If Data(i) Is value Then
                Call RemoveAt(i)
                Exit For
            End If
        Else
            If Data(i) = value Then
                Call RemoveAt(i)
                Exit For
            End If
        End If
    Next i
End Sub

Public Sub RemoveAt(index As Integer)
    Data(index) = Empty
    For i = index To Capacity - 2
        If IsObject(Data(i + 1)) Then
            Set Data(i) = Data(i + 1)
        Else
            Data(i) = Data(i + 1)
        End If
        If IsEmpty(Data(i + 1)) Then
            Exit For
        End If
    Next i
End Sub

Public Function IndexOf(index As Integer) As Variant
    IndexOf = Data(index)
End Function

Public Function Contains(value As Variant) As Boolean
    Contains = False
    For i = 0 To Count - 1
        If IsEmpty(Data(i)) = False Then
            If IsObject(value) Then
                If Data(i) Is value Then
                    Contains = True
                    Exit For
                End If
            Else
                If Data(i) = value Then
                    Contains = True
                    Exit For
                End If
            End If
        End If
    Next i
End Function

Public Sub Sort(isAsc As Boolean)

End Sub

Private Sub Allocate(index As Integer, value As Variant)
    If IsObject(value) Then
        Set Data(index) = value
    Else
        Data(index) = value
    End If
End Sub
