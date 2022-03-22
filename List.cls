VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Capacity As Long
Private Data() As Variant
Public Count As Long

Private Sub Class_Initialize()
    Capacity = 4
    ReDim Data(Capacity - 1)
    Count = 0
End Sub

Private Sub Class_Terminate()

End Sub

Public Sub Init(Capacity As Long)
    Me.Capacity = Capacity
    ReDim Preserve Data(Me.Capacity - 1)
End Sub

Public Sub Add(value As Variant)
    Count = Count + 1
    If Count = Capacity Then
        Capacity = Capacity * 2
        ReDim Preserve Data(Capacity - 1)
    End If
    Call assignAt(Count - 1, value)
End Sub

Public Sub InsertAt(index As Long, value As Variant)
    Dim i As Long
    For i = Count To index Step -1
        If i = index Then
            Call assignAt(i, value)
        Else
            Data(i) = Data(i - 1)
        End If
    Next i
End Sub

Public Sub UpdateAt(index As Long, value As Variant)
    Call assignAt(index, value)
End Sub

Public Sub Remove(value As Variant)
    Dim i As Long
    For i = 0 To Count - 1
        If isEquals(Data(i), value) Then
            Call RemoveAt(i)
            Exit For
        End If
    Next i
End Sub

Public Sub RemoveAt(index As Long)
    Data(index) = Empty
    For i = index To Capacity - 2
        Call assignAt(i, Data(i + 1))
        If IsEmpty(Data(i + 1)) Then
            Exit For
        End If
    Next i
End Sub

Public Function Item(index As Long) As Variant
    Item = Data(index)
End Function

Public Function Contains(value As Variant) As Boolean
    Contains = False
    For i = 0 To Count - 1
        If IsEmpty(Data(i)) = False Then
            If isEquals(Data(i), value) Then
                Contains = True
                Exit For
            End If
        End If
    Next i
End Function

Public Sub Sort(isAsc As Boolean)

End Sub

Public Function ToArray() As Variant()
    Dim arr() As Variant
    ReDim arr(Count - 1)
    For i = 0 To Count - 1
        arr(i) = Data(i)
    Next i
    ToArray = arr
End Function

Public Sub Reverse()
    Dim last As Long
    Dim tmp As Variant
    last = Count / 2
    For i = 0 To last
        If i = last Then
            Exit For
        End If
        If IsObject(Data(i)) Then
            Set tmp = Data(i)
            Set Data(i) = Data(Count - 1 - i)
            Set Data(Count - 1 - i) = tmp
            
        Else
            tmp = Data(i)
            Data(i) = Data(Count - 1 - i)
            Data(Count - 1 - i) = tmp
        End If
    Next i
End Sub

Public Function DeepCopy() As List
    Dim newList as New List
    newList.Init(Capacity)
    For i = 0 To Count - 1
        newList.Add Data(i)
    Next i
    Set DeepCopy = newList
End Function

Private Sub assignAt(index As Long, value As Variant)
    If IsObject(value) Then
        Set Data(index) = value
    Else
        Data(index) = value
    End If
End Sub

Private Function isEquals(value1 As Variant, value2 As Variant) As Boolean
    If IsObject(value1) And IsObject(value2) Then
        If value1 Is value2 Then
            isEquals = True
        Else
            isEquals = False
        End If
    Else
        If value1 = value2 Then
            isEquals = True
        Else
            isEquals = False
        End If
    End If
End Function
