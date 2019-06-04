VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim data As Variant

Private Sub Class_Initialize()
    data = Array()
End Sub

Private Sub Class_Terminate()
    Set data = Nothing
End Sub
Public Function push(ByVal item As Variant)
    ReDim Preserve data(UBound(data) + 1)
    data(UBound(data)) = item
End Function
Public Function insert(ByVal item As Variant, Optional ByVal position As Long = 0)
    Dim i As Long
    Dim tmp, tmp2 As Variant
    
    
    If position >= 0 Then
        tmp = Array()
        ReDim tmp(UBound(data) + 1)
        For i = 0 To UBound(data)
            If i < position Then
                tmp(i) = data(i)
            ElseIf i > position Then
                tmp(i + 1) = data(i)
            Else
                tmp(i) = item
                tmp(i + 1) = data(i)
            End If
        Next i
        
        If position = UBound(tmp) Then
            tmp(UBound(tmp)) = item
        End If
        
        tmp2 = data
        data = tmp
        Set tmp2 = Nothing
        Exit Function
    End If
End Function
Public Function delete(ByVal position As Long)
    Dim i As Long
    Dim tmp, tmp2 As Variant
    
    If position >= LBound(data) And position <= UBound(data) Then
        tmp = Array()
        Select Case UBound(data)
            Case 0
                tmp = Array()
            Case 1
                ReDim tmp(0)
            Case Else
                ReDim tmp(UBound(data) - 1)
        End Select
        For i = 0 To UBound(data)
            If i < position Then
                tmp(i) = data(i)
            ElseIf i > position Then
                If i > 0 Then
                    tmp(i - 1) = data(i)
                End If
            End If
        Next i
        
        tmp2 = data
        data = tmp
        Set tmp2 = Nothing
        Exit Function
    End If
End Function
Public Function drop(ByVal Value As Variant)
    Dim i, j As Long
    Dim tmp, tmp2 As Variant
    
    For i = 0 To UBound(data)
        If Value = data(i) Then
            tmp = Array()
            Select Case UBound(data)
                Case 0
                    tmp = Array()
                Case 1
                    ReDim tmp(0)
                Case Else
                    ReDim tmp(UBound(data) - 1)
            End Select
            For j = 0 To UBound(data)
                If j < i Then
                    tmp(j) = data(j)
                ElseIf j > i Then
                    If j > 0 Then
                        tmp(j - 1) = data(j)
                    End If
                End If
            Next j
            
            tmp2 = data
            data = tmp
            Set tmp2 = Nothing
            Exit Function
        End If
    Next i
End Function
Public Function pop() As Variant
    Static popVal As Variant
    
    Dim tmp As Variant
    
    If UBound(data) >= 0 Then
        popVal = back()
        If UBound(data) > 0 Then
            ReDim Preserve data(UBound(data) - 1)
        Else
            tmp = data
            data = Array()
            Set tmp = Nothing
        End If
    Else
        popVal = Empty
    End If
    pop = popVal
End Function
Public Function front() As Variant
    If UBound(data) >= 0 Then
        front = data(LBound(data))
    Else
        front = Empty
    End If
End Function
Public Function back() As Variant
    If UBound(data) >= 0 Then
        back = data(UBound(data))
    Else
        back = Empty
    End If
End Function
Public Function index(ByVal itr As Long) As Variant
    If itr >= LBound(data) And itr <= UBound(data) Then
        index = data(itr)
    Else
        index = Empty
    End If
End Function
Public Function list(Optional ByVal delimiter As String = ".") As String
    Debug.Print "Length is " & size()
    Debug.Print Join(data, delimiter)
    list = Join(data, delimiter)
End Function
Public Function size() As Long
    size = UBound(data) - LBound(data) + 1
End Function

Public Function last() As Long
    last = UBound(data)
End Function

Public Function clear()
    Set data = Nothing
    data = Array()
End Function
