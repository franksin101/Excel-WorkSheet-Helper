VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PeopleRecommand 
   Caption         =   "�ʽ���"
   ClientHeight    =   4356
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   8076
   OleObjectBlob   =   "PeopleRecommand.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "PeopleRecommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ws As WSHelper
Dim ListRefDict As Object
Dim ListRefDict2 As Object
Dim curMonth As Long
Dim targetCol As Long, targetRow As Long
'targetRow Mod 7 : 3�B4 => �`�]�ԡA5�B6 => ��ԡA0�B1 => �]��
Const refSheetName = "���`�έp"
Const monthAddr = "$K$1"
' ��s���C�� �� QuickSort �i���ƱƧ�
Function AutoUpdateListSortWithTimes(Optional ByVal SheetName As String = refSheetName, Optional ByVal Col As Long = 1, Optional ByVal Row As Long = 1, Optional ByRef sortOrder As Variant = "6")
    Dim wsTimes As New WSHelper
    Dim itrRow As Long
    Dim tmpArr As Variant, tmp As Variant
    
    wsTimes.setSheetName = refSheetName
    
    tmpArr = Array()
    
    For itrRow = Row To wsTimes.maxRow()
        tmp = wsTimes.rGetValueByAddr("$A" & itrRow & ":" & "$I" & itrRow)
        tmp(3) = tmp(3) + tmp(8) ' �`�]�� + ����W�`�]��
        tmp(3) = tmp(3) + tmp(9) ' �`�]�� + ���ǭ��C
        tmp(4) = tmp(4) + tmp(9) ' ��� + ���ǭ��C
        tmp(5) = tmp(5) + tmp(9) ' �]�� + ���ǭ��C
        tmp(6) = tmp(6) + tmp(9) ' ���� + ���ǭ��C
        tmp(7) = tmp(7) + tmp(9) ' �`�� + ���ǭ��C
        
        tmp = Join(Application.Transpose( _
                   Application.Transpose( _
                   Application.index(tmp, 1, 0))), "|")

        ReDim Preserve tmpArr(UBound(tmpArr) + 1)
        tmpArr(UBound(tmpArr)) = tmp
    Next itrRow
    
    Call Quicksort(tmpArr, LBound(tmpArr), UBound(tmpArr), sortOrder)
    
    ListRefDict.RemoveAll
    ListBox1.clear
    
    For itrRow = 0 To UBound(tmpArr)
        tmp = Split(tmpArr(itrRow), "|")
        ListBox1.AddItem (itrRow + 1) & "|" & "�f��" & tmp(0) & " | " & tmp(1)
        ListRefDict.add itrRow, tmp(0)
    Next itrRow
    
    If ListBox1.ListCount > 0 Then
        ListBox1.ListIndex = 0
    End If
    
    If IsArray(tmp) Then
        Erase tmp
    End If
    
    If IsArray(tmpArr) Then
        Erase tmpArr
    End If
    
    Set wsTimes = Nothing
End Function
' ����禡
Function cmp(ByRef A As Variant, ByRef B As Variant, Optional ByRef compareOrder As Variant) As Boolean  'A & B �O �@��r���ơA�Ѱ}�C�X�֪�
    Dim tmpA As Variant, tmpB As Variant
    Dim itrX As Integer
    tmpA = Split(A, "|")
    tmpB = Split(B, "|")
    
    If Not IsArray(compareOrder) Then
        If CLng(tmpA(CInt(compareOrder))) < CLng(tmpB(CInt(compareOrder))) Then
            cmp = True
        Else
            cmp = False
        End If
    Else
        For itrX = LBound(compareOrder) To UBound(compareOrder)
            If CLng(tmpA(CInt(compareOrder(itrX)))) < CLng(tmpB(CInt(compareOrder(itrX)))) Then
                cmp = True
                Exit For
            ElseIf CLng(tmpA(CInt(compareOrder(itrX)))) <> CLng(tmpB(CInt(compareOrder(itrX)))) Then
                cmp = False
                Exit For
            End If
        Next itrX
    End If
    
    ' ����}�C
    If IsArray(tmpA) Then
        Erase tmpA
    End If
    
    ' ����}�C
    If IsArray(tmpB) Then
        Erase tmpB
    End If
End Function
' �ֳt�ƧǺt��k
Function Quicksort(ByRef vArray As Variant, arrLbound As Long, arrUbound As Long, Optional ByVal compareType As Variant = "6")
    'Sorts a one-dimensional VBA array from smallest to largest
    'using a very fast quicksort algorithm variant.
    Dim pivotVal As Variant
    Dim vSwap    As Variant
    Dim tmpLow   As Long
    Dim tmpHi    As Long
    
    tmpLow = arrLbound
    tmpHi = arrUbound
    pivotVal = vArray((arrLbound + arrUbound) / 2)
    
    While (tmpLow <= tmpHi) 'divide
        While (cmp(vArray(tmpLow), pivotVal, compareType) And tmpLow < arrUbound)
            tmpLow = tmpLow + 1
        Wend
        
        While (cmp(pivotVal, vArray(tmpHi), compareType) And tmpHi > arrLbound)
            tmpHi = tmpHi - 1
        Wend
    
        If (tmpLow <= tmpHi) Then
            vSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = vSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    
    If (arrLbound < tmpHi) Then Quicksort vArray, arrLbound, tmpHi, compareType 'conquer
    If (tmpLow < arrUbound) Then Quicksort vArray, tmpLow, arrUbound, compareType 'conquer
End Function
'�۰�������
Function AutoUpdateList(ByVal SheetName, Optional ByVal Col As Long = 1, Optional ByVal Row As Long = 1)
    Dim yItr As Integer
    
    ws.setSheetName = SheetName
    
    ListRefDict.RemoveAll
    ListBox1.clear
    
    For yItr = Row To ws.maxRow() Step 1
        ListBox1.AddItem ws.rGetValueByAxis(Col, yItr)
        ListRefDict.add (ListBox1.ListCount - 1), ws.rGetValueByAxis(1, yItr)
    Next
    
    If ListBox1.ListCount >= 1 Then
        ListBox1.ListIndex = 0
    End If
End Function
'�۰������� (�Ƨǭ��n�ʸ��)
Function AutoUpdateList2(Optional ByVal SheetName = "�H�O���n�ʦ���", Optional ByVal Col As Long = 1, Optional ByVal Row As Long = 1)
    Dim yItr As Integer
    Dim tmpArr As Variant, tmp As Variant
    Dim ws As New WSHelper
    
    
    tmpArr = Array()
    
    ws.setSheetName = SheetName ' �H�O���n�ʦ���
    
    ListRefDict2.RemoveAll
    ListBox2.clear
    
    
    For yItr = Row To ws.maxRow() Step 1
        ReDim Preserve tmpArr(UBound(tmpArr) + 1)
        tmpArr(UBound(tmpArr)) = Join(ws.rGetValueByAddr(ws.colN2A(1) & yItr & ":" & ws.colN2A(4) & yItr), "|") & "|" & yItr
    Next
    
    Call Quicksort(tmpArr, LBound(tmpArr), UBound(tmpArr), "1")
    
    For yItr = LBound(tmpArr) To UBound(tmpArr) Step 1
        tmp = Split(tmpArr(yItr), "|")
        ListBox2.AddItem yItr & "|" & tmp(0)
        ListRefDict2.add (ListBox2.ListCount - 1), Array(tmp(3), tmp(4)) ' 4 �O �ӭȭ쥻��Address Of Row
    Next
    
    If ListBox2.ListCount >= 1 Then
        ListBox2.ListIndex = 0
    End If
    
    If IsArray(tmp) Then
        Erase tmp
    End If
    
    If IsArray(tmpArr) Then
        Erase tmpArr
    End If
    
    Set ws = Nothing
End Function
'�۰ʲŦXList�������ﶵ���
Function AutoMatchList(ByVal Value As String) As Variant
    Dim xItr As Integer
    
    For xItr = 0 To ListBox1.ListCount - 1
        If InStr(1, ListBox1.list(xItr), Value) > 0 Then
            ListBox1.ListIndex = xItr
            AutoMatchList = ListBox1.list(xItr)
            Exit Function
        End If
    Next
    
    Debug.Print ListBox1.ListCount
    
    If ListBox1.ListCount > 0 Then
        ListBox1.ListIndex = 0
    End If
    
    AutoMatchList = False
End Function
' ���ǦV�U/ �v���洫
Private Sub makeDown_Click()
    Dim yItr As Integer, tmp As Integer
    Dim ws As New WSHelper
    
    ws.setSheetName = "�H�O���n�ʦ���" ' �H�O���n�ʦ���
    
    tmp = ws.rGetValueByAxis(2, ListRefDict2(ListBox2.ListIndex)(1))
    
    If ListBox2.ListIndex < ListBox2.ListCount - 1 Then
        Call ws.setValueByAxis(ws.rGetValueByAxis(2, ListRefDict2(ListBox2.ListIndex + 1)(1)), 2, ListRefDict2(ListBox2.ListIndex)(1))
        Call ws.setValueByAxis(tmp, 2, ListRefDict2(ListBox2.ListIndex + 1)(1))
        tmp = ListBox2.ListIndex + 1
        Application.Interactive = False
        Application.ScreenUpdating = False
        Call AutoUpdateList2(Row:=2)
        Application.Interactive = True
        Application.ScreenUpdating = True
        ListBox2.ListIndex = tmp
        
        Call AutoUpdateListSortWithTimes(refSheetName, 2, 2, Array(ListRefDict2(0)(0), _
                                                                   ListRefDict2(1)(0), _
                                                                   ListRefDict2(2)(0), _
                                                                   ListRefDict2(3)(0), _
                                                                   ListRefDict2(4)(0)))
    End If
    
    Set ws = Nothing
End Sub
' ���Ǹm��/ �v�����]
Private Sub makeTop_Click()
    Dim yItr As Integer, minPower As Integer
    Dim ws As New WSHelper
    
    ws.setSheetName = "�H�O���n�ʦ���" ' �H�O���n�ʦ���
    
    For yItr = 0 To (ListBox2.ListCount - 1)
        If ws.rGetValueByAxis(1, ListRefDict2(yItr)(1)) <> ListBox2.list(ListBox2.ListIndex) Then
            Call ws.setValueByAxis(yItr + 2, 2, ListRefDict2(yItr)(1))
        End If
    Next yItr
    
    Call ws.setValueByAxis(1, 2, ListRefDict2(ListBox2.ListIndex)(1))
    
    Call AutoUpdateList2(Row:=2)
    Call AutoUpdateListSortWithTimes(refSheetName, 2, 2, Array(ListRefDict2(0)(0), _
                                                               ListRefDict2(1)(0), _
                                                               ListRefDict2(2)(0), _
                                                               ListRefDict2(3)(0), _
                                                               ListRefDict2(4)(0)))
                                                               
    Set ws = Nothing
End Sub
' ���ǦV�W/ �v���洫
Private Sub makeUp_Click()
    Dim yItr As Integer, tmp As Integer
    Dim ws As New WSHelper
    
    ws.setSheetName = "�H�O���n�ʦ���" ' �H�O���n�ʦ���
    
    tmp = ws.rGetValueByAxis(2, ListRefDict2(ListBox2.ListIndex)(1))
    
    If ListBox2.ListIndex > 0 Then
        Call ws.setValueByAxis(ws.rGetValueByAxis(2, ListRefDict2(ListBox2.ListIndex - 1)(1)), 2, ListRefDict2(ListBox2.ListIndex)(1))
        Call ws.setValueByAxis(tmp, 2, ListRefDict2(ListBox2.ListIndex - 1)(1))
        tmp = ListBox2.ListIndex - 1
        Application.Interactive = False
        Application.ScreenUpdating = False
        Call AutoUpdateList2(Row:=2)
        Application.Interactive = True
        Application.ScreenUpdating = True
        ListBox2.ListIndex = tmp
    End If
    
    Call AutoUpdateListSortWithTimes(refSheetName, 2, 2, Array(ListRefDict2(0)(0), _
                                                               ListRefDict2(1)(0), _
                                                               ListRefDict2(2)(0), _
                                                               ListRefDict2(3)(0), _
                                                               ListRefDict2(4)(0)))
    Set ws = Nothing
End Sub
' �۰ʲŦX�j�M��ﶵ��
Private Sub TextBox1_AfterUpdate()
    If TypeName(AutoMatchList(TextBox1.Value)) = "Boolean" Then
        If Not AutoMatchList(TextBox1.Value) Then
            TextBox1.Value = "�S������ŦX"
        End If
    End If
End Sub
' ��J�s�ȫ�A�۰ʧ�s�έp��ƪ�A�������R����A�s�� + 1�A�ťժ���J�έp + 1
Function inputAndUpdateWithNewData(ByRef Target As Range, ByVal Val As Variant, Optional ByVal updateSheetName As String = "���`�έp")
    Dim oldPNRow As Long, newPNRow As Long
    Dim itrY As Integer, itrX As Integer
    Dim ws As New WSHelper
    Dim workTypeArr As Variant, tmpVar As Variant
    
    workTypeArr = Array("�`�]��", "���", "�]��", "����")
    tmpVar = Target.Resize(1, 1).Value
    Target.Resize(1, 1).Value = Val
        
    ws.setSheetName = updateSheetName
    
    If Len(tmpVar) > 0 Then
        For itrY = 2 To ws.maxRow
            If CInt(ws.rGetValueByAxis(1, itrY)) = tmpVar Then
                oldPNRow = itrY
                Exit For
            End If
        Next itrY
    Else
        oldPNRow = -1
    End If
    
    If Len(Val) > 0 Then
        For itrY = 2 To ws.maxRow
            If CInt(ws.rGetValueByAxis(1, itrY)) = Val Then
                newPNRow = itrY
                Exit For
            End If
        Next itrY
    Else
        newPNRow = -1
    End If
    
    For itrX = 3 To 6
        If oldPNRow <> -1 Then
            Call ws.setValueByAxis(SumOfYear(tmpVar, workTypeArr(itrX - 3), curMonth), itrX, oldPNRow)
        End If
        
        If newPNRow <> -1 Then
            Call ws.setValueByAxis(SumOfYear(Val, workTypeArr(itrX - 3), curMonth), itrX, newPNRow)
        End If
    Next itrX
    
    Erase workTypeArr
    Set ws = Nothing
End Function
' ��ƽT�{��J
Private Sub YES_Click() ' �����R��
    'Selection.Value = ListRefDict(ListBox1.ListIndex)
    'Selection.Offset(0, 1).Value = ListBox1.list(ListBox1.ListIndex)
    'SumOfMonthAllTable (month(Now()))
    Call inputAndUpdateWithNewData(Selection.Resize(1, 1), ListRefDict(ListBox1.ListIndex))
    Unload PeopleRecommand
End Sub
' ��ƽT�{��J
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean) ' �����R��
    'Selection.Value = CStr(ListRefDict(ListBox1.ListIndex))
    'Selection.Offset(0, 1).Value = ListBox1.list(ListBox1.ListIndex)
    'SumOfMonthAllTable (month(Now()))
    Call inputAndUpdateWithNewData(Selection.Resize(1, 1), ListRefDict(ListBox1.ListIndex))
    Unload PeopleRecommand
End Sub
' �M����J
Private Sub Clean_Click()
    'Selection.ClearContents
    Call inputAndUpdateWithNewData(Selection.Resize(1, 1), Empty)
    Unload PeopleRecommand
End Sub
' �����A�h�X���
Private Sub NO_Click()
    Unload PeopleRecommand
End Sub
' �j�M������ܤơA�۰ʲŦX������
Private Sub TextBox1_Change()
    Call AutoMatchList(TextBox1.Value)
End Sub
' ����l��
Private Sub UserForm_Initialize()
    Dim biasRow As Variant

    Set ws = New WSHelper
    ws.setSheetName = refSheetName
    
    ' �����
    biasRow = Selection.Resize(1, 1).Address(ReferenceStyle:=xlR1C1)
    biasRow = Replace(biasRow, "R", "")
    biasRow = Replace(biasRow, "C", vbTab)
    biasRow = Split(biasRow, vbTab)
    targetCol = CLng(biasRow(1))
    targetRow = CLng(biasRow(0))
    
    If IsArray(biasRow) Then
        Erase biasRow
    End If
    
    biasRow = (targetRow - 2) Mod 7
    biasRow = 0 - biasRow
    
    curMonth = Month(Selection.Resize(1, 1).Offset(biasRow, 0))
    If Month(Selection.Resize(1, 1).Offset(biasRow, 0)) <> CLng(ws.rGetValueByAddr(monthAddr)) Then
        Call ws.setValueByAddr(Month(Selection.Resize(1, 1).Offset(biasRow, 0)), monthAddr)
    End If
    
    Call addDictionary(ListRefDict)
    Call addDictionary(ListRefDict2)
    'Call AutoUpdateList(refSheetName, 2, 2)
    Call AutoUpdateList2(Row:=2)
    Call AutoUpdateListSortWithTimes(refSheetName, 2, 2, Array(ListRefDict2(0)(0), _
                                                               ListRefDict2(1)(0), _
                                                               ListRefDict2(2)(0), _
                                                               ListRefDict2(3)(0), _
                                                               ListRefDict2(4)(0)))
End Sub
' �������
Private Sub UserForm_Terminate()
    Call delDictionary(ListRefDict)
    Call delDictionary(ListRefDict2)
    Set ws = Nothing
End Sub

