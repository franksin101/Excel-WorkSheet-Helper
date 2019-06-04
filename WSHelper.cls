VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WSHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'還有 AutoFilter 待做

#Const DebugVersion = False


Const addrDelimiter = "$"
Const rangeArrayDelimiter = ":"
Const axisDelimiter = ","
Const addrPattern As String = "^\$?[A-Z]+\$?[0-9]+$|^\$?[A-Z]+\$?[0-9]+:\$?[A-Z]+\$?[0-9]+$|^\$?[A-Z]+:\$?[A-Z]+$|^\$?[0-9]+:\$?[0-9]+$" ' OK 全部的Pattern
Const addrRangePattern As String = "^\$?[A-Z]+\$?[0-9]+:\$?[A-Z]+\$?[0-9]+$|^\$?[A-Z]+:\$?[A-Z]+$|^\$?[0-9]+:\$?[0-9]+$" ' OK
Const AbsAddrPattern As String = "^\$?[A-Z]+\$?[0-9]+$" ' OK
Const colPattern = "^\$?[A-Z]+$" ' OK

Private wWorkbook As Workbook
Private wWorkSheetName As String
Private regEx As RegExp
'Private DataMapDict As Object
Private Sub Class_Initialize()
    If Not Application.ActiveSheet Is Nothing And Application.ActiveSheet.Name <> "" Then
        wWorkSheetName = Application.ActiveSheet.Name
    Else
        wWorkSheetName = ""
    End If
    
    Set wWorkbook = Application.Workbooks(Application.ActiveWorkbook.Name)
    
    Set regEx = New RegExp
End Sub
Private Sub Class_Terminate()
    Set regEx = Nothing
End Sub
Public Property Let setWorkbook(ByRef pWorkbook As Workbook)
    Set wWorkbook = pWorkbook
End Property
Public Property Get getWorkbook() As Workbook
    Set getWorkbook = wWorkbook
End Property
Public Property Let setSheetName(ByVal SheetName As String)
    wWorkSheetName = SheetName
End Property
Public Property Get getSheetName()
    getSheetName = wWorkSheetName
End Property
Public Function Active() ' 使自己 Activate
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    wWorkbook.Worksheets(wWorkSheetName).Activate
End Function
Public Function self() As Worksheet ' 傳回自己
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    Set self = wWorkbook.Worksheets(wWorkSheetName)
End Function
Public Function myRange(ByVal addr As String) As Range
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    Set myRange = wWorkbook.Worksheets(wWorkSheetName).Range(addr)
End Function
'取得資料表變數
Public Function getValueByAxis(ByRef Value As Variant, ByVal X As Long, ByVal Y As Long) As Variant
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    Value = wWorkbook.Worksheets(wWorkSheetName).Cells(Y, X).Value
End Function
'取得資料表變數
Public Function getValueByAddr(ByRef Value As Variant, ByVal addr As String) As Variant
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = addrPattern
    End With
    
    If regEx.Test(addr) Then
        If IsArray(wWorkbook.Worksheets(wWorkSheetName).Range(addr).Value) Then
            Value = Application.Transpose( _
                    Application.Transpose( _
                    Application.index(wWorkbook.Worksheets(wWorkSheetName).Range(addr).Value, 1, 0)))
        Else
            Value = wWorkbook.Worksheets(wWorkSheetName).Range(addr).Value
        End If
    Else
        Value = Nothing
#If DebugVersion Then
        MsgBox wWorkbook.Caption & " -> " & wWorkSheetName & " 你填的Excel 位子是無效的，請修改"
#End If
    End If
End Function
'取得資料表變數
Public Function rGetValueByAxis(ByVal X As Long, ByVal Y As Long) As Variant
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    rGetValueByAxis = wWorkbook.Worksheets(wWorkSheetName).Cells(Y, X).Value
End Function
'取得資料表變數
Public Function rGetValueByAddr(ByVal addr As String) As Variant
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = addrPattern
    End With
    
    If regEx.Test(addr) Then
        If IsArray(wWorkbook.Worksheets(wWorkSheetName).Range(addr).Value) Then
            rGetValueByAddr = Application.Transpose( _
                              Application.Transpose( _
                              Application.index(wWorkbook.Worksheets(wWorkSheetName).Range(addr).Value, 1, 0)))
        Else
            rGetValueByAddr = wWorkbook.Worksheets(wWorkSheetName).Range(addr).Value
        End If
    Else
#If DebugVersion Then
        MsgBox "你填的Excel 位子是無效的，請修改"
#End If
    rGetValueByAddr = Nothing
    End If
End Function
'在資料表設值
Public Function setValueByAxis(ByVal Value As Variant, ByVal X As Long, ByVal Y As Long)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    wWorkbook.Worksheets(wWorkSheetName).Cells(Y, X).Value = Value
End Function
'在資料表設值
Public Function setValueByAddr(ByVal Value As Variant, ByVal addr As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = addrPattern
    End With
    
    If regEx.Test(addr) Then
        wWorkbook.Worksheets(wWorkSheetName).Range(addr).Value = Value
    Else
#If DebugVersion Then
        MsgBox wWorkSheetName & " 你填的Excel 位子是無效的，請修改"
#End If
    End If
End Function
'在資料表設值
Public Function setValueByOffset(ByVal Value As Variant, ByVal RangeAddr As String, ByVal xOffset As Long, ByVal yOffset As Long)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = addrPattern
    End With
    
    If regEx.Test(RangeAddr) Then
        wWorkbook.Worksheets(wWorkSheetName).Range(RangeAddr).Offset(yOffset, xOffset).Value = Value
        On Error Resume Next
    Else
#If DebugVersion Then
        MsgBox wWorkSheetName & " 你填的Excel 位子是無效的，請修改"
#End If
    End If
End Function
'該表最大列
Public Function maxRow(Optional ByVal Col As Long = 1) As Long
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    maxRow = wWorkbook.Worksheets(wWorkSheetName).Cells(Worksheets(wWorkSheetName).Rows.Count, Col).End(xlUp).Row
End Function
'該表最大欄
Public Function maxCol(Optional ByVal Row As Long = 1) As Long
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    maxCol = wWorkbook.Worksheets(wWorkSheetName).Cells(Row, Worksheets(wWorkSheetName).Columns.Count).End(xlToLeft).Column
End Function
' 搜尋資料 回傳 Address/慢 必要時用建檔字典會更快
Public Function myFindR(ByVal What As String, ByVal addr As String, Optional afterAddr As String = "A1") As Range
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = addrRangePattern
    End With
    
    If regEx.Test(addr) Then
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .pattern = AbsAddrPattern
        End With
        
        If Len(afterAddr) = 0 Then
            afterAddr = Join(Array(colN2A(colA2N(afterAddr)), 1), "")
        End If
        
        If regEx.Test(afterAddr) Then
            On Error Resume Next
            Set myFindR = wWorkbook.Worksheets(wWorkSheetName).Range(addr).Find(What, _
                                                                               After:=Range(afterAddr), _
                                                                               LookIn:=xlValues, _
                                                                               LookAt:=xlWhole, _
                                                                               SearchOrder:=xlByRows, _
                                                                               MatchCase:=True)
            Debug.Print Err.Description
            Err.clear
            On Error GoTo 0
        End If
    Else
        Set myFindR = Nothing
#If DebugVersion Then
        MsgBox "myFind Address Is not corrected, >> " & wWorkSheetName & " 你填的Excel 位子是無效的，請修改"
#End If
    End If
End Function
' 搜尋資料 回傳 Address/慢 必要時用建檔字典會更快
Public Function myFindC(ByVal What As String, ByVal addr As String, Optional afterAddr As String = "A1") As Range
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = addrRangePattern
    End With
    
    If regEx.Test(addr) Then
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .pattern = AbsAddrPattern
        End With
        
        If Len(afterAddr) = 0 Then
            afterAddr = Join(Array(colN2A(colA2N(afterAddr)), 1), "")
        End If
        
        If regEx.Test(afterAddr) Then
            On Error Resume Next
            Set myFindC = wWorkbook.Worksheets(wWorkSheetName).Range(addr).Find(What, _
                                                                               After:=Range(afterAddr), _
                                                                               LookIn:=xlValues, _
                                                                               LookAt:=xlWhole, _
                                                                               SearchOrder:=xlByColumns, _
                                                                               MatchCase:=True)
            Debug.Print Err.Description
            Err.clear
            On Error GoTo 0
        End If
    Else
        Set myFindC = Nothing
#If DebugVersion Then
        MsgBox "myFind Address Is not corrected, >> " & wWorkSheetName & " 你填的Excel 位子是無效的，請修改"
#End If
    End If
End Function
'用 Key 查詢所求欄位資料
' ValCol 是所求欄位 位置(數字)
Public Function getDataWithKey(ByVal KeyCol As Long, ByVal key As String, ByVal ValCol As Long) As Variant
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    Dim findAddress As Variant
    Set findAddress = myFindR(key, Join(Array(colN2A(KeyCol), colN2A(KeyCol)), rangeArrayDelimiter), afterAddr:=colN2A(KeyCol) & "1")
    If Not findAddress Is Nothing Then
        getDataWithKey = rGetValueByAxis(ValCol, findAddress.Row)
    Else
        getDataWithKey = Empty ' 如果查找不到，傳回空
    End If
End Function
' 用 Key 查詢 並填入相應值
Public Function setDataWithKey(ByVal KeyCol As Long, ByVal key As String, ByVal ValCol As Long, ByVal Val As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    Dim findAddress As Variant
    Set findAddress = myFindR(key, Join(Array(colN2A(KeyCol), colN2A(KeyCol)), rangeArrayDelimiter), afterAddr:=colN2A(KeyCol) & "1")
    If Not findAddress Is Nothing Then
        Call setValueByAxis(Val, ValCol, findAddress.Row)
    Else
        Debug.Print wWorkSheetName & "  getDataWithKey : No Such Key -> " & key & " to Fill."
    End If
End Function
' 建立對應型字典 Column 1 TO 1
Public Function addMapDict(ByRef dict As Object, ByVal KeyCol As Long, ByVal ValCol As Long, Optional ByVal startRow As Long = 1)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    Dim yItr As Long
    Dim key As String
    Dim kValue As String
    
    yItr = startRow
    
    If Not dict Is Nothing Then
        Call delDictionary(dict)
        Call addDictionary(dict)
    Else
        Call addDictionary(dict)
    End If
    
    
    Do Until Len(rGetValueByAxis(KeyCol, yItr)) = 0
        key = rGetValueByAxis(KeyCol, yItr)
        kValue = rGetValueByAxis(ValCol, yItr)
        dict.add key, kValue
        yItr = yItr + 1
    Loop
End Function
' 多重搜尋
'自動清理 選定範圍
Public Function clear(ByVal StartAddress As String, ByVal EndAddress As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = AbsAddrPattern
    End With
    
    If regEx.Test(StartAddress) And regEx.Test(EndAddress) Then
        wWorkbook.Worksheets(wWorkSheetName).Range(Join(Array(StartAddress, EndAddress), rangeArrayDelimiter)).clear
    Else
#If DebugVersion Then
        MsgBox wWorkSheetName & " 你填的Excel 位子是無效的，請修改"
#End If
    End If
End Function
'只清理內容
Public Function ClearContent(ByVal StartAddress As String, ByVal EndAddress As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = AbsAddrPattern
    End With
    
    If regEx.Test(StartAddress) And regEx.Test(EndAddress) Then
        wWorkbook.Worksheets(wWorkSheetName).Range(Join(Array(StartAddress, EndAddress), rangeArrayDelimiter)).ClearContents
    Else
#If DebugVersion Then
        MsgBox wWorkSheetName & " 你填的Excel 位子是無效的，請修改"
#End If
    End If
End Function
'自動填入 (下拉式)
Public Function myFill(ByVal rangeOfAddr As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    Dim OldRange As String
    Dim NewRange As String
    Dim tmpStr1() As String
    Dim tmpStr2() As String
    
    Dim x1 As Long
    Dim y1 As Long
    Dim x2 As Long
    Dim y2 As Long
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = addrRangePattern
    End With
    
    OldRange = rangeOfAddr
    
    If regEx.Test(OldRange) Then
        tmpStr1 = Split(OldRange, rangeArrayDelimiter)
        x1 = addrA2N(tmpStr1(0))(0)
        y1 = addrA2N(tmpStr1(0))(1)
        x2 = addrA2N(tmpStr1(1))(0)
        y2 = addrA2N(tmpStr1(1))(1)
        
        If y1 = y2 Then
            If y1 > 1 Then
                y1 = y2 - 1
            Else
                y2 = y1 + 1
            End If
        End If
        
        OldRange = addrN2A(x1, y1) & rangeArrayDelimiter & addrN2A(x2, y1)
        NewRange = addrN2A(x1, y1) & rangeArrayDelimiter & addrN2A(x2, y2)
        
        Application.ScreenUpdating = False
        wWorkbook.Worksheets(wWorkSheetName).Range(OldRange).AutoFill Destination:=Range(NewRange), Type:=xlFillDefault
        Application.ScreenUpdating = True
    Else
#If DebugVersion Then
        MsgBox wWorkSheetName & ">> myFill : 輸入錯誤Address"
#End If
    End If
End Function
'自動排列 (只可選起始列)
Public Function mySort(ByVal addrOfStartRow As String, ByVal hideColCount As Long, ParamArray var() As Variant)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    Dim startRow As Long
    Dim MaxDataRP As Long
    Dim MaxDataCP As Long
    
    Dim e As Variant
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = AbsAddrPattern
    End With
    
    If hideColCount < 0 Then
        hideColCount = 0
    End If
    
    MaxDataRP = maxRow()
    MaxDataCP = maxCol() + hideColCount + 1
    
    If regEx.Test(addrOfStartRow) Then
        startRow = addrA2N(addrOfStartRow)(1)
        If startRow < 2 Then
            startRow = 2
        End If
    Else
#If DebugVersion Then
    MsgBox "mySort : Error On Sort Address, You should pass Address like A2 OR $A$2 OR $A2 OR A$2"
#End If
    End If
    
    wWorkbook.Worksheets(wWorkSheetName).sort.SortFields.clear
    
    For Each e In var
        If IsNumeric(e) Then
            On Error Resume Next
            wWorkbook.Worksheets(wWorkSheetName).sort.SortFields.add key:=Range(addrN2A(e, 1) & rangeArrayDelimiter & addrN2A(e, MaxDataRP)), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        Else
#If DebugVersion Then
            MsgBox wWorkSheetName & " >> mySort 提示 : 請輸入數字!!"
#End If
        End If
    Next
    
    Application.ScreenUpdating = False
    On Error Resume Next
    With wWorkbook.Worksheets(wWorkSheetName).sort
        .SetRange Range(addrN2A(1, startRow - 1) & rangeArrayDelimiter & addrN2A(MaxDataCP, MaxDataRP))
        .Header = xlYes
        .MatchCase = True
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    On Error GoTo 0
    Application.ScreenUpdating = True
End Function
' 資料是否在裡面
Public Function hasValue(ByVal Value As Variant, ByVal addr As String) As Boolean
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    Dim timesDict As Object
    Dim tmpVar() As Variant
    Dim xItr As Long
    Dim yItr As Long
    
    If IsArray(Value) Then
#If DebugVersion Then
        MsgBox "hasValue : 禁止傳一個Array 值進行驗證，若傳入一律視為無值"
#End If
        hasValue = False
        Exit Function
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = addrRangePattern
    End With
    
    Call addDictionary(timesDict)
    
    If regEx.Test(addr) Then
        If IsArray(rGetValueByAddr(addr)) Then
            tmpVar = rGetValueByAddr(addr)
            For yItr = LBound(tmpVar, 1) To UBound(tmpVar, 1)
                For xItr = LBound(tmpVar, 2) To UBound(tmpVar, 2)
                    If Not timesDict.Exists(tmpVar(yItr, xItr)) Then
                        timesDict.add tmpVar(yItr, xItr), 1
                    Else
                        timesDict(tmpVar(yItr, xItr)) = timesDict(tmpVar(yItr, xItr)) + 1
                    End If
                    
                    If timesDict(tmpVar(yItr, xItr)) = 1 And CStr(tmpVar(yItr, xItr)) = Value Then
                        Call delDictionary(timesDict)
                        hasValue = True
                        Exit Function
                    End If
                Next
            Next
        Else
            hasValue = False
            Exit Function
        End If
    Else
#If DebugVersion Then
        MsgBox "hasValue : 錯誤位址格式。"
#End If
    End If
    
    Call delDictionary(timesDict)
    hasValue = False
End Function
'找出範圍內是否有重複資料
Public Function hasDuplicate(ByVal Value As Variant, ByVal addr As String) As Boolean
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    Dim timesDict As Object
    Dim tmpVar() As Variant
    Dim xItr As Long
    Dim yItr As Long
    
    If IsArray(Value) Then
#If DebugVersion Then
        MsgBox "hasDuplicate : 禁止傳一個Array 值進行驗證，若傳入一律視為有重複"
#End If
        hasDuplicate = True
        Exit Function
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = addrRangePattern
    End With
    
    Call addDictionary(timesDict)
    
    If regEx.Test(addr) Then
        If IsArray(rGetValueByAddr(addr)) Then
            tmpVar = rGetValueByAddr(addr)
            For yItr = LBound(tmpVar, 1) To UBound(tmpVar, 1)
                For xItr = LBound(tmpVar, 2) To UBound(tmpVar, 2)
                    If Not timesDict.Exists(tmpVar(yItr, xItr)) Then
                        timesDict.add tmpVar(yItr, xItr), 1
                    Else
                        timesDict(tmpVar(yItr, xItr)) = timesDict(tmpVar(yItr, xItr)) + 1
                    End If
                    
                    If timesDict(tmpVar(yItr, xItr)) > 1 And CStr(tmpVar(yItr, xItr)) = Value Then
                        Call delDictionary(timesDict)
                        hasDuplicate = True
                        Exit Function
                    End If
                Next
            Next
        Else
            Call delDictionary(timesDict)
            hasDuplicate = False
            Exit Function
        End If
    Else
#If DebugVersion Then
        MsgBox "hasDuplicate : 錯誤位址格式。"
#End If
    End If
    
    Call delDictionary(timesDict)
    hasDuplicate = False
End Function
' 數字和文字位置轉換
Public Function addrA2N(ByVal addr As String) As Variant ' 回傳數值陣列
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = AbsAddrPattern
    End With
    
    If regEx.Test(addr) Then
        addrA2N = Array(wWorkbook.Worksheets(wWorkSheetName).Range(addr).Column, Worksheets(wWorkSheetName).Range(addr).Row)
        On Error Resume Next
    Else
#If DebugVersion Then
        MsgBox "你填的Excel 位子是無效的，請修改"
#End If
        addrA2N = Array(-1, -1)
    End If
End Function
'欄名稱 和 位置 轉換
Public Function colN2A(ByVal Number As Long) As String
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    Dim tmpStr As String
    tmpStr = wWorkbook.Worksheets(wWorkSheetName).Cells(1, Number).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    colN2A = Left(tmpStr, Len(tmpStr) - 1)
End Function
Public Function colA2N(ByVal Address As String) As Long
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = colPattern
    End With
    
    If regEx.Test(Address) Then
        colA2N = wWorkbook.Worksheets(wWorkSheetName).Range(CStr(Address & "1")).Column
    Else
#If DebugVersion Then
        MsgBox "你填的Excel 位子是無效的，請修改"
#End If
        colA2N = -1
    End If
End Function
' 數字和文字位置轉換
Public Function addrN2A(ByVal X As Long, ByVal Y As Long) As String
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    addrN2A = CStr(wWorkbook.Worksheets(wWorkSheetName).Cells(Y, X).Address(RowAbsolute:=False, ColumnAbsolute:=False))
End Function
'插入一行
Public Function addALine(ByVal row2Insert As Long)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    wWorkbook.Worksheets(wWorkSheetName).Rows(Join(Array(row2Insert, row2Insert), rangeArrayDelimiter)).insert Shift:=xlUp, CopyOrigin:=xlFormatFromRightOrBelow
End Function
'刪除一行
Public Function delALine(ByVal row2Delete As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    wWorkbook.Worksheets(wWorkSheetName).Rows(Join(Array(row2Delete, row2Delete), rangeArrayDelimiter)).delete Shift:=xlUp
End Function
'資料複製 No Warning ~ !!
Public Function Copy(ByVal dstSheetName As String, ByVal srcAddrRange As String, ByVal dstAddr As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = addrRangePattern
    End With
    
    If regEx.Test(srcAddrRange) And Len(dstSheetName) > 0 Then
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .pattern = AbsAddrPattern
        End With
        If regEx.Test(dstAddr) Then
            wWorkbook.Worksheets(wWorkSheetName).Range(srcAddrRange).Copy _
                      Destination:=Worksheets(dstSheetName).Range(dstAddr)
        End If
    End If
End Function
'合併
Public Function mergeCells(ByVal address2Merge As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = addrPattern
    End With
    
    If regEx.Test(address2Merge) Then
        If isMergeCellsByRange(wWorkSheetName, address2Merge) <> "ErrorAddress" Then
            If Not isMergeCellsByRange(wWorkSheetName, address2Merge) Then
                wWorkbook.Worksheets(wWorkSheetName).Range(address2Merge).Merge
            Else
                MsgBox wWorkSheetName & " >> " & address2Merge & " is already Merged, so It can't merge"
            End If
        End If
    End If
End Function
'不合併
Public Function unMergeCells(ByVal addressOfMerge As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = addrPattern
    End With


    If regEx.Test(addressOfMerge) Then
        If isMergeCellsByRange(wWorkSheetName, addressOfMerge) <> "ErrorAddress" Then
            If isMergeCellsByRange(wWorkSheetName, addressOfMerge) Then
                wWorkbook.Worksheets(wWorkSheetName).Range(addressOfMerge).UnMerge
            Else
#If DebugVersion Then
                MsgBox addressOfMerge & " is not Merged, so It can't UnMerge"
#End If
            End If
        End If
    End If
End Function
'是否為合併儲存格
Public Function isMergeCells(ByVal AbsAddress As String) As Variant
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = False
        .MultiLine = False
        .IgnoreCase = False
        .pattern = AbsAddrPattern
    End With
    
    If Not regEx.Test(AbsAddress) Then
#If DebugVersion Then
        MsgBox "isMergeCells : 輸入無效Excel位址 : 請重新輸入"
#End If
        isMergeCells = "ErrorAddress"
        Exit Function
    End If
    
    If wWorkbook.Worksheets(wWorkSheetName).Range(AbsAddress).mergeCells Then
        isMergeCells = True
        Exit Function
    End If
    
    isMergeCells = False
End Function
'是否為合併儲存格 (範圍)
Public Function isMergeCellsByRange(ByVal wWorkSheetName As String, ByVal RangeAddress As String) As Variant
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = False
        .MultiLine = False
        .IgnoreCase = False
        .pattern = addrRangePattern
    End With
    
    If Not regEx.Test(RangeAddress) Then
#If DebugVersion Then
        MsgBox "isMergeCellsByRange : 輸入無效Excel位址 : 請重新輸入"
#End If
        isMergeCellsByRange = "ErrorAddress"
        Exit Function
    End If
    
    If IsNull(wWorkbook.Worksheets(wWorkSheetName).Range(RangeAddress).mergeCells) Then
        isMergeCellsByRange = True
        Exit Function
    ElseIf wWorkbook.Worksheets(wWorkSheetName).Range(RangeAddress).mergeCells Then
        isMergeCellsByRange = True
        Exit Function
    End If
    
    isMergeCellsByRange = False
End Function
' 如果確定儲存格合併，取得合併範圍
Public Function getMergeAddr(ByVal Address As String) As String
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    With regEx
        .Global = False
        .MultiLine = False
        .IgnoreCase = False
        .pattern = AbsAddrPattern
    End With
    
    If regEx.Test(Address) Then
        If TypeName(isMergeCellsByRange(wWorkSheetName, Address)) = "Boolean" Then
            If isMergeCellsByRange(wWorkSheetName, Address) Then
                getMergeAddr = wWorkbook.Worksheets(wWorkSheetName).Range(Address).MergeArea.Address(RowAbsolute:=False, ColumnAbsolute:=False)
            End If
        Else
            getMergeAddr = Address
        End If
    Else
        getMergeAddr = "ErrorAddress"
    End If
End Function
' 增加新的活頁簿
Public Static Function addSheet(ByVal SheetName As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    Dim ws As Worksheet
    Set ws = wWorkbook.Sheets.add(After:=Worksheets(Worksheets.Count))
    If (isSheetExist(SheetName)) Then
        Call delSheet(SheetName)
        Debug.Print SheetName & " is existed, so we have been deleted the oldest " & SheetName
    End If
    ws.Name = SheetName
End Function
' 查詢活頁簿
Public Static Function isSheetExist(ByVal SheetName As String, Optional ByVal wb As Workbook) As Boolean
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
   Dim sheet As Worksheet

    If wb Is Nothing Then Set wb = wWorkbook
    On Error Resume Next
    Set sheet = wb.Sheets(SheetName)
    On Error GoTo 0
    isSheetExist = Not sheet Is Nothing
End Function
'刪除活頁簿
Public Static Function delSheet(ByVal SheetName As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    If isSheetExist(SheetName) Then
        Application.DisplayAlerts = False
        wWorkbook.Worksheets(SheetName).delete
        Application.DisplayAlerts = True
    End If
End Function
Public Static Function copySheet(ByVal srcSheetName As String, ByVal dstSheetName As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    If isSheetExist(srcSheetName) And StrComp(srcSheetName, dstSheetName) <> 0 Then
        wWorkbook.Worksheets(srcSheetName).Copy After:=wWorkbook.Worksheets(dstSheetName)
    Else
        MsgBox dstSheetName & " is Exist."
    End If
End Function
' 改變活頁簿名稱
Public Static Function changeSheetName(ByVal oldSheetName As String, ByVal newSheetName As String)
    If Application.ActiveWorkbook.Name <> wWorkbook.Name Then
        wWorkbook.Activate
    End If
    
    If isSheetExist(oldSheetName) Then
        wWorkbook.Worksheets(oldSheetName).Name = newSheetName
    End If
End Function
' 複製活頁簿
' 創建一個字典
Public Function addDictionary(ByRef O As Object)
    Set O = CreateObject("Scripting.Dictionary")
    Do Until O.Count = 0
        O.RemoveAll
    Loop
End Function
' 使字典被GC
Public Function delDictionary(ByRef O As Object)
    If Not O Is Nothing Then
        O.RemoveAll
    End If
    Set O = Nothing
End Function



