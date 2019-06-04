VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CurFilePath As String
Private Sub Class_Initialize()
    CurFilePath = ActiveWorkbook.Path
End Sub
Public Function OpenFile() As String
    Dim selectedFile As String
    Dim selItems As Variant
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = ActiveWorkbook.Path
        If .Show = -1 Then
            For Each selItems In .selectedItems
                selectedFile = selItems
            Next selItems
        End If
    End With
    If Len(selectedFile) = 0 Then
        selectedFile = "No_File_Choosed"
    End If
    Debug.Print selectedFile
    CurFilePath = selectedFile
    OpenFile = selectedFile
End Function
Public Function openFolder() As String
    Dim selectedFolder As String
    Dim selItems As Variant
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ActiveWorkbook.Path
        If .Show = -1 Then
            For Each selItems In .selectedItems
                selectedFolder = selItems
            Next selItems
        End If
    End With
    If Len(selectedFolder) = 0 Then
        selectedFolder = "No_Folder_Choosed"
    End If
    Debug.Print selectedFolder
    CurFilePath = selectedFolder
    openFolder = selectedFolder
End Function
Public Function readFile() As String
    Dim tmpStr As String
    Dim allStr As String
    
    allStr = ""
    Open CurFilePath For Input As #1
    
    While Not EOF(1)
        Line Input #1, tmpStr
        allStr = allStr & vbNewLine & tmpStr
    Wend
    
    Close #1
    
    readFile = allStr
End Function
Public Function writeFile(ByVal dstDat As String, Optional ByVal dstPath As String = "")
    If Len(dstPath) > 0 Then
        CurFilePath = dstPath
    End If
    
    Open CurFilePath For Output As #1
        Write #1, dstDat
    Close #1
End Function

