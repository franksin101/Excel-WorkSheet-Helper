VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "�i�����"
   ClientHeight    =   1680
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   6576
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const barWidth = 300

Private Sub UserForm_Initialize()
    DoEvents
    Call UpdateBar(1, "�}�l����A��l�ƳB�z��...")
    DoEvents
End Sub

Public Function UpdateBar(Optional ByVal Val As Double = 100, Optional ByVal msg As String = "�B�z��...", Optional ByVal isWait As Boolean = False)
    DoEvents
    
    Frame1.Caption = msg
    Bar.Width = barWidth * (Val / 100)
    Bar.Caption = Format(Val, "0.00") & "%" ' �䴩�p�ƪ��
    
    If Val = 100 Then
        If isWait Then
            MsgBox "����!"
        End If
        Unload Me
    End If
    
    DoEvents
End Function
Public Property Let Title(ByVal t As String)
    Me.Caption = t
    DoEvents
End Property

