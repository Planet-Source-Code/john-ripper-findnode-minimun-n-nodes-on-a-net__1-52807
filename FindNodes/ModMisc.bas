Attribute VB_Name = "ModMisc"
Option Explicit
Dim SDC As Long
Dim DDC As Long
Public Sub ClearBackBuffer()
    SDC = frmWork.PicClean.hdc
    DDC = frmWork.PicBakBuffer.hdc
    BitBlt DDC, 0, 0, frmWork.PicBakBuffer.ScaleWidth, frmWork.PicBakBuffer.ScaleHeight, SDC, 0, 0, SRCCOPY
End Sub
Public Sub Flip()
    SDC = frmWork.PicBakBuffer.hdc
    DDC = frmMain.Picture1.hdc
    BitBlt DDC, 0, 0, frmWork.PicBakBuffer.ScaleWidth, frmWork.PicBakBuffer.ScaleHeight, SDC, 0, 0, SRCCOPY
End Sub
Public Function FileExists(Path$) As Boolean
    Dim X As Integer
    X = FreeFile(1)

    On Error Resume Next
    Open Path$ For Input As X
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close X
End Function
Function AppPath() As String
Static MiAppPath As String
    
    If MiAppPath = "" Then  'Esta sin Cargar?
        MiAppPath = App.Path
        If Right$(MiAppPath, 1) <> "\" Then 'No es Directorio Raiz
            MiAppPath = MiAppPath & "\"
        End If
    End If
    AppPath = MiAppPath
End Function

