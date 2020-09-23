Attribute VB_Name = "ModAscii"
'##########################################################################################
' Adaptación de la Rutina para escribir Raster-Text que en su dia programé para hacer un
' Scroll de una fuente grafica cualquiera.
' Puedes ver el codigo completo aqui:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=42666&lngWId=1
'##########################################################################################
Option Explicit

Public Type tAscii
    CodeASC As Integer
    PosX As Integer
    Width As Integer
End Type

Public MyAscii(10) As tAscii

'##########################################################################################
Public Sub IniciaAscii()
Dim i As Integer

    For i = 0 To 9
        MyAscii(i + 1).CodeASC = 48 + i
        MyAscii(i + 1).PosX = (i * 6) + 1
        MyAscii(i + 1).Width = 5
    Next i

End Sub

'##########################################################################################
Public Sub WriteMyAscii(hDCOrg As Long, hDCDest As Long, DestWidth As Long, DisplayText As String, Xdest As Long, Ydest As Long, ByRef FinishScroll As Boolean, Optional Leading As Integer = 0, Optional Scrolling As Boolean = False)
Dim i           As Integer
Dim j           As Integer
Dim CounterX    As Long
Dim UnknowAscii As Boolean
Dim OnlyChar    As String
Dim tempAsc     As tAscii
Dim TempUcase   As String
    
    If Len(DisplayText) = 0 Then
        Exit Sub
    End If

    TempUcase = UCase$(DisplayText)
    CounterX = Xdest
        
    For i = 1 To Len(TempUcase)
        OnlyChar = Mid(TempUcase, i, 1)
        UnknowAscii = True
        For j = 1 To UBound(MyAscii)
            If Asc(OnlyChar) = MyAscii(j).CodeASC Then
                tempAsc = MyAscii(j)
                UnknowAscii = False
                Exit For
            End If
        Next j
    
        'only prints necessary Text
        If CounterX > DestWidth Or CounterX < -38 Then
            If Scrolling = False Then
                Exit For
            Else
                CounterX = CounterX + tempAsc.Width - Leading
            End If
        Else
            If i = Len(TempUcase) And CounterX < -38 - Leading Then
                FinishScroll = True
            Else
                FinishScroll = False
            End If
            If UnknowAscii = False Then
                BitBlt hDCDest, CounterX, Ydest, tempAsc.Width, 8, hDCOrg, tempAsc.PosX, 10, SRCAND
                BitBlt hDCDest, CounterX, Ydest, tempAsc.Width, 8, hDCOrg, tempAsc.PosX, 1, SRCPAINT
                CounterX = CounterX + tempAsc.Width - Leading
            End If
        End If
        If i = Len(TempUcase) And CounterX < -38 - Leading Then
            FinishScroll = True
        Else
            FinishScroll = False
        End If
    Next i
    
End Sub
