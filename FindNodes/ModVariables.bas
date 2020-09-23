Attribute VB_Name = "ModVariables"
Option Explicit

Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public NumeroNodos As Long

Public Type tNodo
    X As Long
    Y As Long
End Type
Public Nodo() As tNodo

Public IndiceNodoDrag As Long

Public Type tNextNode
    lNextNode As Long
End Type

Public NumeroNextNodes As Long

Public Type TreeNode
    CurrNode As Long            'arbol binario
    NextNode() As Long          'arbol binario
    Dist() As Double            'arbol binario
    VisitNumber As Long         'Dijstra
    Distance As Double          'Dijstra
    TmpVar As Double            'Dijstra
End Type
Public TreeNodeList() As TreeNode

Public CurrDestNode As Long
Public CurrSrcNode As Long

Public nPathList As Long
Public PATHLIST() As Long

Public NumeroInterconexiones As Long

Public Type tInterconexion
    NodoInicial As Long
    NodoFinal As Long
End Type
Public Interconexion() As tInterconexion
Public InterconexionB() As tInterconexion

Public Analizado As Boolean

Public HayNodoCero As Boolean

Public ImposibleRuta As Boolean

Public FicheroCargado As Boolean

Public NombreFicheroNodos As String

Public argRender As Boolean


