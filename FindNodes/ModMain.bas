Attribute VB_Name = "ModMain"
'##########################################################################################
' Compo0Levle3.exe  Version 1.0
' Objetivo:
' Dada una lista de nodos e interconexiones, encontrar la ruta mas corta entre un nodo de
' ogrien y el de destino
'
' para la CompoVB de http://www.canalvisualbasic.net
'
' Programado por J.A.D. .:A.K.A:. v_sss
'
' Fecha de ultima revisión:14/10/2003
'
' Para una persona no muy familiarizada con las estructuras de datos este problema parece
' mas complejo de lo que en realidad es. (aunque en realidad, es complejo, se este
' familiarizado o no..xD)
' Antes de nada, decir que para el correcto desarrollo de esta Compo0Level3.exe se ha
' utilizado el algoritmo de caminos minimos de Dijkstra.
' Edsger W. Dijkstra fue un mátematico Holandes (y digo fué, porque desgraciadamente creo
' que fallecio el año pasado) que desarrollo este algoritmo alla por la decada de los 60
'
' Si has estudiado asignaturas del Departamento de LSI (Laboratorio de Sistemas Informaticos)
' como por ejemplo EDA (Estructura de Datos y Algoritmos), IEA (Introduccion a los Esquemas
' Algoritmicos), por citar algunas, comprenderás dicho algoritmo
'
' Puedes encontrar enlaces / bibliografia sobre el tema aqui:
' http://www-b2.is.tokushima-u.ac.jp/~ikeda/suuri/dijkstra/Dijkstra.shtml
' http://www.alumnos.unican.es/uc900/Algoritmo.htm
' http://www.cs.utexas.edu/users/EWD/
'
' Incluso, un proyecto final de carrera de la UPC (Universitat Politècnica de Catalunya)
' aqui: (no te sorprendas de la aplicación que se le da al algoritmo xDDD:
' http://www.bacc.info/sensefums/projecteveronica.htm
'
' Existen mas algoritmos similares, optimizados para una tarea en concreto (Kruskal, Prim, etc)
' He incluso modificaciones sobre el algoritmo original de E.W. Dijkstra que permiten
' cambiar la distancia del nodo en tiempo de ejecución, descongestión de redes, etc.
'
' Si no estas familiarizado con este Algoritmo, quizas cuando lo veas no te
' parezca "gran cosa", ya que como verás son apenas poco mas 40 miseras lineas de codigo,
' pero claro, esto siempre le ocurre a los "malos estudiantes", que una vez que ven
' la solución dicen...."Ahhhh, claro..."  (jejeje, un poco de sarcasmo no hace daño)
'
' No me extenderé en los comentarios, ya que el código se explica por si solo
' (recuerdo que tenia un profesor en la facultad que siempre decia que un a * BUEN *
' codigo no le hacian falta comentarios, ya que el codigo se explicaba por si mismo)
'
' P.D.:ya me gustaria ver a semejante elemento ante un programa de codificación genetica
'      o estructuras neuronales de autoaprendizaje...xD
'      ..... o quizas soy yo que soy muy malo programando...vete a saber...xD
'
'
' Bueno, un último comentario:
' El formato del fichero que propuso GiGaHeRz para la realización de esta aplicación
' no era el más adecuado, aunque si el más simple, por que no decirlo:
'    Nº de nodos
'    Nodo de Inicio
'    Nodo de Destino
'    Numero de interconexiones
'    Interconexion Nodo i, Nodo j
'    ...
'    Interconexion n-1 Nodo i, Nodo j
'    Interconexion n Nodo i, Nodo j
'
' Para el correcto funcionamiento del algoritmo de Dijkstra la estructura para "crear"
' un ARBOL BINARIO de nodos deberia ser:
' Nodo i
'     Hijo 1 del nodo i (K)
'     Hijo 2 del nodo i (P)
'     Hijo n del nodo i (N)
' Nodo (K)
'   Hijo 1 del nodo (K)  (K1)
'   Hijo n del nodo (K)  (P1)
' ....
'
' Para simplificar el algoritmo, se ha supuesto que todos los nodos estan a equidistantes
' unos de otros.
'##########################################################################################


Option Explicit
Private OrdenN()    As Long
Private OrdenNB()   As Long

'##########################################################################################
' El programa comienza aqui. Si se pasa el argumento -render se visualizará la malla
' de nodos.
'##########################################################################################
Sub Main()
    'Debug.Print Command
    If UCase(Command$) = "-RENDER" Then
        argRender = True
    Else
        argRender = False
    End If
    
    NombreFicheroNodos = AppPath & "datos.txt"
    
    Inicializa
    
End Sub

'##########################################################################################
Private Sub Inicializa()
    
    'Si no se pasa ningun argumento:
    '   Se lee el fichero de datos
    '   Se aplica el algoritmo
    '   Se Salvan los datos en Resultados.txt
    '   Finaliza el programa
    If argRender = False Then
        LeerFicheroDatos (NombreFicheroNodos)
        ImposibleRuta = False
        
        If Dijkstra(CurrSrcNode, CurrDestNode) = False Then
            ImposibleRuta = True
        End If
        
        GrabaFicheroResultados
        End
    
    Else
        
        frmMain.Show
    
    End If
End Sub

'##########################################################################################
Public Sub GrabaFicheroResultados()
On Error GoTo ErrorHandler
    
    Dim nF          As Integer
    Dim i           As Long
    Dim datoRuta    As String
    
    nF = FreeFile
    
    Open AppPath & "Resultados.txt" For Output As #nF
'        Print #nF, "Nodo Origen:" & CurrSrcNode
'        Print #nF, "Nodo Destino:" & CurrDestNode
'        Print #nF, "Ruta:"
'        Print #nF, "-----"
        If ImposibleRuta = True Then
            Print #nF, "No se pudo encontrar ninguna ruta :("
        Else
            datoRuta = ""
            For i = 1 To (nPathList - 1)
                datoRuta = datoRuta & TreeNodeList(PATHLIST(i)).CurrNode & "," '& TreeNodeList(PATHLIST(i + 1)).CurrNode
            Next i
            datoRuta = datoRuta & TreeNodeList(PATHLIST(i)).CurrNode

            Print #nF, datoRuta

        End If
    Close #nF
    
    Exit Sub
ErrorHandler:

    MsgBox "Se ha producido el error:" & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "El programa se cerrará ahora", , "Error en fichero de Datos"
    End

End Sub

'##########################################################################################
' Aqui esta el meollo "previo" de todo este sarao...xD
' Se convierten los datos del fichero datos.txt a la estructura de arbol binario que se
' ha descrito arriba
'##########################################################################################
Public Sub LeerFicheroDatos(NombreFichero As String)
On Error GoTo ErrorHandler
    Dim nF          As Integer
    Dim i           As Long
    Dim j           As Long
    Dim nodoInicio  As Long
    Dim NodoFinal   As Long
    Dim posComa     As Integer
    Dim DatoFichero As String
    
    nF = FreeFile
    
    Open NombreFichero For Input As #nF
        'numero de nodos
        Line Input #nF, DatoFichero
        If IsNumeric(DatoFichero) = True Then
            NumeroNodos = CLng(DatoFichero)
        Else
            MsgBox "Error en el fichero de datos. El número de nodos no es un dato numérico!!!" & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "El programa se cerrará ahora", , "Error en fichero de Datos"
            End
        End If
        
        'Nodo de origen
        Line Input #nF, DatoFichero
        If IsNumeric(DatoFichero) = True Then
            CurrSrcNode = CLng(DatoFichero)
        Else
            MsgBox "Error en el fichero de datos. El Nodo de Origen no es un dato numérico!!!" & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "El programa se cerrará ahora", , "Error en fichero de Datos"
            End
        End If
        
        'Nodo de destino
        Line Input #nF, DatoFichero
        If IsNumeric(DatoFichero) = True Then
            CurrDestNode = CLng(DatoFichero)
        Else
            MsgBox "Error en el fichero de datos. El Nodo de Destino no es un dato numérico!!!" & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "El programa se cerrará ahora", , "Error en fichero de Datos"
            End
        End If

        'Nodo de destino
        Line Input #nF, DatoFichero
        If IsNumeric(DatoFichero) = True Then
            NumeroInterconexiones = CLng(DatoFichero)
        Else
            MsgBox "Error en el fichero de datos. El Numero de interconexiones no es un dato numérico!!!" & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "El programa se cerrará ahora", , "Error en fichero de Datos"
            End
        End If
        
        'Ahora Lee las interconexiones
        ReDim Interconexion(NumeroInterconexiones)
        HayNodoCero = False
        For i = 0 To NumeroInterconexiones - 1
            Line Input #nF, DatoFichero
            posComa = InStr(1, DatoFichero, ",")
            If posComa = 0 Then
                MsgBox "Error en el fichero de datos. El formato de las interconexiones entre nodos no es el correcto!!!" & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "El programa se cerrará ahora", , "Error en fichero de Datos"
                End
            Else
                'nodo de inicio (interconexion)
                If IsNumeric(Mid(DatoFichero, 1, posComa - 1)) = False Then
                    MsgBox "Error en el fichero de datos. El formato de las interconexiones entre nodos no es el correcto!!!" & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "El programa se cerrará ahora", , "Error en fichero de Datos"
                    End
                Else
                    
                    nodoInicio = Mid(DatoFichero, 1, posComa - 1)
                    If nodoInicio = 0 Then HayNodoCero = True
                End If
                'nodo final (interconexion)
                If IsNumeric(Mid(DatoFichero, posComa + 1)) = False Then
                    MsgBox "Error en el fichero de datos. El formato de las interconexiones entre nodos no es el correcto!!!" & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "El programa se cerrará ahora", , "Error en fichero de Datos"
                    End
                Else
                    
                    NodoFinal = Mid(DatoFichero, posComa + 1)
                    If NodoFinal = 0 Then HayNodoCero = True
                End If
            
                Interconexion(i).NodoInicial = nodoInicio
                Interconexion(i).NodoFinal = NodoFinal
            End If
        Next i
    Close #nF

    If argRender = True Then
        frmMain.lblFicheroNodos.Caption = "Fichero Nodos:" & NombreFichero
    End If
    
    If HayNodoCero = False Then
        'ReDim Preserve Interconexion(UBound(Interconexion) + 1)
        Interconexion(UBound(Interconexion)).NodoFinal = -1
        Interconexion(UBound(Interconexion)).NodoInicial = -1
    End If

'    Debug.Print "*******"
'    For i = 1 To NumeroInterconexiones
'        Debug.Print i & ": (" & Interconexion(i).NodoInicial & "," & Interconexion(i).nodoFinal & ")"
'    Next i
    ReDim InterconexionB(UBound(Interconexion))
        
    For i = 0 To UBound(Interconexion)
        InterconexionB(i) = Interconexion(i)
    Next i
    
    QuickSortInterconexiones Interconexion, 0, NumeroInterconexiones
    
    QuickSortInterconexionesB InterconexionB, 0, NumeroInterconexiones
    
'    Debug.Print "----Interconexion----"
'    For i = 0 To NumeroInterconexiones - 1
'        Debug.Print i & ": (" & Interconexion(i).NodoInicial & "," & Interconexion(i).NodoFinal & ")"
'    Next i
'
'
'    Debug.Print "----InterconexionB----"
'    For i = 0 To NumeroInterconexiones - 1
'        Debug.Print i & ": (" & InterconexionB(i).NodoInicial & "," & InterconexionB(i).NodoFinal & ")"
'    Next i

    NumeroNextNodes = OrdenNGraph + OrdenNBGraph
    
    ReDim TreeNodeList(NumeroNodos)
    
    For i = 0 To NumeroNodos
        For j = 0 To NumeroNextNodes - 1
            ReDim TreeNodeList(i).NextNode(j)
            ReDim TreeNodeList(i).Dist(j)
        Next j
    Next i
    
    For i = 0 To NumeroNodos
        For j = 0 To NumeroNextNodes - 1
            TreeNodeList(i).CurrNode = -1
            TreeNodeList(i).NextNode(j) = -1
            TreeNodeList(i).Dist(j) = 10    'Distancia igual para todos los nodos para
                                            'simplificar el algoritmo
        Next j
    Next i
    
    IniciaBSP   '<---- Construye el arbol binario
    
    Exit Sub
ErrorHandler:
    MsgBox "Se ha producido el error:" & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "El programa se cerrará ahora", , "Error en fichero de Datos"

    End
End Sub
'##########################################################################################
Private Sub IniciaBSP()
Dim i           As Long
Dim j           As Long
Dim k           As Long
Dim aux         As Long
Dim auxb        As Long
Dim contador    As Long
Dim lNodoFinal  As Long

    'interconexion
    aux = 0
    contador = 0
    auxb = 1
    Do
        j = OrdenN(auxb)
        For i = 1 To j
            contador = contador + 1
            TreeNodeList(aux).CurrNode = Interconexion(contador).NodoInicial
            TreeNodeList(aux).NextNode(i - 1) = Interconexion(contador).NodoFinal
        Next i
        auxb = auxb + 1
        aux = aux + 1
    Loop Until auxb > UBound(OrdenN)

    'interconexionB
    'aux = 0
    contador = 0
    auxb = 1
    Do
        j = OrdenNB(auxb)
        lNodoFinal = InterconexionB(contador + 1).NodoFinal
        For k = 0 To UBound(TreeNodeList) - 1
            If TreeNodeList(k).CurrNode = lNodoFinal Or TreeNodeList(k).CurrNode = -1 Then
                aux = k
                Exit For
            End If
        Next k
        k = 0
        For i = 1 To j
            k = i - 1
            contador = contador + 1
            TreeNodeList(aux).CurrNode = InterconexionB(contador).NodoFinal
            Do Until TreeNodeList(aux).NextNode(k) = -1
                k = k + 1
            Loop
            TreeNodeList(aux).NextNode(k) = InterconexionB(contador).NodoInicial
        Next i
        auxb = auxb + 1
        aux = aux + 1
    Loop Until auxb > UBound(OrdenNB)
    
    QuickSortBSP TreeNodeList, 0, UBound(TreeNodeList)
    
End Sub

'##########################################################################################
Private Function OrdenNGraph() As Long
Dim i As Long
Dim aux As Long
Dim lNodoA As Long
Dim lNodoB As Long
Dim TempOrdenN() As Long
    ReDim OrdenN(0)
    
    aux = 1
    For i = 1 To UBound(Interconexion) - 1
        lNodoA = Interconexion(i).NodoInicial
        lNodoB = Interconexion(i + 1).NodoInicial
        If lNodoA = lNodoB Then
            aux = aux + 1
        Else
           ReDim Preserve OrdenN(UBound(OrdenN) + 1)
           OrdenN(UBound(OrdenN)) = aux
           aux = 1
        End If
    Next i
    ReDim Preserve OrdenN(UBound(OrdenN) + 1)
    OrdenN(UBound(OrdenN)) = aux
    
    ReDim TempOrdenN(UBound(OrdenN))
    

    For i = 0 To UBound(OrdenN)
        TempOrdenN(i) = OrdenN(i)
    Next i
    
    QuickSortOrdenN OrdenN, 1, UBound(OrdenN)

'    Debug.Print "///////"
'    For i = 1 To UBound(OrdenN)
'        Debug.Print OrdenN(i)
'    Next i
    
    OrdenNGraph = OrdenN(UBound(OrdenN))
    
    For i = 0 To UBound(OrdenN)
        OrdenN(i) = TempOrdenN(i)
    Next i
        
'    Debug.Print "///////"
'    For i = 1 To UBound(OrdenN)
'        Debug.Print OrdenN(i)
'    Next i
    
End Function

'##########################################################################################
Private Function OrdenNBGraph() As Long
Dim i As Long
Dim aux As Long
Dim lNodoA As Long
Dim lNodoB As Long
Dim TempOrdenN() As Long
    ReDim OrdenNB(0)
    
    aux = 1
    For i = 1 To UBound(InterconexionB) - 1
        lNodoA = InterconexionB(i).NodoFinal
        lNodoB = InterconexionB(i + 1).NodoFinal
        If lNodoA = lNodoB Then
            aux = aux + 1
        Else
           ReDim Preserve OrdenNB(UBound(OrdenNB) + 1)
           OrdenNB(UBound(OrdenNB)) = aux
           aux = 1
        End If
    Next i
    ReDim Preserve OrdenNB(UBound(OrdenNB) + 1)
    OrdenNB(UBound(OrdenNB)) = aux
    
    ReDim TempOrdenN(UBound(OrdenNB))
    

    For i = 0 To UBound(OrdenNB)
        TempOrdenN(i) = OrdenNB(i)
    Next i
    
    QuickSortOrdenN OrdenNB, 1, UBound(OrdenNB)

'    Debug.Print "///////"
'    For i = 1 To UBound(OrdenNB)
'        Debug.Print OrdenNB(i)
'    Next i
    
    OrdenNBGraph = OrdenNB(UBound(OrdenNB))
    
    For i = 0 To UBound(OrdenNB)
        OrdenNB(i) = TempOrdenN(i)
    Next i
        
'    Debug.Print "///////"
'    For i = 1 To UBound(OrdenNB)
'        Debug.Print OrdenNB(i)
'    Next i
    
End Function

' Viva la Recursividad xDDD

'##########################################################################################
Private Sub QuickSortOrdenN(ByRef vntArr() As Long, _
    lngLeft As Long, lngRight As Long)

    Dim i           As Long
    Dim j           As Long
    Dim lngMid      As Long
    Dim vntTestVal  As Variant
    Dim vntTemp     As Long
    
    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2
        vntTestVal = vntArr(lngMid)
        i = lngLeft
        j = lngRight
        Do
            Do While vntArr(i) < vntTestVal
                i = i + 1
            Loop
            Do While vntArr(j) > vntTestVal
                j = j - 1
            Loop
            If i <= j Then
                vntTemp = vntArr(j)
                vntArr(j) = vntArr(i)
                vntArr(i) = vntTemp
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j

        If j <= lngMid Then
            Call QuickSortOrdenN(vntArr, lngLeft, j)
            Call QuickSortOrdenN(vntArr, i, lngRight)
        Else
            Call QuickSortOrdenN(vntArr, i, lngRight)
            Call QuickSortOrdenN(vntArr, lngLeft, j)
        End If
    End If
End Sub

'##########################################################################################
Private Sub QuickSortInterconexiones(ByRef vntArr() As tInterconexion, _
    lngLeft As Long, lngRight As Long)

    Dim i           As Long
    Dim j           As Long
    Dim lngMid      As Long
    Dim vntTestVal  As Variant
    Dim vntTemp     As tInterconexion
    
    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2
        
        vntTestVal = vntArr(lngMid).NodoInicial
        i = lngLeft
        j = lngRight
        Do
        
            Do While vntArr(i).NodoInicial < vntTestVal
                i = i + 1
            Loop
            Do While vntArr(j).NodoInicial > vntTestVal
                j = j - 1
            Loop
            If i <= j Then
                vntTemp = vntArr(j)
                vntArr(j) = vntArr(i)
                vntArr(i) = vntTemp
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j

        If j <= lngMid Then
            Call QuickSortInterconexiones(vntArr, lngLeft, j)
            Call QuickSortInterconexiones(vntArr, i, lngRight)
        Else
            Call QuickSortInterconexiones(vntArr, i, lngRight)
            Call QuickSortInterconexiones(vntArr, lngLeft, j)
        End If
    End If
End Sub

'##########################################################################################
Private Sub QuickSortBSP(ByRef vntArr() As TreeNode, _
    lngLeft As Long, lngRight As Long)

    Dim i           As Long
    Dim j           As Long
    Dim lngMid      As Long
    Dim vntTestVal  As Variant
    Dim vntTemp     As TreeNode
    
    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2
            vntTestVal = vntArr(lngMid).CurrNode
        i = lngLeft
        j = lngRight
        Do
            Do While vntArr(i).CurrNode < vntTestVal
                i = i + 1
            Loop
            Do While vntArr(j).CurrNode > vntTestVal
                j = j - 1
            Loop
            If i <= j Then
                vntTemp = vntArr(j)
                vntArr(j) = vntArr(i)
                vntArr(i) = vntTemp
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j

        If j <= lngMid Then
            Call QuickSortBSP(vntArr, lngLeft, j)
            Call QuickSortBSP(vntArr, i, lngRight)
        Else
            Call QuickSortBSP(vntArr, i, lngRight)
            Call QuickSortBSP(vntArr, lngLeft, j)
        End If
    End If
End Sub

'##########################################################################################
Private Sub QuickSortInterconexionesB(ByRef vntArr() As tInterconexion, _
    lngLeft As Long, lngRight As Long)

    Dim i           As Long
    Dim j           As Long
    Dim lngMid      As Long
    Dim vntTestVal  As Variant
    Dim vntTemp     As tInterconexion
    
    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2
            vntTestVal = vntArr(lngMid).NodoFinal
        i = lngLeft
        j = lngRight
        Do
            Do While vntArr(i).NodoFinal < vntTestVal
                i = i + 1
            Loop
            Do While vntArr(j).NodoFinal > vntTestVal
                j = j - 1
            Loop
            If i <= j Then
                vntTemp = vntArr(j)
                vntArr(j) = vntArr(i)
                vntArr(i) = vntTemp
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j

        If j <= lngMid Then
            Call QuickSortInterconexionesB(vntArr, lngLeft, j)
            Call QuickSortInterconexionesB(vntArr, i, lngRight)
        Else
            Call QuickSortInterconexionesB(vntArr, i, lngRight)
            Call QuickSortInterconexionesB(vntArr, lngLeft, j)
        End If
    End If
End Sub

'##########################################################################################
' Devuelve el valor minimo entre 2 numeros
' Nota: quizas el VB6 ya tenga esta instrucción, en mi VB5, no :(
Private Function MIN(A As Double, B As Double) As Double
    If A < B Then MIN = A
    If A > B Then MIN = B
    If A = B Then MIN = A
End Function

'##########################################################################################
' Nota:Dejo los comentario originales del algoritmo de Dijkstra que tenia perdido
'      por mi disco duro
'      No recuerdo la URL de donde lo descargé, sorries
Public Function Dijkstra(NodeSrc As Long, NodeDest As Long) As Boolean

    Dim i                  As Long
    Dim j                  As Long
    Dim bRunning           As Boolean
    Dim CurrentVisitNumber As Long   'Which visit the current node will be
    Dim CurrNode           As Long   'Which node we are scanning...
    Dim LowestNodeFound    As Long   'For when we are searching for the lowest temporary value
    Dim LowestValFound     As Double 'For above variable
    Dim cuentaCero         As Integer
    
    If HayNodoCero = True Then
        cuentaCero = 0
    Else
        cuentaCero = 1
    End If
    
    If NodeSrc = NodeDest Then
        'we're already there...
        nPathList = 2
        ReDim PATHLIST(2) As Long
        PATHLIST(1) = NodeSrc
        PATHLIST(2) = NodeDest
        Dijkstra = True
        Exit Function
    End If

    'A)Setup all the data we need
    For i = cuentaCero To NumeroNodos
        TreeNodeList(i).VisitNumber = -1 '-1 indicates not visited
        TreeNodeList(i).Distance = -1    'Unknown distance
        TreeNodeList(i).TmpVar = 99999   'A high number that can easily be beaten
    Next i
    
    'B)Set the first variable
    TreeNodeList(NodeSrc).VisitNumber = 1
    CurrentVisitNumber = 1 'Initialise
    CurrNode = NodeSrc
    TreeNodeList(NodeSrc).Distance = 0
    TreeNodeList(NodeSrc).TmpVar = 0

    'C)Start scanning
    'We're going to keep looping till we find the destination
    Do While bRunning = False
        '1. Go to each node that the current one touches
                'and make it's temporary variable = source distance + weight of the arc
                For j = 0 To NumeroNextNodes - 1
                    If Not (TreeNodeList(CurrNode).NextNode(j) = -1) Then TreeNodeList(TreeNodeList(CurrNode).NextNode(j)).TmpVar = MIN(TreeNodeList(CurrNode).Dist(j) + TreeNodeList(CurrNode).Distance, TreeNodeList(TreeNodeList(CurrNode).NextNode(j)).TmpVar)
                Next j
                
        '2. Decide which node has the lowest temporary variable (Free choice if multiple)
                LowestValFound = 100999 'Hopefully the graph isn't this big :)
                For i = cuentaCero To NumeroNodos  'If we have more than 1000-2000 nodes this part will be horribly slow...
                    If (TreeNodeList(i).TmpVar <= LowestValFound) And (TreeNodeList(i).TmpVar >= 0) And (TreeNodeList(i).VisitNumber < 0) Then 'make sure we ignore the -1's and visited nodes
                        'We have a new lowest value
                        LowestValFound = TreeNodeList(i).TmpVar
                        LowestNodeFound = i
                    End If
                Next i
                '**NB: If there are multiple lowest values then this method will choose the last one found...
        
        '3. Mark this node with the next visit number and copy the tmpvar -> distance
                CurrentVisitNumber = CurrentVisitNumber + 1
                TreeNodeList(LowestNodeFound).VisitNumber = CurrentVisitNumber
                TreeNodeList(LowestNodeFound).Distance = TreeNodeList(LowestNodeFound).TmpVar
                CurrNode = LowestNodeFound 'Copy the variable for next time...
        
        '4. If this node IS NOT the destination then go onto the next iteration...
                If CurrNode = NodeDest Then
                    bRunning = True 'We've gotten to the destination
                Else
                    bRunning = False 'Still not there yet
                End If
    Loop
    
    'D Work out the route that was taken...
    bRunning = False
    CurrNode = NodeDest 'Start at the end, and work backwards...
    Dim lngTimeTaken As Long
    lngTimeTaken = GetTickCount
    
    nPathList = 1
    ReDim PATHLIST(nPathList) As Long
    PATHLIST(1) = NodeDest 'Put the first node in...
    
        Do While bRunning = False
            'First we check that the current node isn't actually the start
                'because if it is then we've found the path already
                If CurrNode = NodeSrc Then
                    bRunning = True
                    GoTo EndGoal:
                ElseIf GetTickCount - lngTimeTaken > 1000 Then
                    'Break out if we haven't found a solution in under 1 second
                    bRunning = True
                    Dijkstra = False
                    Exit Function
                    GoTo EndGoal:
                End If
        
            'Scan through each node that we visited
            
            For j = 0 To NumeroNextNodes - 1
                If (TreeNodeList(CurrNode).NextNode(j) >= 0) Then 'Only if there is a node in this direction
                    If (TreeNodeList(TreeNodeList(CurrNode).NextNode(j)).VisitNumber >= 0) Then 'Only if we visited this node...
                        If TreeNodeList(CurrNode).Distance - TreeNodeList(TreeNodeList(CurrNode).NextNode(j)).Distance = TreeNodeList(CurrNode).Dist(j) Then
                            'NextNode(0) is part of the route home
                                nPathList = nPathList + 1
                                ReDim Preserve PATHLIST(nPathList) As Long
                                PATHLIST(nPathList) = TreeNodeList(CurrNode).NextNode(j)
                                CurrNode = TreeNodeList(CurrNode).NextNode(j)
                                Exit For
                        End If
                    End If
                End If

            Next j
        
        Loop
EndGoal:
'For ease of use we're going to invert the array.
    'currently we go Dest-Source, Source-Dest is more useful/easier
    Dim TmpArray() As Long
    ReDim TmpArray(nPathList) As Long
    For i = nPathList To 1 Step -1
        TmpArray(i) = PATHLIST(((nPathList - i) + 1))
    Next i
    For i = 1 To nPathList
        PATHLIST(i) = TmpArray(i)
    Next i
    
Dijkstra = True
End Function

