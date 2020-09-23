VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10155
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDialog1 
      Left            =   3600
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer tmrRender 
      Enabled         =   0   'False
      Left            =   3120
      Top             =   6240
   End
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   0
      ScaleHeight     =   421
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   673
      TabIndex        =   0
      Top             =   0
      Width           =   10155
   End
   Begin VB.Label lblFicheroNodos 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3180
      TabIndex        =   2
      Top             =   6420
      Width           =   6975
   End
   Begin VB.Label lblStatusNodoDrag 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6420
      Width           =   2775
   End
   Begin VB.Menu mArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuArchivo 
         Caption         =   "Abrir Archivo de Nodos..."
         Index           =   0
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "Salir"
         Index           =   2
      End
   End
   Begin VB.Menu mnuAnalizar 
      Caption         =   "Analizar"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##########################################################################################
' Este formulario solo será visible si se especifica el parametro -render al ejecutar el
' progama
'##########################################################################################
Option Explicit

'##########################################################################################
Private Sub Form_Load()
    
    Me.Caption = App.EXEName & ".exe  .: Por v_sss :. Ronda #0 para competición http://www.canalvisualbasic.net"
    
    IniciaAscii 'inicia tabla look-up para BLITear el raster-text de los nodos
    
End Sub

'##########################################################################################
Private Sub InicializaB()

    IndiceNodoDrag = -1     'no se tiene ningun Nodo DRAGeado
    
    Analizado = False       'Aun no se ha realizado el análisis
    
    FicheroCargado = False  'Usa tu imaginación....^_^
    
    If FileExists(NombreFicheroNodos) = True Then
        
        LeerFicheroDatos (NombreFicheroNodos)   'carga el fichero de datos
                                                'e inicializa el arbol binario
        ReDim Nodo(NumeroNodos)
        
        InitNodos                               'Genera Nodos "aleatorios" para Renderizarlos
                                                'posteriormente
        
        tmrRender.Interval = 100                'activa el timer para la renderización
        tmrRender.Enabled = True
    
    End If
    
End Sub

'##########################################################################################
'Genera Nodos "aleatorios" para Renderizarlos posteriormente
'##########################################################################################
Private Sub InitNodos()
Dim i       As Integer
Dim xRnd    As Long     'Posición X aleatoria
Dim yRnd    As Long     'Posición Y aleatoria

    Randomize           'inicia "semilla"
    
    For i = 0 To UBound(Nodo)
        xRnd = CLng(Int((Picture1.ScaleWidth * Rnd) + 1))
        yRnd = CLng(Int((Picture1.ScaleHeight * Rnd) + 1))
        Nodo(i).X = xRnd
        Nodo(i).Y = yRnd
    Next i
    
    FicheroCargado = True
End Sub

'##########################################################################################
'Renderiza los nodos y los arcos que los unen. Tambien se imprime el Numero correspondiente
'del Nodo en cuestión.
'El nodo de origen esta representado de color Verde
'El nodo de destino esta representado de color Rojo
'El volcado se realiza en un BackBuffer y se BLITea desde alli al MainBuffer mediante FLIP
'Esto se hace asi para evitar parpadeos y ver correctamente el Render, ya que todos los
'cambios se realizan en el BackBuffer (invisible para el user)
'##########################################################################################
Private Sub RenderNodos()
Dim i       As Integer
Dim j       As Integer
Dim cRojo   As Integer
Dim cVerde  As Integer
Dim cAzul   As Integer
Dim cuentaCero As Integer
    
    ClearBackBuffer     'Borra BackBuffer
    
    If HayNodoCero = True Then
        cuentaCero = 0
    Else
        cuentaCero = 1
    End If
    
    For i = cuentaCero To UBound(Nodo)
           
        'Pinta las lineas entre nodos:
        For j = 0 To NumeroNextNodes - 1
            If Not (TreeNodeList(i).NextNode(j) = -1) Then frmWork.PicBakBuffer.Line (Nodo(TreeNodeList(i).CurrNode).X, Nodo(TreeNodeList(i).CurrNode).Y)-(Nodo(TreeNodeList(i).NextNode(j)).X, Nodo(TreeNodeList(i).NextNode(j)).Y), RGB(0, 0, 255)
        Next j
        
        'Pinta la ruta mas corta si ya se ha procedido al analisis
        If Analizado = True Then
            For j = 1 To (nPathList - 1)
                frmWork.PicBakBuffer.Line (Nodo(TreeNodeList(PATHLIST(j)).CurrNode).X, Nodo(TreeNodeList(PATHLIST(j)).CurrNode).Y)-(Nodo(TreeNodeList(PATHLIST(j + 1)).CurrNode).X, Nodo(TreeNodeList(PATHLIST(j + 1)).CurrNode).Y), RGB(0, 255, 0)
            Next j
        End If
        
        'Nodo de Origen: (verde)
        If TreeNodeList(i).CurrNode = CurrSrcNode Then
            cRojo = 0
            cVerde = 255
            cAzul = 64
        'Nodo de Destion (Rojo)
        ElseIf TreeNodeList(i).CurrNode = CurrDestNode Then
            cRojo = 255
            cVerde = 0
            cAzul = 64
        'Resto de Nodos (Blancos)
        Else
            cRojo = 255
            cVerde = 255
            cAzul = 255
        End If
        
        'Los nodos de origen y destino se pintan mas "grandes"
        'Si, en este punto ya se que se pudo usar PSET con un determinado DrawWith,
        'pero el PSET es increiblemente Lento
        If TreeNodeList(i).CurrNode = CurrSrcNode Or TreeNodeList(i).CurrNode = CurrDestNode Then
            SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X - 1, Nodo(i).Y - 1, RGB(cRojo, cVerde, cAzul)
            SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X + 1, Nodo(i).Y - 1, RGB(cRojo, cVerde, cAzul)
            SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X - 1, Nodo(i).Y + 1, RGB(cRojo, cVerde, cAzul)
            SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X + 1, Nodo(i).Y + 1, RGB(cRojo, cVerde, cAzul)
            SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X - 2, Nodo(i).Y, RGB(cRojo, cVerde, cAzul)
            SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X + 2, Nodo(i).Y, RGB(cRojo, cVerde, cAzul)
            SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X, Nodo(i).Y + 2, RGB(cRojo, cVerde, cAzul)
            SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X, Nodo(i).Y - 2, RGB(cRojo, cVerde, cAzul)
        End If
        
        SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X - 1, Nodo(i).Y, RGB(cRojo, cVerde, cAzul)
        SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X + 1, Nodo(i).Y, RGB(cRojo, cVerde, cAzul)
        SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X, Nodo(i).Y + 1, RGB(cRojo, cVerde, cAzul)
        SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X, Nodo(i).Y - 1, RGB(cRojo, cVerde, cAzul)
        
        SetPixel frmWork.PicBakBuffer.hdc, Nodo(i).X, Nodo(i).Y, RGB(cRojo, cVerde, cAzul)
    
        'Escribe el Texto del Nodo Adecuado
        WriteMyAscii frmWork.PicNumeros.hdc, frmWork.PicBakBuffer.hdc, frmWork.PicBakBuffer.ScaleWidth, Str(TreeNodeList(i).CurrNode), Nodo(i).X + 5, Nodo(i).Y - 3, False
    Next i
    
    'Vuelca del BackBuffer al MainBuffer
    Flip
End Sub

'##########################################################################################
'Formula clasica para hayar la distancia entre dos puntos en un plano
'##########################################################################################
Private Function FormulaDistancia2D(X As Long, Y As Long, X1 As Long, Y1 As Long) As Long
    FormulaDistancia2D = Sqr(((X - X1) ^ 2) + ((Y - Y1) ^ 2))
End Function

'##########################################################################################
'Funcion que devolverá el Nodo mas proximo al hacer click con el ratón en la zona de render
'##########################################################################################
Private Function BuscaNodoMasProximo(X As Long, Y As Long) As Long
Dim i           As Long
Dim aux         As Long
Dim Distancia2D As Long
Dim cuentaCero  As Integer
   
    If HayNodoCero = True Then
        cuentaCero = 0
    Else
        cuentaCero = 1
    End If
    
    Distancia2D = 9999999
    
    For i = cuentaCero To UBound(Nodo)
        If FormulaDistancia2D(Nodo(i).X, Nodo(i).Y, X, Y) < Distancia2D Then
            Distancia2D = FormulaDistancia2D(Nodo(i).X, Nodo(i).Y, X, Y)
            aux = i
        End If
    Next i

    BuscaNodoMasProximo = aux
End Function

'##########################################################################################
Private Sub Form_Paint()
    Static A As Integer
    
    If A <> -1 Then
        A = -1
        InicializaB
    End If
End Sub

'##########################################################################################
Private Sub mnuAnalizar_Click()
    If Dijkstra(CurrSrcNode, CurrDestNode) = False Then
        ImposibleRuta = True
        GrabaFicheroResultados
        MsgBox "Se generó el archivo " & AppPath & "Resultados.txt" & vbCrLf & vbCrLf & "No se pudo encontrar ninguna ruta :(", , App.EXEName
    Else
        Analizado = True
        ImposibleRuta = False
        GrabaFicheroResultados
        MsgBox "Se generó el archivo " & AppPath & "Resultados.txt" & vbCrLf & vbCrLf & "con la información de la ruta mas corta", , App.EXEName
    End If
End Sub

'##########################################################################################
Private Sub mnuArchivo_Click(Index As Integer)
On Error GoTo ErrorHandler
    
Dim lret        As Long
Dim lFichero    As String
    
    Select Case Index
        Case 0 'Cargar un fichero de nodos de una ubicación especifica
            CDialog1.InitDir = AppPath
            CDialog1.DialogTitle = "Seleccione el fichero un fichero de Nodos"
            CDialog1.Filter = "Archivos de texto (*.txt)|*.txt"
            CDialog1.ShowOpen
            lFichero = CDialog1.filename
            If Trim(lFichero) <> "" Then
                NombreFicheroNodos = lFichero
                LeerFicheroDatos NombreFicheroNodos
                tmrRender.Enabled = False
            
                IndiceNodoDrag = -1
                Analizado = False
                FicheroCargado = True
                ReDim Nodo(NumeroNodos)
                InitNodos

                tmrRender.Interval = 100
                tmrRender.Enabled = True
            End If

        Case 2 'Salir
            lret = MsgBox("¿Estas segur@ que quieres salir?", vbYesNo, App.EXEName)
            If lret = vbYes Then
                End
            End If
    End Select
    
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 32755
            'Se pulso el botón Cancel en el DialogBox
            Exit Sub
        Case Else
            MsgBox "Ocurrio el siguiente error:" & Err.Number & " - " & Err.Description, , App.EXEName
            Exit Sub
    End Select
End Sub

'##########################################################################################
'Estamos haciendo un DRAG de un determinado Nodo
'##########################################################################################
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    IndiceNodoDrag = BuscaNodoMasProximo(CLng(X), CLng(Y))

    lblStatusNodoDrag.Caption = "Nodo:" & TreeNodeList(IndiceNodoDrag).CurrNode
    
    If TreeNodeList(IndiceNodoDrag).CurrNode = CurrSrcNode Then
        lblStatusNodoDrag.Caption = lblStatusNodoDrag.Caption & " - Nodo Origen"
    ElseIf TreeNodeList(IndiceNodoDrag).CurrNode = CurrDestNode Then
        lblStatusNodoDrag.Caption = lblStatusNodoDrag.Caption & " - Nodo Destino"
    End If

End Sub

'##########################################################################################
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IndiceNodoDrag <> -1 And FicheroCargado = True Then
        Nodo(IndiceNodoDrag).X = X
        Nodo(IndiceNodoDrag).Y = Y
    End If
End Sub

'##########################################################################################
'Hemos hecho el DROP del Nodo que habiamos DRAGeado anteriormente
'##########################################################################################
Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndiceNodoDrag = -1
    lblStatusNodoDrag.Caption = ""
End Sub

'##########################################################################################
Private Sub tmrRender_Timer()
    RenderNodos
    'DoEvents
End Sub
