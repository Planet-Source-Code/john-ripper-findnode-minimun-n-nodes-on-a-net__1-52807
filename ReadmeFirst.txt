Structure of data File:

    Nº nodes
    Start Node
    End Node
    Interconections Number
    Interconections between Node i, Node j
    ...
    Interconections between n-1 Node i, Node j
    Interconections between n Node i, Node j

Example:
	
    4
    1
    4
    5
    1,2
    1,3
    2,3
    1,4
    2,4


	1     2
	*---*
        |\ /\
	| /  |
	|/ \  \
       3*   \  |
	     \ |
	      \*4

(As you can see, i'm not an ASCII Artist! :-)
    

Default Data file its located on AppPath & "datos.txt"

You can execute the portram with de -RENDER param to display best traceroute.


With -RENDER param you can drag & drop nodes with de mouse where ever you want.

if you execute the program with -RENDER param, you can choose another data file sample on folder \MoreTestData

The result of the shortest path its saved on AppPath & "\resultados.txt" and the format its:
	node 1, node 2, node 3, ...., node n-1, node n


	