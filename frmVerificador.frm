VERSION 5.00
Begin VB.Form frmVerificador 
   BackColor       =   &H00800000&
   Caption         =   "VERIFICADOR SUDOKUS 4X4"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   " RESULTADOS "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   3840
      TabIndex        =   22
      Top             =   240
      Width           =   7095
      Begin VB.ListBox lstSoluciones 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3930
         Left            =   3600
         TabIndex        =   26
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CommandButton cmdImprimirResultados 
         BackColor       =   &H00FF8080&
         Caption         =   "IMPRIMIR RESULTADOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5400
         Width           =   6615
      End
      Begin VB.Label lblCondicion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "SOLUCIONES ENCONTRADAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3600
         TabIndex        =   25
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "CONDICIÓN DEL PLANTEAMIENTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   " PANEL DEL PROBLEMA "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3495
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   21
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   16
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   20
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   15
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   19
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   14
         Left            =   960
         MaxLength       =   1
         TabIndex        =   18
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   240
         MaxLength       =   1
         TabIndex        =   17
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   15
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   960
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   240
         MaxLength       =   1
         TabIndex        =   13
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   960
         MaxLength       =   1
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   240
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   960
         MaxLength       =   1
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   240
         MaxLength       =   1
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdVerificar 
         BackColor       =   &H00FF8080&
         Caption         =   "VERIFICAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3480
         Width           =   3015
      End
   End
   Begin VB.Frame framLineaProblemas 
      BackColor       =   &H00FFC0C0&
      Caption         =   " LÍNEA PARA CARGAR PROBLEMAS   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   4680
      Width           =   3495
      Begin VB.CommandButton cmdExtraerProblema 
         BackColor       =   &H00FF8080&
         Caption         =   "EXTRAER PROBLEMA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdCargarProblema 
         BackColor       =   &H00FF8080&
         Caption         =   "CARGAR PROBLEMA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtCargaProblema 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         TabIndex        =   1
         Text            =   "1234341241232341"
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmVerificador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : VERIFICADOR DEL SUDOKU 4X4
'* CONTENIDO     : VERIFICA LAS SOLUCIONES DEL SUDOKUS DE 4X4
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 17 DE ABRIL DE 2014
'* ACTUALIZACION : 17 DE ABRIL DE 2014
'****************************************************************************************
Option Explicit

' DECLARACION DE TIPOS DE VARIABLES
' LOS POSIBLES PROBLEMAS QUE SE PUEDEN PLANTEAR
Private Type misProblemas
    miNumero As Long
    miCasilla(1 To 16) As Integer
    miPosible As Boolean
End Type

' LAS 288 SOLUCIONES QUE EXISTEN PARA EL SUDOKU DE 4X4
Private Type misSoluciones
    miNumero As Long
    miCasilla(1 To 16) As Integer
End Type

' VECTOR QUE SERÁ ANALIZADO POR EL PROGRAMA
Private Type misAnalizados
    miNumero As Long
    miCasilla(1 To 16) As Integer
End Type

' DECLARACION DE VARIABLES
Dim miProblema(1 To 1000000)  As misProblemas
Dim miSolucion(1 To 288) As misSoluciones
Dim miAnalizado As misAnalizados
Dim miVectorAnalisis(1 To 16) As Integer
Dim miCantidadProblemas As Long
Dim miCantidadProblemasEncontrados As Long
Dim miLineOutput As String
Dim miContadorCeros As Integer
Dim miContadorTotalEstudiados As Long
Dim miContadorTotalEncontrados As Long

Dim miIdPlanteados As Long
Dim miPistas As Integer
Dim miIncognitas As Integer

Dim miCuentaCasosEstadisticos As Long
Dim miCuentaBuenosEstadisticos As Long
Dim miCuentaMalosEstadisticos As Long
Dim miCuentaAmbiguosEstadisticos As Long

' IMPRIME LOS RESULTADOS OBTENIDOS
Private Sub cmdImprimirResultados_Click()
    Dim i As Integer
    
    Open "RESULTADOS.txt" For Output As #7
    ' EL PROBLEMA PANTEADO
    Print #7, ""
    Print #7, " El Problema Planteado es : "
    Print #7, ""
    
    miLineOutput = "   "
    For i = 1 To 16
        If Val(txtProblema(i)) = 0 Then
            miLineOutput = miLineOutput + "0"
        Else
            miLineOutput = miLineOutput + Trim(txtProblema(i))
        End If
        If i = 4 Or i = 8 Or i = 12 Then
            miLineOutput = miLineOutput + "."
        End If
    Next i
    Print #7, miLineOutput
    Print #7, ""
    
    ' LA CONDICION OBTENIDA
    Print #7, ""
    Print #7, " La condición obtenida del planteamiento es : "
    Print #7, ""
    Print #7, "   " + lblCondicion.Caption
    Print #7, ""
    
    ' LAS SOLUCIONES
    Print #7, ""
    Print #7, " Las soluciones son: "
    Print #7, ""
    For i = 0 To (lstSoluciones.ListCount - 1)
        lstSoluciones.ListIndex = i
        Print #7, "   " + lstSoluciones.Text
    Next i
    Print #7, ""
    
    Close #7
End Sub

' VERIFICAR EL PROBLEMA PLANTEADO
Private Sub cmdVerificar_Click()
    lblCondicion = ""
    lstSoluciones.Clear
    Call miCargaVector
    Call miAnalizaVector
End Sub

' CARGA LOS DATOS EN EL VECTOR DE ANALISIS
Private Sub miCargaVector()
    Dim i As Integer
    ' CARGA LOS DATOS PARA ANALIZARLOS
    For i = 1 To 16
        miVectorAnalisis(i) = Val(txtProblema(i).Text)
    Next i
End Sub

' ANALIZA EL VECTOR PARA VER SI ES UN PROBLEMA PLANTEABLE O NO
Private Sub miAnalizaVector()
    Dim i As Integer
    Dim j As Integer
    Dim miSirve As Boolean
    Dim miContadorIguales As Integer
     
    ' IMPRIME DATOS EN EL ARCHIVO DE TEXTO
    miLineOutput = ""
    For i = 1 To 16
        miLineOutput = miLineOutput + Str(miVectorAnalisis(i))
    Next i
    'Print #10, miLineOutput
    
    ' REVISA CONTRA TODAS LAS SOLUCIONES
    miContadorIguales = 0
    For j = 1 To 288
        miSirve = True
        For i = 1 To 16
            If miVectorAnalisis(i) <> 0 Then
                If miVectorAnalisis(i) <> miSolucion(j).miCasilla(i) Then
                    miSirve = False
                End If
            End If
        Next i
        If miSirve = True Then
            ' CUENTA LA SOLUCION
            miContadorIguales = miContadorIguales + 1
        
            ' MUESTRA SOLUCIONES ENCONTRADAS EN EL LISTBOX
            miLineOutput = ""
            For i = 1 To 16
                miLineOutput = miLineOutput + Trim(Str(miSolucion(j).miCasilla(i)))
                If i = 4 Or i = 8 Or i = 12 Then
                    miLineOutput = miLineOutput + "."
                End If
                
            Next i
            lstSoluciones.AddItem miLineOutput
        End If
    Next j
    
    ' CUENTA EL CASO ESTUDIADO COMO UN PROBLEMA INCORRECTO
    If miContadorIguales = 0 Then
        miCuentaMalosEstadisticos = miCuentaMalosEstadisticos + 1
    End If
    
    ' CUENTA LA CANTIDAD DE CEROS EN EL VECTOR
    miContadorCeros = 0
    For i = 1 To 16
        If miVectorAnalisis(i) = 0 Then
            miContadorCeros = miContadorCeros + 1
        End If
    Next i
    
    miCantidadProblemas = miCantidadProblemas + 1
    
    ' CARGA LA BASE DE DATOS DE LOS PROBLEMAS PLANTEABLES
    miIncognitas = miContadorCeros
    miPistas = 16 - miContadorCeros
    miIdPlanteados = miIdPlanteados + 1
    
    ' DETERMINA SI EL PLANTAMIENTO ES AMBIGUO O NO
    ' NO ES AMBIGUO
    If miContadorIguales = 1 Then
        miProblema(miCantidadProblemas).miNumero = miCantidadProblemas
        For i = 1 To 16
            miProblema(miCantidadProblemas).miCasilla(i) = miVectorAnalisis(i)
        Next i
        miProblema(miCantidadProblemas).miPosible = True
        
        ' MUESTRA RESULTADOS EN EL FORMULARIO
        miCantidadProblemasEncontrados = miCantidadProblemasEncontrados + 1
        'lblEncontrados = miCantidadProblemasEncontrados
        'DoEvents
    
        ' CUENTA EL CASO ESTUDIADO COMO UN PROBLEMA BIEN PLANTEADO
        miCuentaBuenosEstadisticos = miCuentaBuenosEstadisticos + 1
    End If
    
    ' ES AMBIGUO
    If miContadorIguales > 1 Then
        miProblema(miCantidadProblemas).miNumero = miCantidadProblemas
        For i = 1 To 16
            miProblema(miCantidadProblemas).miCasilla(i) = miVectorAnalisis(i)
        Next i
        miProblema(miCantidadProblemas).miPosible = True
        
        ' CUENTA EL CASO ESTUDIADO COMO UN PROBLEMA AMBIGUO
        miCuentaAmbiguosEstadisticos = miCuentaAmbiguosEstadisticos + 1
    End If
    
    ' EXPONE LAS RESPUESTAS ENCONTRADAS
    If miContadorIguales = 0 Then
        lblCondicion = "El problema planteado no tiene solución"
    End If
    If miContadorIguales = 1 Then
        lblCondicion = "El problema planteado tiene solución única"
    End If
    If miContadorIguales > 1 Then
        lblCondicion = "El problema planteado tiene " + Trim(Str(miContadorIguales)) + " soluciones"
    End If
    DoEvents
End Sub

' AL MOMENTO DE CARGAR EL FORMULARIO INICIAL
Private Sub Form_Load()
    ' CARGA VALORES INICIALES PARA LAS SOLUCIONES
    ' ABRE EL ARCHIVO CON LAS 288 SOLUCIONES
    Dim X As Integer
    Dim miLineInput As String
    Dim miNumero As Integer
    Open "miSolucionesTotal.txt" For Input As #10
    Do Until EOF(10)
        Line Input #10, miLineInput
        miNumero = Val(Mid(miLineInput, 32, 3))
        ' CARGA EL NUMERO DE LA GRILLA
        miSolucion(miNumero).miNumero = Val(Mid(miLineInput, 32, 3))
        ' CARGA LOS VALORES DE LOS DÍGITOS QUE COMPONEN LA SOLUCION
        For X = 1 To 16
            miSolucion(miNumero).miCasilla(X) = Val(Mid(miLineInput, X, 1))
        Next X
    Loop
    Close #10
End Sub

' CARGAR PROBLEMA DESDE LA LINEA
Private Sub cmdCargarProblema_Click()
    CargarProblema (txtCargaProblema)
End Sub

' EXTRAER PROBLEMA HACIA LA LINEA
Private Sub cmdExtraerProblema_Click()
    txtCargaProblema = ExtraerProblema()
End Sub

' CARGA PROBLEMA QUE SE PRESENTAN EN FORMA DE LINEA
Private Function CargarProblema(miEnviado As String)
    Dim X As Integer
    Call Limpia
    For X = 1 To 16
        If Val(Mid(miEnviado, X, 1)) = 0 Then
            txtProblema(X) = ""
        Else
            txtProblema(X) = Mid(miEnviado, X, 1)
        End If
    Next X
End Function

' EXTRAER PROBLEMA HACIA UNA VARIABLE DE TIPO CARACTER
Private Function ExtraerProblema() As String
    Dim X As Integer
    Dim ProblemaExtraido As String
    ProblemaExtraido = ""
    For X = 1 To 16
        If txtProblema(X) = "" Then
            ProblemaExtraido = ProblemaExtraido & "0"
        Else
            ProblemaExtraido = ProblemaExtraido & txtProblema(X)
        End If
    Next X
    ExtraerProblema = ProblemaExtraido
End Function

' LIMPIA LOS VALORES VISIBLES EN TODO EL FORMULARIO
Private Sub Limpia()
    Dim X As Integer
    Dim i As Integer
    For X = 1 To 16
        txtProblema(X) = ""
        txtProblema(X).Enabled = True
    Next X
    lblCondicion = ""
    lstSoluciones.Clear
End Sub

