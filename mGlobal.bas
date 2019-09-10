Attribute VB_Name = "mGlobal"
' ------------------------------------------------------
' Name: mGlobal
' Kind: Módulo
' Purpose: Variables, Constantes, Enumeraciones y Procedimientos generales
' Author: RicardoDiaz
' Date: 9/09/2019
' ------------------------------------------------------
Option Explicit
Option Private Module

' Variables globales
Global EstadoProteccionLibro As Boolean
Global EstadoCalculo As XlCalculation
Global EstadoEventos As Boolean
Global EstadoAlertas As Boolean
Global EstadoActualizaPantalla As Boolean
Global EstadoBarraEstados As Boolean
Global EstadoSaltoPagina As Boolean

' Constantes


' Enumeraciones
Public Enum enTipoEstilo
    evTodosEstilos = 1
    evEstiloNativo = 2
    evEstiloNoNativo = 3
End Enum

Public Enum enIncluyeExcluye
    evIncluye = 1
    evExcluye = 2
End Enum

Public Enum enMostrarOcultar
    evMostrar = 1
    evNoMostrar = 2
    evOcultar = 3
End Enum

' ----------------------------------------------------------------
' Procedure Name: ApagarTodo
' Purpose: Desactivar propiedades de la aplicación y el archivo para aumentar velocidad de ejecución del código
' Procedure Kind: Sub
' Procedure Access: Public
' Author: RicardoDiaz
' Date: 9/09/2019
' ----------------------------------------------------------------
Public Sub ApagarTodo()

    Call GuardarEstados
    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
        .ScreenUpdating = False
        .DisplayStatusBar = False
    End With
    
    ActiveSheet.DisplayPageBreaks = False
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: GuardarEstados
' Purpose: Guardar propiedades de la aplicación y el archivo antes de ejecutar código VBA
' Procedure Kind: Sub
' Procedure Access: Public
' Author: RicardoDiaz
' Date: 9/09/2019
' ----------------------------------------------------------------
Public Sub GuardarEstados()

    With Application
        EstadoCalculo = .Calculation
        EstadoEventos = .EnableEvents
        EstadoAlertas = .DisplayAlerts
        EstadoActualizaPantalla = .ScreenUpdating
        EstadoBarraEstados = .DisplayStatusBar

    End With
    
    EstadoSaltoPagina = ActiveSheet.DisplayPageBreaks
        
End Sub

' ----------------------------------------------------------------
' Procedure Name: PrenderTodo
' Purpose: Retornar propiedades de la aplicación y el archivo al estado de antes de ejecutar el código VBA
' Procedure Kind: Sub
' Procedure Access: Public
' Author: RicardoDiaz
' Date: 9/09/2019
' ----------------------------------------------------------------
Public Sub PrenderTodo()
    With Application
        .Calculation = EstadoCalculo
        .EnableEvents = EstadoEventos
        .DisplayAlerts = EstadoAlertas
        .ScreenUpdating = EstadoActualizaPantalla
        .DisplayStatusBar = EstadoBarraEstados
    End With
    
    ActiveSheet.DisplayPageBreaks = EstadoSaltoPagina
End Sub
