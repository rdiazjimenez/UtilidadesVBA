Attribute VB_Name = "mPlantillaModulo"
' ------------------------------------------------------
' Name: mPlantillaModulo
' Kind: Módulo
' Purpose: Plantillas con estructuras de procedimientos y funciones para agilizar producción de código VBA estándar
' Author: RicardoDiaz
' Date: 9/09/2019
' ------------------------------------------------------
Option Explicit
Option Private Module

' ----------------------------------------------------------------
' Procedure Name: NombreProcedimiento
' Purpose: Escribir propósito del procedimiento
' Procedure Kind: Sub
' Procedure Access: Public
' Author: RicardoDiaz
' Date: 9/09/2019
' ----------------------------------------------------------------
Public Sub NombreProcedimiento()

    ' Declarar objetos
    
    ' Declarar variables
    
    ' Iniciar control de errores
    On Error GoTo ManejarError
    
    ' Apagar todo
    Call ApagarTodo

    ' Inicializar variables generales

    ' Inicializar objetos

    ' Iniciarlizar otras variables

    ' Inicio código

    
SalirProcedimiento:
    Call PrenderTodo
    Exit Sub

ManejarError:
    If ManejarError(Err.Number, Err.Description, Err.Source) Then
        Stop
        Resume
    End If
    GoTo SalirProcedimiento
    
End Sub
