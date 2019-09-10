Attribute VB_Name = "mPlantillaModulo"
' ------------------------------------------------------
' Name: mPlantillaModulo
' Kind: M�dulo
' Purpose: Plantillas con estructuras de procedimientos y funciones para agilizar producci�n de c�digo VBA est�ndar
' Author: RicardoDiaz
' Date: 9/09/2019
' ------------------------------------------------------
Option Explicit
Option Private Module

' ----------------------------------------------------------------
' Procedure Name: NombreProcedimiento
' Purpose: Escribir prop�sito del procedimiento
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

    ' Inicio c�digo

    
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
