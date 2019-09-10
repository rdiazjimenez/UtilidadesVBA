Attribute VB_Name = "mPruebas"
' ------------------------------------------------------
' Name: mPruebas
' Kind: Módulo
' Purpose: Módulo de pruebas (descartable)
' Author: RicardoDiaz
' Date: 9/09/2019
' ------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: ProbarProcedimientos
' Purpose: Llamadas de procedimientos de prueba
' Procedure Kind: Sub
' Procedure Access: Public
' Author: RicardoDiaz
' Date: 10/09/2019
' ----------------------------------------------------------------
Sub ProbarProcedimientos()
    
    Call mEstilos.Reemplazar(BuscarEnLibro:=False)
    'Call mEstilos.BuscarEstilos(TiposEstilos:=evTodosEstilos, IncluyeExcluyeCaracteresNombre:=evIncluye, CaracteresNombre:="es", BuscarEnLibro:=True)
    'Call mEstilos.Eliminar(TiposEstilos:=evTodosEstilos, IncluyeExcluyeCaracteresNombre:=evExcluye, CaracteresNombre:="bu", MostrarMensaje:=evMostrar)
    
End Sub
