Attribute VB_Name = "mGestionErrores"
' ------------------------------------------------------
' Name: mGestionErrores
' Kind: M�dulo
' Purpose: Gestionar los errores generados en el archivo
' Author: RicardoDiaz
' Date: 9/09/2019
' ------------------------------------------------------
Option Explicit
Option Private Module
' Creditos: https://github.com/spences10



' ----------------------------------------------------------------
' Procedure Name: ManejarError
' Purpose: Mostrar un mensaje de error y devolver verdadero si el usuario lo indica
' Procedure Kind: Function
' Procedure Access: Public
' Parameter ErrNumber (Long): N�mero del error generado
' Parameter ErrDesc (String): Descripci�n del error generado
' Parameter ErrSource (String): Fuente del error generado
' Return Type: Boolean
' Author: RicardoDiaz
' Date: 9/09/2019
' ----------------------------------------------------------------
Public Function ManejarError(ByVal ErrNumber As Long, ByVal ErrDesc As String, ByVal ErrSource As String) As Boolean

    Dim accionUsuario As Integer

    accionUsuario = MsgBox( _
        Prompt:= _
        "Se produjo un error inesperado. Por favor reportarlo a " & _
        "https://github.com/rdiazjimenez/UtilidadesVBA/issues" & vbNewLine & vbNewLine & _
        "N�mero de error: " & ErrNumber & vbNewLine & _
        "Descripci�n: " & ErrDesc & vbNewLine & _
        "Fuente: " & ErrSource & vbNewLine & vbNewLine & _
        "Quiere visualizar el c�digo?", _
        Buttons:=vbYesNo + vbDefaultButton2, _
        Title:="Error inesperado")

    ManejarError = accionUsuario = vbYes

End Function

