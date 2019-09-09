Option Explicit
Option Private Module
' Creditos: https://github.com/spences10


'// Mostrar un mensaje de error y devolver verdadero si el usuario lo indica
Public Function ManejarError(ByVal ErrNumber As Long, ByVal ErrDesc As String, ByVal ErrSource As String) As Boolean

    Dim UserAction As Integer

    UserAction = MsgBox( _
        Prompt:= _
            "Se produjo un error inesperado. Por favor reportarlo a " & _
            "https://github.com/rdiazjimenez/UtilidadesVBA/issues" & vbNewLine & vbNewLine & _
            "Número de error: " & ErrNumber & vbNewLine & _
            "Descripción: " & ErrDesc & vbNewLine & _
            "Fuente: " & ErrSource & vbNewLine & vbNewLine & _
            "Quiere visualizar el código?", _
        Buttons:=vbYesNo + vbDefaultButton2, _
        Title:="Error inesperado")

    HandleCrash = UserAction = vbYes

End Function