' ----------------------------------------------------------------
' Nombre del procedimiento: EscribirNombre
' Objetivo: EscribirObjetivo
' Tipo: Sub
' Acceso: Public
' Autor: RicardoDiaz
' Fecha: 9/09/2019
' ----------------------------------------------------------------
Public Sub NombreProcedimiento()

    ' Declarar objetos
    Dim oSt As Style
    Dim oCell As Range

    ' Declarar variables
    Dim lCount As Long
    Dim CurStyle As Style
    
    ' Iniciar control de errores
    On Error GoTo manejarError
    

    ' 
    For Each CurStyle In ThisWorkbook.Styles
        
        lCount = lCount + 1
        
        'oStylesh.Cells(lCount, 1).Value = CurStyle.Name
        'oStylesh.Cells(lCount, 2).Value = CurStyle.NameLocal
                
        ' Borrar estilo
        If InStr(CurStyle, "% -") > 0 And CurStyle.BuiltIn = False Then CurStyle.Delete
        
    Next CurStyle
    
salirSub:
    Exit Sub

manejarError:
    If ManejarError(Err.Number, Err.Description, Err.Source) Then
        Stop
        Resume
    End If
    GoTo salirSub
    
End Sub