Attribute VB_Name = "mHojas"
' ------------------------------------------------------
' Name: mHojas
' Kind: Módulo
' Purpose: Extender las opciones para gestionar las hojas del libro
' Author: RicardoDiaz
' Date: 9/09/2019
' ------------------------------------------------------
Option Explicit
Option Private Module

' ----------------------------------------------------------------
' Procedure Name: AgregarReferenciarHoja
' Purpose: Agregar una hoja al libro y si ya existe hacer referencia a ella
' Procedure Kind: Function
' Procedure Access: Public
' Parameter NombreHoja (String): Nombre de la hoja a insertar o hacer referencia
' Parameter EliminarExistente (Boolean): Eliminar la hoja si existe para quitar todo el contenido
' Parameter BorrarContenido (Boolean): Borrar el contenido de la hoja existente si se quiere empezar en blanco
' Return Type: Worksheet
' Author: RicardoDiaz
' Date: 9/09/2019
' ----------------------------------------------------------------
Public Function AgregarReferenciarHoja(NombreHoja As String, _
    Optional EliminarExistente As Boolean = False, _
    Optional BorrarContenido As Boolean = False) As Worksheet
    
    
    
    ' Declarar objetos
    Dim hoja As Worksheet
    
    ' Declarar variables
    Dim hojaExiste As Boolean
    Dim agregarHoja As Boolean
    
    ' Inicializar variables generales
    hojaExiste = (Evaluate("ISREF('" & NombreHoja & "'!A1)") = True)
    
    
    If (hojaExiste = True) And (EliminarExistente = True) Then
        ThisWorkbook.Sheets(NombreHoja).Delete
        agregarHoja = True
    ElseIf (hojaExiste = False) Then
        agregarHoja = True
    ElseIf (hojaExiste = True) And (EliminarExistente = False) Then
        agregarHoja = False
    End If
    
    If agregarHoja = True Then
        Set hoja = ThisWorkbook.Sheets.Add
        hoja.Name = NombreHoja
    Else
        Set hoja = ThisWorkbook.Worksheets(NombreHoja)
    End If

    If BorrarContenido = True Then hoja.UsedRange.Clear
    
    Set AgregarReferenciarHoja = hoja

End Function



' ----------------------------------------------------------------
' Procedure Name: EncontrarHojasPorCriterio
' Purpose: Encontrar hojas que cumplen criterio de caracteres en el nombre (devuelve colección)
' Procedure Kind: Function
' Procedure Access: Public
' Parameter IncluyeExcluyeCaracteresNombre (enIncluyeExcluye): Incluir o excluir caracteres
' Parameter CaracteresNombre (String): Caracteres que debe contener el nombre de la hoja
' Return Type: Collection
' Author: RicardoDiaz
' Date: 10/09/2019
' ----------------------------------------------------------------
Public Function EncontrarHojasPorCriterio(IncluyeExcluyeCaracteresNombre As enIncluyeExcluye, _
    CaracteresNombre As String) As Collection

    ' Declarar objetos
    Dim hoja As Worksheet
    Dim colHojas As New Collection
    
    ' Declarar variables
    Dim incluirHojaPorNombre As Boolean
    
    ' Inicializar variables generales
    incluirHojaPorNombre = True
        
    ' Buscar estilos a eliminar
    For Each hoja In ThisWorkbook.Sheets

        ' Incluir si contiene caracteres en el nombre
        incluirHojaPorNombre = ((CaracteresNombre = "") Or _
            (IncluyeExcluyeCaracteresNombre = evIncluye And InStr(LCase(hoja.Name), CaracteresNombre) > 0) Or _
            (IncluyeExcluyeCaracteresNombre = evExcluye And InStr(LCase(hoja.Name), CaracteresNombre) = 0))
                
        ' Agregar a colección de estilos a eliminar
        If incluirHojaPorNombre = True Then
                
            colHojas.Add hoja
                
        End If
    
    Next hoja
    
    Set EncontrarHojasPorCriterio = colHojas
    
End Function

