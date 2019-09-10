Attribute VB_Name = "mEstilos"
' ------------------------------------------------------
' Name: mEstilos
' Kind: Módulo
' Purpose: Extender las opciones para gestionar los estilos del libro
' Author: RicardoDiaz
' Date: 9/09/2019
' ------------------------------------------------------
Option Explicit
Option Private Module

' ----------------------------------------------------------------
' Procedure Name: BuscarCeldasConEstilos
' Purpose: Buscar celdas en hojas determinadas con estilos específicos
' Procedure Kind: Function
' Procedure Access: Public
' Parameter ColEstilos (Collection): Colección de estilos a buscar
' Parameter colHojas (Collection): Colección de hojas donde buscar los estilos
' Return Type: Collection
' Author: RicardoDiaz
' Date: 10/09/2019
' ----------------------------------------------------------------
Public Function BuscarCeldasConEstilos(ColEstilos As Collection, colHojas As Collection) As Collection

    ' Declarar objetos
    Dim colResultado As Collection
    Dim hoja As Worksheet
    Dim celda As Range
    Dim estilo As Style
    
    ' Declarar variables


    ' Inicializar variables generales

    ' Inicializar objetos
    Set colResultado = New Collection
    
    ' Iniciarlizar otras variables

    ' Inicio código
    For Each estilo In ColEstilos

        For Each hoja In colHojas
            
            For Each celda In hoja.UsedRange.Cells
            
                If celda.Style = estilo Then
                    colResultado.Add celda
                End If
            
            Next celda
    
        Next hoja

    Next estilo
    
    Set BuscarCeldasConEstilos = colResultado
    
End Function

' ----------------------------------------------------------------
' Procedure Name: BuscarEstilos
' Purpose: Buscar estilos en un libro y generar un reporte
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter TiposEstilos (enTipoEstilo): Tipos de estilos a eliminar
' Parameter IncluyeExcluyeCaracteresNombre (enIncluyeExcluye): Incluir o excluir caracteres
' Parameter CaracteresNombre (String): Caracteres que debe contener el nombre del estilo
' Parameter BuscarEnLibro (Boolean): Buscar en todo el libro (true) o solo la hoja activa (false)
' Author: RicardoDiaz
' Date: 10/09/2019
' ----------------------------------------------------------------
Public Sub BuscarEstilos(Optional TiposEstilos As enTipoEstilo = evEstiloNoNativo, _
    Optional IncluyeExcluyeCaracteresNombre As enIncluyeExcluye = evIncluye, _
    Optional CaracteresNombre As String = vbNullString, Optional BuscarEnLibro As Boolean = True)
    
    ' Declarar objetos
    Dim hojaResultados As Worksheet
    Dim tablaResultados As ListObject
    
    Dim ColEstilos As Collection
    Dim colHojas As Collection
    Dim colCeldas As Collection
        
    ' Declarar variables
    Dim nombreHojaResultados As String
    Dim nombreTablaResultados As String
    Dim listadoColumnasResultados As Variant
    Dim mensajeSinCoincidencias As String


    ' Iniciar control de errores
    On Error GoTo ManejarError
    
    ' Apagar todo
    Call ApagarTodo
    
    ' Inicializar variables generales
    nombreHojaResultados = "AnalisisEstilos"
    nombreTablaResultados = "TablaAnalisisEstilos"
    listadoColumnasResultados = Array("Hoja", "Celda", "Estilo")
    mensajeSinCoincidencias = "No se encontraron celdas con estilos que cumplan los criterios definidos"
    
    ' Encontrar estilos que cumplen criterios de tipo de estilo y caracteres en el nombre (devuelve colección)
    Set ColEstilos = EncontrarEstilosPorCriterio(TiposEstilos, IncluyeExcluyeCaracteresNombre, CaracteresNombre)
    
    If BuscarEnLibro = True Then
        ' Encontrar hojas que cumplen criterio de caracteres en el nombre (devuelve colección)
        Set colHojas = mHojas.EncontrarHojasPorCriterio(evExcluye, nombreHojaResultados)
    Else
        Set colHojas = New Collection
        colHojas.Add ActiveSheet
    End If
    
    ' Buscar celdas con estilos que cumplen criterios
    Set colCeldas = BuscarCeldasConEstilos(ColEstilos, colHojas)
    
    ' Agregar hoja de resultados
    Set hojaResultados = mHojas.AgregarReferenciarHoja(nombreHojaResultados, True, True)

    If colCeldas.Count = 0 Then
        hojaResultados.Range("A1").Value2 = mensajeSinCoincidencias
        Exit Sub
    End If
    
    ' Preparar hoja de resultados
    Set tablaResultados = PrepararTablaEstructurada(hojaResultados, nombreTablaResultados, listadoColumnasResultados)
    
    ' Registrar información en hoja de resultados
    Call RegistrarInfoEstiloCelda(colCeldas, tablaResultados)
    
    ' Mostrar resultado a usuario
    hojaResultados.Activate

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

' ----------------------------------------------------------------
' Procedure Name: Eliminar
' Purpose: Eliminar estilos del libro que cumplen criterios
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter TiposEstilos (enTipoEstilo): Tipos de estilos a eliminar
' Parameter IncluyeExcluyeCaracteresNombre (enIncluyeExcluye): Incluir o excluir caracteres
' Parameter CaracteresNombre (String): Caracteres que debe contener el nombre del estilo
' Parameter MostrarMensaje (enMostrarOcultar): Mostrar mensajes de confirmación y resultados al usuario
' Author: RicardoDiaz
' Date: 9/09/2019
' ----------------------------------------------------------------
Public Sub Eliminar(Optional TiposEstilos As enTipoEstilo = evEstiloNoNativo, _
    Optional IncluyeExcluyeCaracteresNombre As enIncluyeExcluye = evIncluye, _
    Optional CaracteresNombre As String = vbNullString, _
    Optional MostrarMensaje As enMostrarOcultar = evMostrar)

    ' Declarar objetos
    Dim ColEstilos As Collection
    
    ' Declarar variables
    Dim cancelar As Boolean
    Dim contadorEliminados As Single

    ' Iniciar control de errores
    On Error GoTo ManejarError
    
    ' Apagar todo
    Call ApagarTodo
    
    ' Encontrar estilos que cumplen criterios de tipo de estilo y caracteres en el nombre (devuelve colección)
    Set ColEstilos = EncontrarEstilosPorCriterio(TiposEstilos, IncluyeExcluyeCaracteresNombre, CaracteresNombre)
    
    ' Eliminar estilos encontrados
    contadorEliminados = EliminarEstilosEnColeccion(ColEstilos, MostrarMensaje, cancelar)
    
    ' Mostrar resultados al usuario
    If MostrarMensaje = evMostrar And cancelar <> True Then MostrarResultadoEliminarEstilos (contadorEliminados)

    
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

' ----------------------------------------------------------------
' Procedure Name: EliminarEstilosEnColeccion
' Purpose: Eliminar estilos que cumplen criterio. Devuelve la cantidad de estilos eliminados
' Procedure Kind: Function
' Procedure Access: Private
' Parameter ColEstilos (Collection): Colección de estilos que cumplen los criterios para ser eliminados
' Parameter MostrarMensaje (enMostrarOcultar): Mostrar o no la confirmación de eliminación de estilos
' Parameter cancelar (Boolean): Cancelar la eliminación si el usuario lo decide
' Return Type: Single
' Author: RicardoDiaz
' Date: 10/09/2019
' ----------------------------------------------------------------
Private Function EliminarEstilosEnColeccion(ColEstilos As Collection, _
    MostrarMensaje As enMostrarOcultar, _
    cancelar As Boolean) As Single
    
    ' Declarar objetos
    Dim estilo As Variant
        
    ' Declarar variables
    Dim tituloConfirmar As String
    Dim mensajeConfirmar As String
    Dim mensajeusuario As String
    
    Dim continuarEliminar As VbMsgBoxResult
    
    Dim contadorEncontrados As Single
    Dim contadorEliminados As Single
        
    ' Inicializar variables generales
    tituloConfirmar = "Confirmar operación"
    mensajeConfirmar = "Estilos encontrados: (<contador>) para eliminar, está seguro que quiere continuar?"
    
    contadorEncontrados = ColEstilos.Count

    ' Escoger mensaje dependiendo de si hay estilos para eliminar
    If MostrarMensaje <> evMostrar Then
        continuarEliminar = True
    ElseIf MostrarMensaje = evMostrar And contadorEncontrados > 0 Then
        mensajeusuario = Replace(mensajeConfirmar, "<contador>", contadorEncontrados)

        ' Confirmar con usuario si eliminar
        cancelar = (MsgBox(mensajeusuario, vbYesNo, tituloConfirmar) <> vbYes)
    End If
    
    ' Eliminar estilos encontrados
    If (MostrarMensaje = evMostrar And cancelar = False) Or MostrarMensaje = evNoMostrar Then
    
        For Each estilo In ColEstilos
            
            estilo.Delete
                
            contadorEliminados = contadorEliminados + 1
            
        Next estilo
            
    End If
    
    EliminarEstilosEnColeccion = contadorEliminados
    
End Function

' ----------------------------------------------------------------
' Procedure Name: EncontrarEstilosPorCriterio
' Purpose: Encontrar estilos que cumplen criterios de tipo de estilo y caracteres en el nombre (devuelve colección)
' Procedure Kind: Function
' Procedure Access: Private
' Parameter TiposEstilos (enTipoEstilo): Tipos de estilos a eliminar
' Parameter IncluyeExcluyeCaracteresNombre (enIncluyeExcluye): Incluir o excluir caracteres
' Parameter CaracteresNombre (String): Caracteres que debe contener el nombre del estilo
' Return Type: Collection
' Author: RicardoDiaz
' Date: 9/09/2019
' ----------------------------------------------------------------
Private Function EncontrarEstilosPorCriterio(TiposEstilos As enTipoEstilo, _
    IncluyeExcluyeCaracteresNombre As enIncluyeExcluye, _
    CaracteresNombre As String) As Collection

    ' Declarar objetos
    Dim estilo As Style
    Dim ColEstilos As New Collection
    
    ' Declarar variables
    Dim incluirEstiloPorTipo As Boolean
    Dim incluirEstiloPorNombre As Boolean
    
    ' Inicializar variables generales
    incluirEstiloPorNombre = True
        
    ' Buscar estilos a eliminar
    For Each estilo In ThisWorkbook.Styles
        
        ' Incluir por tipo de estilo
        incluirEstiloPorTipo = (TiposEstilos = evTodosEstilos) Or _
            (TiposEstilos = evEstiloNativo And estilo.BuiltIn = True) Or _
            (TiposEstilos = evEstiloNoNativo And estilo.BuiltIn = False)
        
        ' Incluir si contiene caracteres en el nombre
        incluirEstiloPorNombre = ((CaracteresNombre = "") Or _
            (IncluyeExcluyeCaracteresNombre = evIncluye And InStr(LCase(estilo.Name), CaracteresNombre) > 0) Or _
            (IncluyeExcluyeCaracteresNombre = evExcluye And InStr(LCase(estilo.Name), CaracteresNombre) = 0)) _
            And estilo.Name <> "Normal"
                
        ' Agregar a colección de estilos a eliminar
        If incluirEstiloPorTipo = True And incluirEstiloPorNombre = True Then
                
            ColEstilos.Add estilo
                
        End If
    
    Next estilo
    
    Set EncontrarEstilosPorCriterio = ColEstilos
    
End Function

' ----------------------------------------------------------------
' Procedure Name: MostrarResultadoEliminarEstilos
' Purpose: Mostrar resultados de la eliminación al usuario
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter ContadorEstilosEliminados (Single): Contador de estilos eliminados
' Author: RicardoDiaz
' Date: 9/09/2019
' ----------------------------------------------------------------
Private Sub MostrarResultadoEliminarEstilos(ContadorEstilosEliminados As Single)

    ' Declarar variables
    Dim tituloResultado As String
    Dim mensajeResultado As String
    Dim mensajeSinCoincidencias As String
    Dim mensajeusuario As String
    
    ' Inicializar variables generales
    tituloResultado = "Resultado de la operación"
    mensajeResultado = "Estilos eliminados: (<contador>)"
    mensajeSinCoincidencias = "No se encontraron estilos para eliminar"
    
    ' Mostrar mensaje resultado
    Select Case ContadorEstilosEliminados
        Case 0
            MsgBox mensajeSinCoincidencias, vbInformation, tituloResultado
        Case Is > 0
            mensajeusuario = Replace(mensajeResultado, "<contador>", ContadorEstilosEliminados)
            MsgBox mensajeusuario, vbInformation, tituloResultado
    End Select
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: PrepararTablaEstructurada
' Purpose: Agregar una tabla estructurada para registrar información resultados
' Procedure Kind: Function
' Procedure Access: Private
' Parameter hojaResultados (Worksheet): Hoja donde se almacena la tabla
' Parameter nombreTabla (String): Nombre de la tabla
' Parameter listadoColumnas (Variant): Listado con nombres de columnas
' Return Type: ListObject
' Author: RicardoDiaz
' Date: 10/09/2019
' ----------------------------------------------------------------
Private Function PrepararTablaEstructurada(hojaResultados As Worksheet, nombreTabla As String, listadoColumnas As Variant) As ListObject

    Dim tablaEstructurada As ListObject
    
    ' Declarar variables
    Dim tablaExiste As Boolean
    
    tablaExiste = Evaluate("ISREF(" & nombreTabla & ")")
    
    If tablaExiste = False Then
        ' Agregar tabla estructurada
        Set tablaEstructurada = hojaResultados.ListObjects.Add
        
        With tablaEstructurada
            .Name = nombreTabla
            tablaEstructurada.Resize .Range.Resize(, UBound(listadoColumnas) + 1)
            .HeaderRowRange.Value2 = listadoColumnas
        End With
    Else
        Set tablaEstructurada = hojaResultados.ListObjects(nombreTabla)
    End If
    
    Set PrepararTablaEstructurada = tablaEstructurada

End Function

' ----------------------------------------------------------------
' Procedure Name: Reemplazar
' Purpose: Reemplazar según una tabla estilos actuales por nuevos
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter BuscarEnLibro (Boolean): Buscar en todo el libro (true) o solo la hoja activa (false)
' Author: RicardoDiaz
' Date: 10/09/2019
' ----------------------------------------------------------------
Public Sub Reemplazar(Optional BuscarEnLibro As Boolean = True)
    
    ' Declarar objetos
    Dim colHojas As Collection
    
    Dim hojaEstilos As Worksheet
    Dim nombreHojaEstilos As String
    
    ' Declarar variables
    Dim tablaEstilos As ListObject
    Dim nombreTablaEstilos As String
    Dim listadoColumnasEstilos As Variant
    
    ' Inicializar variables generales
    nombreHojaEstilos = "Estilos"
    nombreTablaEstilos = "TablaEstilos"
    listadoColumnasEstilos = Array("Estilo origen", "Estilo reemplazar")
    
    'mensajeSinCoincidencias = "No se encontraron celdas con estilos que cumplan los criterios definidos"
    
    If BuscarEnLibro = True Then
        ' Encontrar hojas que cumplen criterio de caracteres en el nombre (devuelve colección)
        Set colHojas = mHojas.EncontrarHojasPorCriterio(evExcluye, nombreHojaEstilos)
    Else
        Set colHojas = New Collection
        colHojas.Add ActiveSheet
    End If
    
    ' Agregar hoja de búsqueda/reemplazo de estilos
    Set hojaEstilos = mHojas.AgregarReferenciarHoja(nombreHojaEstilos, False, False)
    
    ' Preparar tabla de estilos
    Set tablaEstilos = PrepararTablaEstructurada(hojaEstilos, nombreTablaEstilos, listadoColumnasEstilos)
    
    If tablaEstilos.ListRows.Count = 0 Then
        MsgBox "No hay estilos para reemplazar, se ha agregado la hoja Estilos para definir los estilos que quiera reemplazar. Diligencie la tabla y después vuelva a correr este proceso"
        hojaEstilos.Activate
        Exit Sub
    End If
    
    Call ReemplazarEstilosEnHojas(tablaEstilos, colHojas)




End Sub

' ----------------------------------------------------------------
' Procedure Name: ReemplazarEstiloEnHoja
' Purpose: Reemplazar un estilo por otro en una hoja determinada
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter hoja (Worksheet): Hoja donde se realizará el reemplazo
' Parameter estiloOriginal (Style): Estilo original
' Parameter estiloNuevo (Style): Estilo nuevo
' Author: RicardoDiaz
' Date: 10/09/2019
' ----------------------------------------------------------------
Private Sub ReemplazarEstiloEnHoja(hoja As Worksheet, estiloOriginal As Style, estiloNuevo As Style)
    
    Dim celdaEvaluada As Range
    
    For Each celdaEvaluada In hoja.UsedRange.Cells
    
        If celdaEvaluada.Style = estiloOriginal And celdaEvaluada.MergeCells = False Then
            celdaEvaluada.Style = estiloNuevo
        End If
    
    Next celdaEvaluada


End Sub

' ----------------------------------------------------------------
' Procedure Name: ReemplazarEstilosEnHojas
' Purpose: Reemplazar los estilos definidos en una tabla (anteriores y nuevos) en una colección de hojas
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter tablaEstilos (ListObject): Tabla que contiene los estilos anteriores (actuales) y los nuevos que los reemplazarán
' Parameter colHojas (Collection): Colección de hojas dónde se realizará el reemplazo
' Author: RicardoDiaz
' Date: 10/09/2019
' ----------------------------------------------------------------
Private Sub ReemplazarEstilosEnHojas(tablaEstilos As ListObject, colHojas As Collection)

    Dim hoja As Worksheet
    Dim celda As Range

    Dim estiloAnterior As Style
    Dim estiloNuevo As Style

    For Each celda In tablaEstilos.DataBodyRange.Columns(1).Cells

        Set estiloAnterior = ThisWorkbook.Styles(celda.Value)
        Set estiloNuevo = ThisWorkbook.Styles(celda.Offset(0, 1).Value)
    
        For Each hoja In colHojas
            Call ReemplazarEstiloEnHoja(hoja, estiloAnterior, estiloNuevo)
        Next hoja

    Next celda
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: RegistrarInfoEstiloCelda
' Purpose: Registrar la información de las celdas que cumplen los criterios en la tabla de resultados
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter colCeldas (Collection): Colección de celdas que cumplen los criterios
' Parameter tablaResultados (ListObject): Tabla de resultados para registrar la información
' Author: RicardoDiaz
' Date: 10/09/2019
' ----------------------------------------------------------------
Private Sub RegistrarInfoEstiloCelda(colCeldas As Collection, tablaResultados As ListObject)
    
    Dim celda As Range
    Dim fila As ListRow
    
    Dim contador As Long
    
    For Each celda In colCeldas
        Set fila = tablaResultados.ListRows.Add
        With fila
            .Range.Columns(1).Value2 = celda.Parent.Name
            .Range.Columns(2).Value2 = celda.Address
            .Range.Columns(3).Value2 = celda.Style.Name
        End With
        contador = contador + 1
    Next celda
    
End Sub



Public Sub Duplicar()

    ' Declarar objetos
    Dim colHojas As Collection
    
    Dim hojaEstilos As Worksheet
    Dim nombreHojaEstilos As String
    
    ' Declarar variables
    Dim tablaEstilos As ListObject
    Dim nombreTablaEstilos As String
    Dim listadoColumnasEstilos As Variant
    
    ' Inicializar variables generales
    nombreHojaEstilos = "Estilos"
    nombreTablaEstilos = "TablaEstilos"
    listadoColumnasEstilos = Array("Estilo origen", "Nuevo estilo")
    
    'mensajeSinCoincidencias = "No se encontraron celdas con estilos que cumplan los criterios definidos"
    
    
    ' Agregar hoja de búsqueda/reemplazo de estilos
    Set hojaEstilos = mHojas.AgregarReferenciarHoja(nombreHojaEstilos, False, False)
    
    ' Preparar tabla de estilos
    Set tablaEstilos = PrepararTablaEstructurada(hojaEstilos, nombreTablaEstilos, listadoColumnasEstilos)
    
    If tablaEstilos.ListRows.Count = 0 Then
        MsgBox "No hay estilos para duplicar, se ha agregado la hoja Estilos para definir los estilos que quiera duplicar. Diligencie la tabla y después vuelva a correr este proceso"
        hojaEstilos.Activate
        Exit Sub
    End If
    

    Dim celda As Range
    
    Dim nombreEstiloOriginal As String
    Dim nombreEstiloNuevo As String

    For Each celda In tablaEstilos.ListColumns(1).DataBodyRange.Cells
        
        nombreEstiloOriginal = celda.Value2
        nombreEstiloNuevo = celda.Offset(0, 1).Value2
        
        Call DuplicarEstilo(nombreEstiloOriginal, nombreEstiloNuevo)

    Next celda



End Sub


' ----------------------------------------------------------------
' Procedure Name: DuplicarEstiloEnHoja
' Purpose: Duplicar un estilo con un nuevo nombre
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter nombreEstiloOriginal (String): Estilo original
' Parameter nombreEstiloNuevo (String): Estilo nuevo
' Author: RicardoDiaz
' Date: 10/09/2019
' ----------------------------------------------------------------
Private Sub DuplicarEstilo(nombreEstiloOriginal As String, nombreEstiloNuevo As String)
    
    Dim estiloOriginal As Style
    Dim estiloNuevo As Style
    
    On Error Resume Next
    
    ' Recrear estilo con nuevo nombre
    Set estiloOriginal = ThisWorkbook.Styles(nombreEstiloOriginal)
    Set estiloNuevo = ThisWorkbook.Styles.Add(Name:=nombreEstiloNuevo)
    
    estiloNuevo.Interior.i
    
    With estiloNuevo
        .AddIndent = estiloOriginal.AddIndent
        .Borders(xlDiagonalDown).Color = estiloOriginal.Borders(xlDiagonalDown).Color
        .Borders(xlDiagonalDown).ColorIndex = estiloOriginal.Borders(xlDiagonalDown).ColorIndex
        .Borders(xlDiagonalDown).LineStyle = estiloOriginal.Borders(xlDiagonalDown).LineStyle
        .Borders(xlDiagonalDown).ThemeColor = estiloOriginal.Borders(xlDiagonalDown).ThemeColor
        .Borders(xlDiagonalDown).TintAndShade = estiloOriginal.Borders(xlDiagonalDown).TintAndShade
        .Borders(xlDiagonalDown).Weight = estiloOriginal.Borders(xlDiagonalDown).Weight
    
        .Borders(xlDiagonalUp).Color = estiloOriginal.Borders(xlDiagonalUp).Color
        .Borders(xlDiagonalUp).ColorIndex = estiloOriginal.Borders(xlDiagonalUp).ColorIndex
        .Borders(xlDiagonalUp).LineStyle = estiloOriginal.Borders(xlDiagonalUp).LineStyle
        .Borders(xlDiagonalUp).ThemeColor = estiloOriginal.Borders(xlDiagonalUp).ThemeColor
        .Borders(xlDiagonalUp).TintAndShade = estiloOriginal.Borders(xlDiagonalUp).TintAndShade
        .Borders(xlDiagonalUp).Weight = estiloOriginal.Borders(xlDiagonalUp).Weight
    
        .Borders(xlEdgeBottom).Color = estiloOriginal.Borders(xlEdgeBottom).Color
        .Borders(xlEdgeBottom).ColorIndex = estiloOriginal.Borders(xlEdgeBottom).ColorIndex
        .Borders(xlEdgeBottom).LineStyle = estiloOriginal.Borders(xlEdgeBottom).LineStyle
        .Borders(xlEdgeBottom).ThemeColor = estiloOriginal.Borders(xlEdgeBottom).ThemeColor
        .Borders(xlEdgeBottom).TintAndShade = estiloOriginal.Borders(xlEdgeBottom).TintAndShade
        .Borders(xlEdgeBottom).Weight = estiloOriginal.Borders(xlEdgeBottom).Weight
    
        .Borders(xlEdgeLeft).Color = estiloOriginal.Borders(xlEdgeLeft).Color
        .Borders(xlEdgeLeft).ColorIndex = estiloOriginal.Borders(xlEdgeLeft).ColorIndex
        .Borders(xlEdgeLeft).LineStyle = estiloOriginal.Borders(xlEdgeLeft).LineStyle
        .Borders(xlEdgeLeft).ThemeColor = estiloOriginal.Borders(xlEdgeLeft).ThemeColor
        .Borders(xlEdgeLeft).TintAndShade = estiloOriginal.Borders(xlEdgeLeft).TintAndShade
        .Borders(xlEdgeLeft).Weight = estiloOriginal.Borders(xlEdgeLeft).Weight
    
        .Borders(xlEdgeRight).Color = estiloOriginal.Borders(xlEdgeRight).Color
        .Borders(xlEdgeRight).ColorIndex = estiloOriginal.Borders(xlEdgeRight).ColorIndex
        .Borders(xlEdgeRight).LineStyle = estiloOriginal.Borders(xlEdgeRight).LineStyle
        .Borders(xlEdgeRight).ThemeColor = estiloOriginal.Borders(xlEdgeRight).ThemeColor
        .Borders(xlEdgeRight).TintAndShade = estiloOriginal.Borders(xlEdgeRight).TintAndShade
        .Borders(xlEdgeRight).Weight = estiloOriginal.Borders(xlEdgeRight).Weight
    
        .Borders(xlEdgeTop).Color = estiloOriginal.Borders(xlEdgeTop).Color
        .Borders(xlEdgeTop).ColorIndex = estiloOriginal.Borders(xlEdgeTop).ColorIndex
        .Borders(xlEdgeTop).LineStyle = estiloOriginal.Borders(xlEdgeTop).LineStyle
        .Borders(xlEdgeTop).ThemeColor = estiloOriginal.Borders(xlEdgeTop).ThemeColor
        .Borders(xlEdgeTop).TintAndShade = estiloOriginal.Borders(xlEdgeTop).TintAndShade
        .Borders(xlEdgeTop).Weight = estiloOriginal.Borders(xlEdgeTop).Weight
    
        .Borders(xlInsideHorizontal).Color = estiloOriginal.Borders(xlInsideHorizontal).Color
        .Borders(xlInsideHorizontal).ColorIndex = estiloOriginal.Borders(xlInsideHorizontal).ColorIndex
        .Borders(xlInsideHorizontal).LineStyle = estiloOriginal.Borders(xlInsideHorizontal).LineStyle
        .Borders(xlInsideHorizontal).ThemeColor = estiloOriginal.Borders(xlInsideHorizontal).ThemeColor
        .Borders(xlInsideHorizontal).TintAndShade = estiloOriginal.Borders(xlInsideHorizontal).TintAndShade
        .Borders(xlInsideHorizontal).Weight = estiloOriginal.Borders(xlInsideHorizontal).Weight
    
        .Borders(xlInsideVertical).Color = estiloOriginal.Borders(xlInsideVertical).Color
        .Borders(xlInsideVertical).ColorIndex = estiloOriginal.Borders(xlInsideVertical).ColorIndex
        .Borders(xlInsideVertical).LineStyle = estiloOriginal.Borders(xlInsideVertical).LineStyle
        .Borders(xlInsideVertical).ThemeColor = estiloOriginal.Borders(xlInsideVertical).ThemeColor
        .Borders(xlInsideVertical).TintAndShade = estiloOriginal.Borders(xlInsideVertical).TintAndShade
        .Borders(xlInsideVertical).Weight = estiloOriginal.Borders(xlInsideVertical).Weight
    
        '.Creator = estiloOriginal.Creator
        '.Font.Background = estiloOriginal.Font.Background
        .Font.Bold = estiloOriginal.Font.Bold
        .Font.Color = estiloOriginal.Font.Color
        .Font.ColorIndex = estiloOriginal.Font.ColorIndex
        '.Font.FontStyle = estiloOriginal.Font.FontStyle
        .Font.Italic = estiloOriginal.Font.Italic
        .Font.Name = estiloOriginal.Font.Name
        .Font.Size = estiloOriginal.Font.Size
        .Font.Strikethrough = estiloOriginal.Font.Strikethrough
        .Font.Subscript = estiloOriginal.Font.Subscript
        .Font.Superscript = estiloOriginal.Font.Superscript
        .Font.ThemeColor = estiloOriginal.Font.ThemeColor
        .Font.ThemeFont = estiloOriginal.Font.ThemeFont
        .Font.TintAndShade = estiloOriginal.Font.TintAndShade
        .Font.Underline = estiloOriginal.Font.Underline
    
        .FormulaHidden = estiloOriginal.FormulaHidden
        .HorizontalAlignment = estiloOriginal.HorizontalAlignment
        
        .IncludeAlignment = estiloOriginal.IncludeAlignment
        .IncludeBorder = estiloOriginal.IncludeBorder
        .IncludeFont = estiloOriginal.IncludeFont
        .IncludeNumber = estiloOriginal.IncludeNumber
        .IncludePatterns = estiloOriginal.IncludePatterns
        .IncludeProtection = estiloOriginal.IncludeProtection
        '.IndentLevel = estiloOriginal.IndentLevel
        .Interior.Color = estiloOriginal.Interior.Color
        '.Interior.ColorIndex = estiloOriginal.Interior.ColorIndex
        '.Interior.Gradient = estiloOriginal.Interior.Gradient
        '.Interior.InvertIfNegative = estiloOriginal.Interior.InvertIfNegative
        .Interior.Pattern = estiloOriginal.Interior.Pattern
        .Interior.PatternColor = estiloOriginal.Interior.PatternColor
        .Interior.PatternColorIndex = estiloOriginal.Interior.PatternColorIndex
        '.Interior.PatternThemeColor = estiloOriginal.Interior.PatternThemeColor
        .Interior.PatternTintAndShade = estiloOriginal.Interior.PatternTintAndShade
        .Interior.ThemeColor = estiloOriginal.Interior.ThemeColor
        .Interior.TintAndShade = estiloOriginal.Interior.TintAndShade
    
        .Locked = estiloOriginal.Locked
        '.MergeCells = estiloOriginal.MergeCells
        .NumberFormat = estiloOriginal.NumberFormat
        .NumberFormatLocal = estiloOriginal.NumberFormatLocal
        .Orientation = estiloOriginal.Orientation
        .ReadingOrder = estiloOriginal.ReadingOrder
        .ShrinkToFit = estiloOriginal.ShrinkToFit
        .VerticalAlignment = estiloOriginal.VerticalAlignment
        .WrapText = estiloOriginal.WrapText
    
    
    End With


End Sub
