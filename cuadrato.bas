Option Explicit

'============================================================================
' cuadrato.bas
' Funciones VBA para extraer valores desde tablas personalizadas de SPSS
' hacia planillas de informe en Excel.
'
' Autor: Cristian Ayala
' Revisado y optimizado: 2026-07-24
'
' Cambios de esta version:
'  - cuadratoRango reescrita (antes no compilaba: fil1 duplicada, fil1_range
'    / fil2 / posicion / error inexistentes en el ambito).
'  - Eliminada la variable rgaux no declarada en cuadrato2 (rompia Option
'    Explicit).
'  - Corregido el espacio despues del "_" de continuacion en varias lineas.
'  - Quitados los toggles de Application.* (Calculation / ScreenUpdating /
'    EnableEvents): dentro de una UDF son ineficaces y ademas forzaban el
'    libro a calculo Automatico al terminar.
'  - Busqueda con Application.Match + IsError en vez de WorksheetFunction.Match
'    con On Error: evita el costo de lanzar y atrapar una excepcion cada vez
'    que un lookup no existe (que es el caso frecuente). Esta es la principal
'    ganancia de rendimiento.
'  - fil / col declaradas como Long (antes Integer, con overflow > 32.767).
'  - fil2 acepta tanto un valor como una celda/Range (antes cuadrato2 exigia
'    un Range y fallaba si se le pasaba un literal).
'  - Las funciones ...Rango suman solo valores numericos (evita el type
'    mismatch de "" + numero) y devuelven 0 cuando no hay coincidencias.
'  - Parametro 'error' renombrado a 'alError': 'error' pisa la funcion Error
'    de VBA. Las llamadas desde celdas son posicionales, no se ven afectadas.
'============================================================================


' ---------------------------------------------------------------------------
' cuadrato: valor de la tabla para pregunta (fil1), alternativa (fil2)
'           y un nivel de segmento (col1).
' ---------------------------------------------------------------------------
Function cuadrato(fil1 As String, fil2 As Variant, col1 As String, rgMat As String, _
                  Optional posicion As Long = 0, Optional alError As Boolean = False) As Variant

    Dim rg As Range, celda As Range
    Dim mFil As Variant, mSub As Variant, mCol As Variant
    Dim fil As Long, col As Long, aux As Long

    On Error GoTo Falla
    Set rg = Range(rgMat)

    ' Fila de la pregunta en la primera columna.
    mFil = Application.Match(fil1, rg.Columns(1), 0)
    If IsError(mFil) Then GoTo Falla
    fil = CLng(mFil)

    ' Numero de filas del bloque combinado de la pregunta.
    aux = rg.Cells(fil, 1).MergeArea.Count

    ' Fila de la alternativa dentro del bloque, en la segunda columna.
    mSub = Application.Match(NormalizaClave(fil2), rg.Cells(fil, 2).Resize(aux, 1), 0)
    If IsError(mSub) Then GoTo Falla
    fil = CLng(mSub) + fil - 1

    ' Columna del segmento en la segunda fila.
    mCol = Application.Match(col1, rg.Rows(2), 0)
    If IsError(mCol) Then GoTo Falla
    col = CLng(mCol) + posicion

    Set celda = rg.Cells(fil, col)
    cuadrato = EscalaPorcentaje(celda.Value, celda)
    Exit Function

Falla:
    cuadrato = ResultadoVacio(alError)
End Function


' ---------------------------------------------------------------------------
' cuadratoRango: suma cuadrato() sobre un rango de alternativas (fil2_range)
'                para la misma pregunta y el mismo segmento.
' ---------------------------------------------------------------------------
Function cuadratoRango(fil1 As String, fil2_range As Range, col1 As String, rgMat As String, _
                       Optional posicion As Long = 0, Optional alError As Boolean = False) As Variant

    Dim celda As Range, v As Variant, total As Double

    For Each celda In fil2_range.Cells
        If Not IsEmpty(celda.Value) Then
            v = cuadrato(fil1, celda.Value, col1, rgMat, posicion, alError)
            If IsNumeric(v) Then total = total + v
        End If
    Next celda

    cuadratoRango = total
End Function


' ---------------------------------------------------------------------------
' cuadrato2: valor de la tabla con dos niveles de segmento (col1 y col2).
' ---------------------------------------------------------------------------
Function cuadrato2(fil1 As String, fil2 As Variant, col1 As String, col2 As String, rgMat As String, _
                   Optional posicion As Long = 0, Optional alError As Boolean = False) As Variant

    Dim rg As Range, celda As Range
    Dim m As Variant
    Dim fil As Long, col As Long, aux As Long

    On Error GoTo Falla
    Set rg = Range(rgMat)

    ' Fila de la pregunta en la primera columna.
    m = Application.Match(fil1, rg.Columns(1), 0)
    If IsError(m) Then GoTo Falla
    fil = CLng(m)
    aux = rg.Cells(fil, 1).MergeArea.Count

    ' Fila de la alternativa dentro del bloque, en la segunda columna.
    m = Application.Match(NormalizaClave(fil2), rg.Cells(fil, 2).Resize(aux, 1), 0)
    If IsError(m) Then GoTo Falla
    fil = CLng(m) + fil - 1

    ' Columna del primer segmento en la segunda fila.
    m = Application.Match(col1, rg.Rows(2), 0)
    If IsError(m) Then GoTo Falla
    col = CLng(m)
    aux = rg.Cells(2, col).MergeArea.Count

    ' Columna del segundo segmento dentro del bloque, en la cuarta fila.
    m = Application.Match(col2, rg.Cells(4, col).Resize(1, aux), 0)
    If IsError(m) Then GoTo Falla
    col = CLng(m) + col - 1 + posicion

    Set celda = rg.Cells(fil, col)
    cuadrato2 = EscalaPorcentaje(celda.Value, celda)
    Exit Function

Falla:
    cuadrato2 = ResultadoVacio(alError)
End Function


' ---------------------------------------------------------------------------
' cuadrato2Rango: suma cuadrato2() sobre un rango de alternativas.
' ---------------------------------------------------------------------------
Function cuadrato2Rango(fil1 As String, fil2_range As Range, col1 As String, col2 As String, rgMat As String, _
                        Optional posicion As Long = 0, Optional alError As Boolean = False) As Variant

    Dim celda As Range, v As Variant, total As Double

    For Each celda In fil2_range.Cells
        If Not IsEmpty(celda.Value) Then
            v = cuadrato2(fil1, celda.Value, col1, col2, rgMat, posicion, alError)
            If IsNumeric(v) Then total = total + v
        End If
    Next celda

    cuadrato2Rango = total
End Function


' ---------------------------------------------------------------------------
' ratocuad: variante con la pregunta en columnas y los segmentos en filas.
' NOTA: se conserva el "+ 1" del codigo original en la busqueda de fila1.
'       No pude verificarlo contra una tabla real; si los resultados
'       aparecen corridos una columna, revisar esa linea.
' ---------------------------------------------------------------------------
Function ratocuad(col1 As String, fila1 As String, fila2 As String, rgMat As Range, _
                  Optional alError As Boolean = False) As Variant

    Dim rg As Range
    Dim mFil As Variant, mCol As Variant, mSub As Variant
    Dim fil As Long, col As Long, aux As Long

    On Error GoTo Falla
    Set rg = rgMat

    ' Fila de interes en la primera columna.
    mFil = Application.Match(col1, rg.Columns(1), 0)
    If IsError(mFil) Then GoTo Falla
    fil = CLng(mFil)

    ' Columna en la segunda fila (se conserva el + 1 del codigo original).
    mCol = Application.Match(fila1, rg.Rows(2), 0)
    If IsError(mCol) Then GoTo Falla
    col = CLng(mCol) + 1
    aux = rg.Cells(2, col).MergeArea.Count

    ' Columna del subnivel dentro del bloque, en la tercera fila.
    mSub = Application.Match(fila2, rg.Cells(3, col).Resize(1, aux), 0)
    If IsError(mSub) Then GoTo Falla
    col = CLng(mSub) + col - 1

    ratocuad = rg.Cells(fil, col).Value
    Exit Function

Falla:
    ratocuad = ResultadoVacio(alError)
End Function


'============================================================================
' Auxiliares (privadas: no aparecen como funciones de hoja).
'============================================================================

' Devuelve el valor de una clave, sea un valor directo o una celda/Range.
Private Function NormalizaClave(v As Variant) As Variant
    If IsObject(v) Then
        NormalizaClave = v.Value
    Else
        NormalizaClave = v
    End If
End Function

' Multiplica por 100 si la celda tiene formato de porcentaje.
Private Function EscalaPorcentaje(ByVal valor As Variant, celda As Range) As Variant
    If IsNumeric(valor) Then
        If Right$(celda.NumberFormat, 1) = "%" Then
            valor = valor * 100
        End If
    End If
    EscalaPorcentaje = valor
End Function

' Resultado cuando el lookup falla: 0 si alError = True, en caso contrario vacio.
Private Function ResultadoVacio(ByVal alError As Boolean) As Variant
    If alError Then
        ResultadoVacio = 0
    Else
        ResultadoVacio = vbNullString
    End If
End Function
