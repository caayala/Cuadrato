Option Explicit

Function cuadrato(fil1 As String, _ 
                  fil2 As Variant, _
                  col1 As String, _ 
                  rgMat As string, _ 
                  Optional posicion As Integer = 0, _
                  Optional error As Boolean = False) As Variant

' Funcion para encontrar valores en una tabla personalizada de SPSS con un nivel de segmento
' by Cristian Ayala

Dim fil As Integer, col As Integer, aux As Integer
Dim rg As Range
Dim Msg As String

' Set up error handling
' On Error Resume Next
On Error GoTo BadEntry

Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.ScreenUpdating = False

Set rg = Range(rgMat)
With rg
    ' Busca el numero de fila en la primera columna
    fil = WorksheetFunction.Match(fil1, .Columns(1).value, 0)
    ' Aux es el numero de fila donde termina el rango de la pregunta en la que busco
    aux = .Cells(fil, 1).MergeArea.Count
    
    ' Busca el numero de fila de en la segunda columna
    fil = WorksheetFunction.Match(fil2, Range(.Cells(fil, 2), .Cells(fil + aux - 1, 2)).value, 0) + fil - 1
    
    ' Busca el numero de columna en la segunda fila
    col = WorksheetFunction.Match(col1, .Rows(2).value, 0)
    
    ' Ajusta el numero de columna a la posicion en que esta el dato de interes
    cuadrato = .Cells(fil, col + posicion).Value

    If Right(.Cells(fil, col + posicion).NumberFormat, 1) = "%" Then
        cuadrato = cuadrato * 100
    End If

End With

Set rg = Nothing

GoTo Done

BadEntry:
    Set rg = Nothing
    
    If error = True Then
        cuadrato = 0
    Else
        cuadrato = vbNullString ' Igual que "" pero no ocupa memoria.
    End If

Done:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

'    Msg = "An error occurred." & vbNewLine
'    MsgBox Msg
End Function

Function cuadratoRango( _ 
    fil1 As String, fil2_range As Range, _
    col1 As String, _ 
    rgMat As String, Optional posicion As Integer = 0, _
    Optional error As Boolean = False) As Variant

' suma los valores de la preguta (fil1) asignados 
' al rango de alternativas ingresadas (fil2_range) para el segmento (col1)
' necesita de funcion cuadrato para funcionar

Dim fil1 As Range

For Each fil1 In fil1_range.Cells
    If IsEmpty(fil1) = False Then
        cuadratoRango = cuadrato(fil1, fil2, col1, rgMat, posicion, error) + cuadratoRango
    End If
Next

End Function


Function cuadrato2( _
    fil1 As String, fil2 As Variant, _
    col1 As String, col2 As String, _
    rgMat As String, _
    Optional posicion As Integer = 0, Optional error As Boolean = False) As Variant

' Funcion para encontrar valores en una tabla personalizada de SPSS con dos niveles de segmentos
' by Cristian Ayala

Dim fil As Integer, col As Integer, aux As Integer
Dim rg As Range
Dim Msg As String

' Set up error handling
' On Error Resume Next
On Error GoTo BadEntry

Application.Calculation = xlCalculationManual
' Application.EnableEvents = False
Application.ScreenUpdating = False

Set rg = Range(rgMat)

With rg
    ' Busca el numero de fila en la primera columna
    fil = WorksheetFunction.Match(fil1, .Columns(1).Value, 0)
    ' Aux es el numero de fila donde termina el rango de la pregunta en la que busco
    aux = .Cells(fil, 1).MergeArea.Count
    
    ' Busca el numero de fila de en la segunda columna
    fil = WorksheetFunction.Match(fil2.Value, Range(.Cells(fil, 2), .Cells(fil + aux - 1, 2)).Value, 0) + fil - 1

    ' Busca el numero de columna en la segunda fila
    col = WorksheetFunction.Match(col1, .Rows(2).Value, 0)
    ' Aux es el numero de columna donde termina el rango del segmento en que busco
    aux = .Cells(2, col).MergeArea.Count

    ' Busca el numero de fila de en la cuarta columna
    col = WorksheetFunction.Match(col2, Range(.Cells(4, col), .Cells(4, col + aux - 1)).Value, 0) + col - 1
    
    ' Ajusta el numero de columna a la posicion en que esta el dato de interes
    cuadrato2 = .Cells(fil, col + posicion).Value

    If Right(.Cells(fil, col + posicion).NumberFormat, 1) = "%" Then
        cuadrato2 = cuadrato2 * 100
    End If

End With

Set rg = Nothing
Set rgaux = Nothing

GoTo Done

BadEntry:
    Set rg = Nothing
    Set rgaux = Nothing

    If error = True Then
        cuadrato2 = 0
    Else
        cuadrato2 = vbNullString ' Igual que "" pero no ocupa memoria.
    End If

'    Msg = "An error occurred." & vbNewLine
'    MsgBox Msg

Done:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ' Application.EnableEvents = True

End Function

Function cuadrato2Rango( _ 
    fil1 As String, fil2_range As Range, _
    col1 As String, col2 As String, _ 
    rgMat As String, _ 
    Optional posicion As Integer = 0, _ 
    Optional error As Boolean = False) As Variant

' suma los valores asignados al rango de alternativas ingresadas.
' necesita de funcion cuadrato2 para funcionar

Dim fil2 As Range

For Each fil2 In fil2_range.Cells
    If IsEmpty(fil2) = False Then
        cuadrato2Rango = cuadrato2(fil1, fil2, col1, col2, rgMat, posicion, error) + cuadrato2Rango
    End If
Next

End Function

Function ratocuad(col1 As String, fila1 As String, _
    fila2 As String, rgMat As Range, _
    Optional error As Boolean = False) As Variant

' by Cristian Ayala

Dim fil As Integer, col As Integer, aux As Integer
Dim rg As Range
Dim Msg As String

' Set up error handling
' On Error Resume Next
On Error GoTo BadEntry

Set rg = rgMat

With rg
    ' Busca el numero de fila en la primera columna
    fil = WorksheetFunction.Match(col1, .Columns(1), 0)
    
    ' Busca el numero de columna en la segunda fila
    col = WorksheetFunction.Match(fila1, .Rows(2), 0) + 1
    ' Aux es el numero de columna donde termina el rango de la pregunta en la que busco
        aux = .Cells(2, col).MergeArea.Count

    ' Busca el numero de columna de en la tercera fila
    col = WorksheetFunction.Match(fila2, Range(.Cells(3, col), .Cells(3, col + aux - 1)), 0) + col - 1

    ratocuad = .Cells(fil, col).Value

End With

Set rg = Nothing

Exit Function

BadEntry:
    If error = True Then
        ratocuad = 0
    Else
        ratocuad = vbNullString ' Igual que "" pero no ocupa memoria.
    End If

'    Msg = "An error occurred." & vbNewLine
'    MsgBox Msg
End Function