Sub ProcesarReservas()
    ' ------------------------------------------------------------------------
    ' Macro: ProcesarReservas
    ' Descripción: Lee la hoja "Sheet" del libro activo, filtra por marca
    '              (ALAMO, ENTERPRISE, NATIONAL) según la ubicación de recogida,
    '              ordena los datos, agrega columnas de formato y crea un nuevo
    '              libro con tres pestañas con el formato y colores especificados.
    ' ------------------------------------------------------------------------
    
    Dim wbOrigen As Workbook
    Dim wsOrigen As Worksheet
    Dim wbDestino As Workbook
    Dim wsAlamo As Worksheet, wsEnterprise As Worksheet, wsNational As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim datos() As Variant
    Dim filasAlamo() As Variant, filasEnterprise() As Variant, filasNational() As Variant
    Dim countAlamo As Long, countEnterprise As Long, countNational As Long
    Dim ubicRecCol As Long, agenciaCol As Long
    Dim simboloFLCol As Long, simboloMarcaCol As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set wbOrigen = ThisWorkbook
    On Error Resume Next
    Set wsOrigen = wbOrigen.Sheets("Sheet")
    On Error GoTo 0
    If wsOrigen Is Nothing Then
        MsgBox "No se encontró la hoja 'Sheet' en el libro origen.", vbCritical
        Exit Sub
    End If
    
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, "B").End(xlUp).Row
    If ultimaFila < 2 Then
        MsgBox "No hay datos en la hoja.", vbCritical
        Exit Sub
    End If
    
    simboloFLCol = 1
    simboloMarcaCol = 3
    ubicRecCol = 10
    agenciaCol = 18
    
    Dim rangoDatos As Range
    Set rangoDatos = wsOrigen.Range(wsOrigen.Cells(2, 1), wsOrigen.Cells(ultimaFila, agenciaCol))
    datos = rangoDatos.Value
    
    ReDim filasAlamo(1 To UBound(datos, 1))
    ReDim filasEnterprise(1 To UBound(datos, 1))
    ReDim filasNational(1 To UBound(datos, 1))
    countAlamo = 0
    countEnterprise = 0
    countNational = 0
    
    For i = 1 To UBound(datos, 1)
        Dim ubicRec As String
        ubicRec = Trim(CStr(datos(i, ubicRecCol)))
        If ubicRec <> "" Then
            If InStr(1, ubicRec, "E71", vbTextCompare) > 0 Or InStr(1, ubicRec, "T71", vbTextCompare) > 0 Then
                countAlamo = countAlamo + 1
                filasAlamo(countAlamo) = i
            ElseIf InStr(1, ubicRec, "C61", vbTextCompare) > 0 Or InStr(1, ubicRec, "T61", vbTextCompare) > 0 Then
                countEnterprise = countEnterprise + 1
                filasEnterprise(countEnterprise) = i
            ElseIf InStr(1, ubicRec, "E01", vbTextCompare) > 0 Or InStr(1, ubicRec, "T01", vbTextCompare) > 0 Then
                countNational = countNational + 1
                filasNational(countNational) = i
            End If
        End If
    Next i
    
    ReDim Preserve filasAlamo(1 To countAlamo)
    ReDim Preserve filasEnterprise(1 To countEnterprise)
    ReDim Preserve filasNational(1 To countNational)
    
    Set wbDestino = Workbooks.Add
    Do While wbDestino.Sheets.Count > 1
        wbDestino.Sheets(wbDestino.Sheets.Count).Delete
    Loop
    wbDestino.Sheets(1).Name = "ALAMO"
    Set wsAlamo = wbDestino.Sheets("ALAMO")
    Set wsEnterprise = wbDestino.Sheets.Add(After:=wsAlamo)
    wsEnterprise.Name = "ENTERPRISE"
    Set wsNational = wbDestino.Sheets.Add(After:=wsEnterprise)
    wsNational.Name = "NATIONAL"
    
    Call EscribirHoja(wsAlamo, datos, filasAlamo, "ALAMO", "I")
    Call EscribirHoja(wsEnterprise, datos, filasEnterprise, "ENTERPRISE", "P")
    Call EscribirHoja(wsNational, datos, filasNational, "NATIONAL", "E")
    
    Dim ws As Worksheet
    For Each ws In wbDestino.Sheets
        ws.Columns.AutoFit
    Next ws
    
    wbDestino.Sheets("ALAMO").Activate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Dim nombreArchivo As Variant
    nombreArchivo = Application.GetSaveAsFilename(InitialFileName:="Reservas_Procesadas.xlsx", _
        fileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If nombreArchivo <> False Then
        wbDestino.SaveAs nombreArchivo, FileFormat:=xlOpenXMLWorkbook
        MsgBox "Archivo guardado exitosamente en:" & vbCrLf & nombreArchivo, vbInformation
    Else
        MsgBox "No se guardó el archivo. Puede guardarlo manualmente.", vbExclamation
    End If
    
End Sub

Private Sub EscribirHoja(ws As Worksheet, datos As Variant, filas() As Variant, marcaNombre As String, letraMarca As String)
    
    Dim numFilas As Long
    numFilas = UBound(filas) - LBound(filas) + 1
    If numFilas = 0 Then Exit Sub
    
    Dim ordenArray() As Variant
    ReDim ordenArray(1 To numFilas, 1 To 11)
    
    Dim i As Long, filaOri As Long
    For i = 1 To numFilas
        filaOri = filas(i)
        ordenArray(i, 1) = datos(filaOri, 2)
        ordenArray(i, 2) = datos(filaOri, 4)
        ordenArray(i, 3) = datos(filaOri, 6)
        ordenArray(i, 4) = datos(filaOri, 8)
        ordenArray(i, 5) = datos(filaOri, 9)
        ordenArray(i, 6) = datos(filaOri, 10)
        ordenArray(i, 7) = datos(filaOri, 11)
        ordenArray(i, 8) = datos(filaOri, 13)
        ordenArray(i, 9) = datos(filaOri, 18)
        ordenArray(i, 10) = datos(filaOri, 1)
        ordenArray(i, 11) = datos(filaOri, 3)
    Next i
    
    Call QuickSortArray(ordenArray, LBound(ordenArray, 1), UBound(ordenArray, 1), 2)
    
    Dim encabezados As Variant
    encabezados = Array("Letra", "N.º reserva", "Marca", "**", "Nombre", "Clase", "**MAIN VIP**", _
                        "Fecha recogida", "Fecha entrega", "Ubic. rec.", "Ubic. ent.", "Tarifa Diaria", "Agencia de recom.")
    
    With ws
        For i = LBound(encabezados) To UBound(encabezados)
            .Cells(1, i + 1).Value = encabezados(i)
            .Cells(1, i + 1).Font.Bold = True
            .Cells(1, i + 1).Font.Name = "Calibri"
        Next i
    End With
    
    Dim estrella As String
    estrella = ChrW(&H2B50)
    
    Dim filaActual As Long, letraGrupo As String, letraAnterior As String
    Dim valorFL As String, valorMarcaSimbolo As String, valorCombinado As String
    Dim mainVip As String
    Dim ubicRecTemp As String, agenciaTemp As String
    
    For i = 1 To numFilas
        filaActual = i + 1
        
        ' ---------- Columna A (Letra) - Solo A-Z, ignora números y otros caracteres ----------
        letraGrupo = UCase(Left(Trim(ordenArray(i, 2)), 1))
        If i = 1 Then
            If letraGrupo >= "A" And letraGrupo <= "Z" Then
                ws.Cells(filaActual, 1).Value = letraGrupo
            End If
            letraAnterior = letraGrupo
        ElseIf letraGrupo <> letraAnterior Then
            If letraGrupo >= "A" And letraGrupo <= "Z" Then
                ws.Cells(filaActual, 1).Value = letraGrupo
            End If
            letraAnterior = letraGrupo
        Else
            ws.Cells(filaActual, 1).Value = ""
        End If
        
        ' ---------- Columna B ----------
        ws.Cells(filaActual, 2).Value = ordenArray(i, 1)
        
        ' ---------- Columna C ----------
        ws.Cells(filaActual, 3).Value = marcaNombre & "/"
        
        ' ---------- Columna D ----------
        valorFL = ""
        Dim simboloFL As String
        simboloFL = CStr(ordenArray(i, 10))
        If InStr(1, simboloFL, "*", vbTextCompare) > 0 Or InStr(1, simboloFL, "~", vbTextCompare) > 0 Then
            valorFL = estrella & "FL" & estrella
        End If
        valorMarcaSimbolo = ""
        If InStr(1, CStr(ordenArray(i, 11)), "+", vbTextCompare) > 0 Then
            valorMarcaSimbolo = letraMarca
        End If
        
        If valorFL <> "" And valorMarcaSimbolo <> "" Then
            valorCombinado = estrella & valorMarcaSimbolo & "-FL" & estrella
        ElseIf valorFL <> "" Then
            valorCombinado = valorFL
        ElseIf valorMarcaSimbolo <> "" Then
            valorCombinado = valorMarcaSimbolo
        Else
            valorCombinado = ""
        End If
        ws.Cells(filaActual, 4).Value = valorCombinado
        
        ' ---------- Columna E ----------
        ws.Cells(filaActual, 5).Value = ordenArray(i, 2)
        
        ' ---------- Columna F ----------
        ws.Cells(filaActual, 6).Value = ordenArray(i, 3)
        
        ' ---------- Columna G ----------
        mainVip = ""
        ubicRecTemp = CStr(ordenArray(i, 6))
        agenciaTemp = CStr(ordenArray(i, 9))
        If Right(ubicRecTemp, 3) = "E71" Or Right(ubicRecTemp, 3) = "C61" Or Right(ubicRecTemp, 3) = "E01" Then
            mainVip = "* MAIN VIP*"
        End If
        If InStr(1, agenciaTemp, "11617270", vbTextCompare) > 0 Then
            If mainVip <> "" Then
                mainVip = mainVip & " * EXPEDIA*"
            Else
                mainVip = "* EXPEDIA*"
            End If
        End If
        ws.Cells(filaActual, 7).Value = mainVip
        
        ' ---------- Columnas H a M ----------
        ws.Cells(filaActual, 8).Value = ordenArray(i, 4)
        ws.Cells(filaActual, 9).Value = ordenArray(i, 5)
        ws.Cells(filaActual, 10).Value = ordenArray(i, 6)
        ws.Cells(filaActual, 11).Value = ordenArray(i, 7)
        ws.Cells(filaActual, 12).Value = ordenArray(i, 8)
        ws.Cells(filaActual, 13).Value = ordenArray(i, 9)
    Next i
    
    ' ---------- FORMATOS GENERALES ----------
    Dim rngTodo As Range
    Set rngTodo = ws.UsedRange
    With rngTodo
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
    End With
    
    ' Columna A (Letra): rojo
    Dim rngLetra As Range
    Set rngLetra = ws.Range(ws.Cells(2, 1), ws.Cells(numFilas + 1, 1))
    rngLetra.Font.Color = RGB(255, 0, 0)
    rngLetra.Font.Bold = True
    
    ' Columna C (Marca): color según marca
    Dim rngMarca As Range
    Set rngMarca = ws.Range(ws.Cells(2, 3), ws.Cells(numFilas + 1, 3))
    Select Case marcaNombre
        Case "ALAMO"
            rngMarca.Font.Color = RGB(0, 112, 192)
        Case "ENTERPRISE"
            rngMarca.Font.Color = RGB(0, 0, 0)
        Case "NATIONAL"
            rngMarca.Font.Color = RGB(0, 176, 80)
    End Select
    rngMarca.Font.Bold = True
    
    ' Columna D (**): rojo
    Dim rngD As Range
    Set rngD = ws.Range(ws.Cells(2, 4), ws.Cells(numFilas + 1, 4))
    rngD.Font.Color = RGB(255, 0, 0)
    rngD.Font.Bold = True
    
    ' Columna E (Nombre): color según marca
    Dim rngNombre As Range
    Set rngNombre = ws.Range(ws.Cells(2, 5), ws.Cells(numFilas + 1, 5))
    Select Case marcaNombre
        Case "ALAMO"
            rngNombre.Font.Color = RGB(0, 112, 192)
        Case "ENTERPRISE"
            rngNombre.Font.Color = RGB(0, 0, 0)
        Case "NATIONAL"
            rngNombre.Font.Color = RGB(0, 176, 80)
    End Select
    rngNombre.Font.Bold = True
    
    ' =========================================================
    ' BORDE SUPERIOR EN FILAS DE LETRA (columna A con valor)
    ' Aplica una línea horizontal arriba de cada grupo de letra
    ' =========================================================
    Dim j As Long
    Dim numCols As Long
    numCols = 13 ' Columnas A hasta M
    
    For i = 1 To numFilas
        filaActual = i + 1
        If ws.Cells(filaActual, 1).Value <> "" Then
            For j = 1 To numCols
                With ws.Cells(filaActual, j).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = RGB(0, 0, 0)
                End With
            Next j
        End If
    Next i
    ' =========================================================
    
End Sub

Private Sub QuickSortArray(ByRef arr As Variant, ByVal left As Long, ByVal right As Long, ByVal col As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant
    Dim temp As Variant
    Dim colIdx As Long
    i = left
    j = right
    pivot = arr((left + right) \ 2, col)
    While i <= j
        While arr(i, col) < pivot And i < right
            i = i + 1
        Wend
        While pivot < arr(j, col) And j > left
            j = j - 1
        Wend
        If i <= j Then
            For colIdx = LBound(arr, 2) To UBound(arr, 2)
                temp = arr(i, colIdx)
                arr(i, colIdx) = arr(j, colIdx)
                arr(j, colIdx) = temp
            Next colIdx
            i = i + 1
            j = j - 1
        End If
    Wend
    If left < j Then QuickSortArray arr, left, j, col
    If i < right Then QuickSortArray arr, i, right, col
End Sub
