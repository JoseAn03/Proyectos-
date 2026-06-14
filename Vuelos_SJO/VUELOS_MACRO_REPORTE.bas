'----------------------------------------------------------------------------------
' MACRO: GenerarReporteVuelos
' Lee los vuelos de la hoja activa y genera el reporte en hoja "Reporte"
' Columnas: A=Hora | B=N°Vuelo | C=Destino | D=Aerolínea
' USO: Alt+F8 → GenerarReporteVuelos → Ejecutar
'----------------------------------------------------------------------------------

Sub GenerarReporteVuelos()

    Dim wsDatos As Worksheet
    Set wsDatos = ActiveSheet

    Dim excluirDestinos(13) As String
    excluirDestinos(0)  = "cobano"
    excluirDestinos(1)  = "tamarindo"
    excluirDestinos(2)  = "puerto jimenez"
    excluirDestinos(3)  = "limon"
    excluirDestinos(4)  = "golfito"
    excluirDestinos(5)  = "liberia"
    excluirDestinos(6)  = "quepos"
    excluirDestinos(7)  = "drake bay"
    excluirDestinos(8)  = "nosara"
    excluirDestinos(9)  = "la fortuna"
    excluirDestinos(10) = "bocas del toro"
    excluirDestinos(11) = "managua"
    excluirDestinos(12) = "bocas"
    excluirDestinos(13) = "tortuguero"

    Dim excluirAerolineas(5) As String
    excluirAerolineas(0) = "sansa"
    excluirAerolineas(1) = "latam"
    excluirAerolineas(2) = "getjet"
    excluirAerolineas(3) = "netjets"
    excluirAerolineas(4) = "viva"
    excluirAerolineas(5) = "mas air"

    Dim arrHora(300)      As String
    Dim arrAerolinea(300) As String
    Dim arrDestino(300)   As String
    Dim totalVuelos       As Integer
    totalVuelos = 0

    Dim ultimaFila As Long
    ultimaFila = wsDatos.Cells(wsDatos.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim j As Integer

    For i = 1 To ultimaFila
        Dim celVal       As Variant
        Dim celDestino   As String
        Dim celAerolinea As String
        Dim hora         As String
        Dim destino      As String
        Dim aerolinea    As String

        celVal       = wsDatos.Cells(i, 1).Value
        celDestino   = Trim(CStr(wsDatos.Cells(i, 3).Value))
        celAerolinea = Trim(CStr(wsDatos.Cells(i, 4).Value))

        hora = ""
        If IsEmpty(celVal) Or IsNull(celVal) Then GoTo SiguienteFila

        ' Manejar hora como Date, String o número
        On Error Resume Next
        If TypeName(celVal) = "Date" Then
            hora = Format(celVal, "hh:mm")
        ElseIf TypeName(celVal) = "String" Then
            If CStr(celVal) Like "##:##" Then
                hora = CStr(celVal)
            End If
        ElseIf IsNumeric(celVal) Then
            hora = Format(CDate(CDbl(celVal)), "hh:mm")
        End If
        On Error GoTo 0

        If hora = "" Or hora = "00:00" Then GoTo SiguienteFila

        ' Limpiar destino
        Dim posP As Integer
        posP = InStr(celDestino, "(")
        If posP > 1 Then
            destino = Trim(Left(celDestino, posP - 1))
        Else
            destino = celDestino
        End If
        destino = Replace(destino, Chr(160), "")
        destino = Trim(destino)

        ' Limpiar aerolínea
        aerolinea = Replace(celAerolinea, Chr(160), "")
        aerolinea = Trim(aerolinea)

        If aerolinea = "" Or aerolinea = "-" Then GoTo SiguienteFila

        Dim excluir As Boolean
        excluir = False
        For j = 0 To 5
            If InStr(1, LCase(aerolinea), excluirAerolineas(j), vbTextCompare) > 0 Then
                excluir = True : Exit For
            End If
        Next j
        If excluir Then GoTo SiguienteFila

        For j = 0 To 13
            If InStr(1, LCase(destino), excluirDestinos(j), vbTextCompare) > 0 Then
                excluir = True : Exit For
            End If
        Next j
        If excluir Then GoTo SiguienteFila

        arrHora(totalVuelos)      = hora
        arrAerolinea(totalVuelos) = aerolinea
        arrDestino(totalVuelos)   = destino
        totalVuelos = totalVuelos + 1

SiguienteFila:
    Next i

    If totalVuelos = 0 Then
        MsgBox "No se encontraron vuelos internacionales." & vbCrLf & _
               "Asegúrate de estar en la hoja con los vuelos pegados.", vbInformation
        Exit Sub
    End If

    ' Crear o limpiar hoja Reporte
    Dim wsReporte As Worksheet
    On Error Resume Next
    Set wsReporte = ThisWorkbook.Sheets("Reporte")
    On Error GoTo 0

    If wsReporte Is Nothing Then
        Set wsReporte = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsReporte.Name = "Reporte"
    Else
        wsReporte.Cells.UnMerge
        wsReporte.Cells.Clear
    End If

    wsReporte.Columns("A").ColumnWidth = 24
    wsReporte.Columns("B").ColumnWidth = 28
    wsReporte.Columns("C").ColumnWidth = 10

    ' Título
    wsReporte.Range("A1:C1").Merge
    With wsReporte.Range("A1")
        .Value = "Vuelos del Día - Aeropuerto Juan Santamaría"
        .Font.Name = "Arial" : .Font.Bold = True : .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    wsReporte.Range("A1:C1").Borders.LineStyle = xlContinuous
    wsReporte.Range("A1:C1").Borders.Color = RGB(170, 170, 170)

    ' Fecha
    wsReporte.Range("A2:C2").Merge
    With wsReporte.Range("A2")
        .Value = Format(Now(), "dddd d \d\e mmmm \d\e yyyy")
        .Font.Name = "Arial" : .Font.Bold = True : .Font.Size = 11
        .Font.Color = RGB(31, 78, 121)
        .Interior.Color = RGB(214, 228, 240)
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With
    wsReporte.Range("A2:C2").Borders.LineStyle = xlContinuous
    wsReporte.Range("A2:C2").Borders.Color = RGB(170, 170, 170)

    ' Encabezados
    Dim enc As Variant
    enc = Array("Aerolínea", "Destino", "Hora")
    Dim col As Integer
    For col = 1 To 3
        With wsReporte.Cells(3, col)
            .Value = enc(col - 1)
            .Font.Name = "Arial" : .Font.Bold = True : .Font.Size = 10
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
            .RowHeight = 18
        End With
    Next col
    wsReporte.Range("A3:C3").Borders.LineStyle = xlContinuous
    wsReporte.Range("A3:C3").Borders.Color = RGB(170, 170, 170)

    ' Bloques horarios
    Dim bloques(5)   As String
    Dim bloqueIni(5) As String
    Dim bloqueFin(5) As String
    bloques(0) = "05 a 09" : bloqueIni(0) = "05:00" : bloqueFin(0) = "08:59"
    bloques(1) = "09 a 12" : bloqueIni(1) = "09:00" : bloqueFin(1) = "11:59"
    bloques(2) = "12 a 15" : bloqueIni(2) = "12:00" : bloqueFin(2) = "14:59"
    bloques(3) = "15 a 18" : bloqueIni(3) = "15:00" : bloqueFin(3) = "17:59"
    bloques(4) = "18 a 21" : bloqueIni(4) = "18:00" : bloqueFin(4) = "20:59"
    bloques(5) = "21 a 05" : bloqueIni(5) = "21:00" : bloqueFin(5) = "23:59"

    Dim fila     As Integer
    Dim b        As Integer
    Dim v        As Integer
    Dim contFila As Integer
    fila = 4

    For b = 0 To 5
        Dim hayVuelos As Boolean
        hayVuelos = False
        For v = 0 To totalVuelos - 1
            If EnBloque(arrHora(v), bloqueIni(b), bloqueFin(b), b = 5) Then
                hayVuelos = True : Exit For
            End If
        Next v
        If Not hayVuelos Then GoTo SiguienteBloque

        wsReporte.Range(wsReporte.Cells(fila, 1), wsReporte.Cells(fila, 3)).Merge
        With wsReporte.Cells(fila, 1)
            .Value = bloques(b)
            .Font.Name = "Arial" : .Font.Bold = True : .Font.Size = 10
            .Font.Color = RGB(31, 78, 121)
            .Interior.Color = RGB(189, 215, 238)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
            .RowHeight = 16
        End With
        wsReporte.Range(wsReporte.Cells(fila, 1), wsReporte.Cells(fila, 3)).Borders.LineStyle = xlContinuous
        wsReporte.Range(wsReporte.Cells(fila, 1), wsReporte.Cells(fila, 3)).Borders.Color = RGB(170, 170, 170)
        fila = fila + 1

        contFila = 0
        For v = 0 To totalVuelos - 1
            If EnBloque(arrHora(v), bloqueIni(b), bloqueFin(b), b = 5) Then
                Dim cf As Long
                cf = IIf(contFila Mod 2 = 0, RGB(222, 234, 241), RGB(255, 255, 255))

                With wsReporte.Cells(fila, 1)
                    .Value = arrAerolinea(v)
                    .Font.Name = "Arial" : .Font.Size = 10
                    .Interior.Color = cf
                    .HorizontalAlignment = xlLeft : .VerticalAlignment = xlCenter
                    .RowHeight = 15
                End With
                With wsReporte.Cells(fila, 2)
                    .Value = arrDestino(v)
                    .Font.Name = "Arial" : .Font.Size = 10
                    .Interior.Color = cf
                    .HorizontalAlignment = xlLeft : .VerticalAlignment = xlCenter
                End With
                With wsReporte.Cells(fila, 3)
                    .Value = arrHora(v)
                    .Font.Name = "Arial" : .Font.Size = 10
                    .Interior.Color = cf
                    .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
                End With

                wsReporte.Range(wsReporte.Cells(fila, 1), wsReporte.Cells(fila, 3)).Borders.LineStyle = xlContinuous
                wsReporte.Range(wsReporte.Cells(fila, 1), wsReporte.Cells(fila, 3)).Borders.Color = RGB(170, 170, 170)

                fila = fila + 1
                contFila = contFila + 1
            End If
        Next v

SiguienteBloque:
    Next b

    wsReporte.Activate
    wsReporte.Range("A1").Select
    MsgBox "¡Listo! " & totalVuelos & " vuelos internacionales en la hoja 'Reporte'.", _
           vbInformation, "Vuelos del Día"

End Sub

Function EnBloque(hora As String, ini As String, fin As String, esUltimo As Boolean) As Boolean
    If esUltimo Then
        EnBloque = (hora >= ini)
    Else
        EnBloque = (hora >= ini And hora <= fin)
    End If
End Function
