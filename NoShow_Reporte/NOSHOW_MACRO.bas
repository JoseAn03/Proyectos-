'---------------------------------------------------------------------------
'  MACRO: GenerarResManifest
'  DESCRIPCIÓN: Lee la hoja "Sheet" del archivo fuente y genera una nueva
'               hoja formateada con las columnas:
'               A=Letra | B=N.º reserva | C=Marca | D=** | E=Nombre |
'               F=Clase | G=**MAIN VIP** | H=Fecha recogida |
'               I=Fecha entrega | J=Ubic. rec. | K=Ubic. ent. |
'               L=Tarifa Diaria | M=Agencia de recom.
'
'  CÓMO USAR:
'    1. Abre el archivo Excel fuente (ResManifest_.xlsx)
'    2. Presiona Alt + F11  →  se abre el Editor de VBA
'    3. En el menú: Insertar → Módulo
'    4. Pega TODO este código en el módulo en blanco
'    5. Cierra el editor (Alt + F4 o simplemente la X)
'    6. Presiona Alt + F8, selecciona "GenerarResManifest" y haz clic en Ejecutar
'    7. Se creará una nueva hoja llamada "Manifiesto" con el resultado
'---------------------------------------------------------------------------

Option Explicit

Sub GenerarResManifest()

    '--- Variables de color (RGB() no puede usarse en Const en VBA) ---
    Dim COL_ROJO      As Long: COL_ROJO      = RGB(255, 0, 0)
    Dim COL_AZUL      As Long: COL_AZUL      = RGB(0, 112, 192)   ' Alamo
    Dim COL_NEGRO     As Long: COL_NEGRO     = RGB(0, 0, 0)        ' Enterprise
    Dim COL_VERDE     As Long: COL_VERDE     = RGB(0, 176, 80)     ' National
    Dim COL_BLANCO    As Long: COL_BLANCO    = RGB(255, 255, 255)
    Dim COL_HEADER_BG As Long: COL_HEADER_BG = RGB(31, 78, 121)   ' Azul oscuro cabecera

    Dim wsOrigen  As Worksheet
    Dim wsDest    As Worksheet
    Dim wbThis    As Workbook

    Set wbThis = ThisWorkbook

    '--- Verificar hoja origen ---
    On Error Resume Next
    Set wsOrigen = wbThis.Sheets("Sheet")
    On Error GoTo 0
    If wsOrigen Is Nothing Then
        MsgBox "No se encontró la hoja 'Sheet'. Verifica el nombre de la hoja fuente.", vbCritical
        Exit Sub
    End If

    '--- Eliminar hoja destino si ya existe ---
    Application.DisplayAlerts = False
    On Error Resume Next
    wbThis.Sheets("Manifiesto").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    '--- Crear hoja destino ---
    Set wsDest = wbThis.Sheets.Add(After:=wbThis.Sheets(wbThis.Sheets.Count))
    wsDest.Name = "Manifiesto"

    '==========================================================
    '  PASO 1 – Leer y almacenar datos del origen
    '==========================================================
    Dim ultimaFila As Long
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 2).End(xlUp).Row  ' col B = N.º reserva

    ' Índices columnas ORIGEN (1-based)
    ' Col1=A flag (~/*), Col2=B reserva, Col3=C (+), Col4=D nombre
    ' Col6=F clase, Col8=H f.recogida, Col9=I f.entrega
    ' Col10=J ubic.rec, Col11=K ubic.ent, Col13=M tarifa, Col18=R agencia

    Dim nFilas As Long
    nFilas = ultimaFila - 1  ' sin cabecera

    ' Arrays para guardar datos
    ReDim aFlag(1 To nFilas)        As String
    ReDim aReserva(1 To nFilas)     As Variant
    ReDim aPlus(1 To nFilas)        As String
    ReDim aNombre(1 To nFilas)      As String
    ReDim aClase(1 To nFilas)       As String
    ReDim aFRec(1 To nFilas)        As Variant
    ReDim aFEnt(1 To nFilas)        As Variant
    ReDim aUbicRec(1 To nFilas)     As String
    ReDim aUbicEnt(1 To nFilas)     As String
    ReDim aTarifa(1 To nFilas)      As Variant
    ReDim aAgencia(1 To nFilas)     As String

    Dim i As Long
    Dim srcRow As Long
    srcRow = 2  ' empieza en fila 2 (salta cabecera)

    For i = 1 To nFilas
        aFlag(i)    = Trim(CStr(IIf(IsNull(wsOrigen.Cells(srcRow, 1).Value), "", wsOrigen.Cells(srcRow, 1).Value)))
        aReserva(i) = wsOrigen.Cells(srcRow, 2).Value
        aPlus(i)    = Trim(CStr(IIf(IsNull(wsOrigen.Cells(srcRow, 3).Value), "", wsOrigen.Cells(srcRow, 3).Value)))
        aNombre(i)  = Trim(CStr(IIf(IsNull(wsOrigen.Cells(srcRow, 4).Value), "", wsOrigen.Cells(srcRow, 4).Value)))
        aClase(i)   = Trim(CStr(IIf(IsNull(wsOrigen.Cells(srcRow, 6).Value), "", wsOrigen.Cells(srcRow, 6).Value)))
        aFRec(i)    = wsOrigen.Cells(srcRow, 8).Value
        aFEnt(i)    = wsOrigen.Cells(srcRow, 9).Value
        aUbicRec(i) = Trim(CStr(IIf(IsNull(wsOrigen.Cells(srcRow, 10).Value), "", wsOrigen.Cells(srcRow, 10).Value)))
        aUbicEnt(i) = Trim(CStr(IIf(IsNull(wsOrigen.Cells(srcRow, 11).Value), "", wsOrigen.Cells(srcRow, 11).Value)))
        aTarifa(i)  = wsOrigen.Cells(srcRow, 13).Value
        aAgencia(i) = Trim(CStr(IIf(IsNull(wsOrigen.Cells(srcRow, 18).Value), "", wsOrigen.Cells(srcRow, 18).Value)))
        srcRow = srcRow + 1
    Next i

    '==========================================================
    '  PASO 2 – Ordenar por Nombre (A-Z) con BubbleSort simple
    '==========================================================
    Dim j As Long
    Dim tmpStr As String, tmpVar As Variant

    For i = 1 To nFilas - 1
        For j = i + 1 To nFilas
            If UCase(aNombre(i)) > UCase(aNombre(j)) Then
                ' Intercambiar todos los campos
                tmpStr = aFlag(i):    aFlag(i) = aFlag(j):       aFlag(j) = tmpStr
                tmpVar = aReserva(i): aReserva(i) = aReserva(j): aReserva(j) = tmpVar
                tmpStr = aPlus(i):    aPlus(i) = aPlus(j):       aPlus(j) = tmpStr
                tmpStr = aNombre(i):  aNombre(i) = aNombre(j):   aNombre(j) = tmpStr
                tmpStr = aClase(i):   aClase(i) = aClase(j):     aClase(j) = tmpStr
                tmpVar = aFRec(i):    aFRec(i) = aFRec(j):       aFRec(j) = tmpVar
                tmpVar = aFEnt(i):    aFEnt(i) = aFEnt(j):       aFEnt(j) = tmpVar
                tmpStr = aUbicRec(i): aUbicRec(i) = aUbicRec(j): aUbicRec(j) = tmpStr
                tmpStr = aUbicEnt(i): aUbicEnt(i) = aUbicEnt(j): aUbicEnt(j) = tmpStr
                tmpVar = aTarifa(i):  aTarifa(i) = aTarifa(j):   aTarifa(j) = tmpVar
                tmpStr = aAgencia(i): aAgencia(i) = aAgencia(j): aAgencia(j) = tmpStr
            End If
        Next j
    Next i

    '==========================================================
    '  PASO 3 – Escribir cabecera en hoja destino
    '==========================================================
    Dim headers(1 To 13) As String
    headers(1)  = "Letra"
    headers(2)  = "N." & Chr(186) & " reserva"
    headers(3)  = "Marca"
    headers(4)  = "**"
    headers(5)  = "Nombre"
    headers(6)  = "Clase"
    headers(7)  = "**MAIN VIP**"
    headers(8)  = "Fecha recogida"
    headers(9)  = "Fecha entrega"
    headers(10) = "Ubic. rec."
    headers(11) = "Ubic. ent."
    headers(12) = "Tarifa Diaria"
    headers(13) = "Agencia de recom."

    Dim k As Long
    For k = 1 To 13
        With wsDest.Cells(1, k)
            .Value = headers(k)
            .Font.Name = "Calibri"
            .Font.Bold = True
            .Font.Color = COL_BLANCO
            .Font.Size = 11
            .Interior.Color = COL_HEADER_BG
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    Next k
    wsDest.Rows(1).RowHeight = 30

    '==========================================================
    '  PASO 4 – Escribir filas de datos con formato
    '==========================================================
    Dim destRow   As Long
    Dim letraAnt  As String
    Dim letraAct  As String
    Dim marcaStr  As String
    Dim colMarca  As Long
    Dim colNombre As Long
    Dim colD      As String
    Dim colG      As String
    Dim hasFl     As Boolean
    Dim hasPlus   As Boolean
    Dim letraVal  As String

    destRow  = 2
    letraAnt = ""

    For i = 1 To nFilas

        '--- Determinar Marca por Ubic. rec. ---
        Dim ubic As String
        ubic = UCase(aUbicRec(i))

        If InStr(ubic, "E71") > 0 Or InStr(ubic, "T71") > 0 Then
            marcaStr  = "ALAMO/"
            colMarca  = COL_AZUL
            colNombre = COL_AZUL
        ElseIf InStr(ubic, "C61") > 0 Or InStr(ubic, "T61") > 0 Then
            marcaStr  = "ENTERPRISE/"
            colMarca  = COL_NEGRO
            colNombre = COL_NEGRO
        ElseIf InStr(ubic, "E01") > 0 Or InStr(ubic, "T01") > 0 Then
            marcaStr  = "NATIONAL/"
            colMarca  = COL_VERDE
            colNombre = COL_VERDE
        Else
            marcaStr  = ""
            colMarca  = COL_NEGRO
            colNombre = COL_NEGRO
        End If

        '--- Columna A: Letra del grupo (solo A-Z, ignorar números y caracteres) ---
        Dim primerChar As String
        primerChar = ""
        If Len(aNombre(i)) > 0 Then
            primerChar = UCase(Left(aNombre(i), 1))
        End If

        ' Solo aceptar letras del abecedario A-Z
        If primerChar >= "A" And primerChar <= "Z" Then
            letraAct = primerChar
        Else
            letraAct = ""   ' No es letra válida → no se asigna grupo
        End If

        If letraAct <> "" And letraAct <> letraAnt Then
            letraVal = letraAct
            letraAnt = letraAct
        Else
            letraVal = ""
        End If

        '--- Columna D: Indicadores FL / Marca ---
        hasFl   = (InStr(aFlag(i), "~") > 0 Or InStr(aFlag(i), "*") > 0)
        hasPlus = (InStr(aPlus(i), "+") > 0)

        Dim brandLetter As String
        brandLetter = ""
        If hasPlus Then
            If marcaStr = "ALAMO/"      Then brandLetter = "I"
            If marcaStr = "ENTERPRISE/" Then brandLetter = "P"
            If marcaStr = "NATIONAL/"   Then brandLetter = "E"
        End If

        If hasFl And brandLetter <> "" Then
            colD = ChrW(11088) & brandLetter & "-FL" & ChrW(11088)  ' ⭐X-FL⭐
        ElseIf hasFl Then
            colD = ChrW(11088) & "FL" & ChrW(11088)                 ' ⭐FL⭐
        ElseIf brandLetter <> "" Then
            colD = brandLetter
        Else
            colD = ""
        End If

        '--- Columna G: MAIN VIP / EXPEDIA ---
        Dim isVip     As Boolean
        Dim isExpedia As Boolean
        isVip     = (Right(aUbicRec(i), 3) = "E71" Or Right(aUbicRec(i), 3) = "E01" Or Right(aUbicRec(i), 3) = "C61")
        isExpedia = (InStr(aAgencia(i), "11617270") > 0)

        If isVip And isExpedia Then
            colG = "* MAIN VIP* * EXPEDIA*"
        ElseIf isVip Then
            colG = "* MAIN VIP*"
        ElseIf isExpedia Then
            colG = "* EXPEDIA*"
        Else
            colG = ""
        End If

        '--- Formatear fechas ---
        Dim strFRec As String, strFEnt As String
        If IsDate(aFRec(i)) Then
            strFRec = Format(CDate(aFRec(i)), "DD/MM/YYYY HH:MM")
        Else
            strFRec = CStr(aFRec(i))
        End If
        If IsDate(aFEnt(i)) Then
            strFEnt = Format(CDate(aFEnt(i)), "DD/MM/YYYY HH:MM")
        Else
            strFEnt = CStr(aFEnt(i))
        End If

        '--- Escribir valores ---
        wsDest.Cells(destRow, 1).Value  = letraVal
        wsDest.Cells(destRow, 2).Value  = aReserva(i)
        wsDest.Cells(destRow, 3).Value  = marcaStr
        wsDest.Cells(destRow, 4).Value  = colD
        wsDest.Cells(destRow, 5).Value  = aNombre(i)
        wsDest.Cells(destRow, 6).Value  = aClase(i)
        wsDest.Cells(destRow, 7).Value  = colG
        wsDest.Cells(destRow, 8).Value  = strFRec
        wsDest.Cells(destRow, 9).Value  = strFEnt
        wsDest.Cells(destRow, 10).Value = aUbicRec(i)
        wsDest.Cells(destRow, 11).Value = aUbicEnt(i)
        wsDest.Cells(destRow, 12).Value = aTarifa(i)
        wsDest.Cells(destRow, 13).Value = aAgencia(i)

        '--- Aplicar formato base a toda la fila (Calibri, negrita, negro) ---
        With wsDest.Range(wsDest.Cells(destRow, 1), wsDest.Cells(destRow, 13))
            .Font.Name  = "Calibri"
            .Font.Bold  = True
            .Font.Color = COL_NEGRO
            .HorizontalAlignment = xlCenter
            .VerticalAlignment   = xlCenter
            .WrapText = True
        End With

        '--- Col A: Letra en ROJO ---
        If letraVal <> "" Then
            With wsDest.Cells(destRow, 1).Font
                .Color = COL_ROJO
                .Size  = 12
            End With
            '--- Borde superior en toda la fila cuando cambia la letra ---
            With wsDest.Range(wsDest.Cells(destRow, 1), wsDest.Cells(destRow, 13)).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight    = xlMedium
                .Color     = COL_NEGRO
            End With
        End If

        '--- Col C: Marca con color de marca ---
        With wsDest.Cells(destRow, 3).Font
            .Color = colMarca
        End With

        '--- Col D: indicador en ROJO ---
        If colD <> "" Then
            wsDest.Cells(destRow, 4).Font.Color = COL_ROJO
        End If

        '--- Col E: Nombre con color de marca ---
        wsDest.Cells(destRow, 5).Font.Color = colNombre

        wsDest.Rows(destRow).RowHeight = 18
        destRow = destRow + 1
    Next i

    '==========================================================
    '  PASO 5 – Anchos de columna
    '==========================================================
    wsDest.Columns(1).ColumnWidth  = 8    ' Letra
    wsDest.Columns(2).ColumnWidth  = 15   ' N.º reserva
    wsDest.Columns(3).ColumnWidth  = 14   ' Marca
    wsDest.Columns(4).ColumnWidth  = 12   ' **
    wsDest.Columns(5).ColumnWidth  = 30   ' Nombre
    wsDest.Columns(6).ColumnWidth  = 8    ' Clase
    wsDest.Columns(7).ColumnWidth  = 22   ' **MAIN VIP**
    wsDest.Columns(8).ColumnWidth  = 20   ' Fecha recogida
    wsDest.Columns(9).ColumnWidth  = 20   ' Fecha entrega
    wsDest.Columns(10).ColumnWidth = 12   ' Ubic. rec.
    wsDest.Columns(11).ColumnWidth = 12   ' Ubic. ent.
    wsDest.Columns(12).ColumnWidth = 12   ' Tarifa Diaria
    wsDest.Columns(13).ColumnWidth = 40   ' Agencia

    '--- Inmovilizar cabecera ---
    wsDest.Activate
    wsDest.Cells(2, 1).Select
    ActiveWindow.FreezePanes = True

    MsgBox "¡Listo! Se generó la hoja 'Manifiesto' con " & (destRow - 2) & " filas de datos.", vbInformation

End Sub
