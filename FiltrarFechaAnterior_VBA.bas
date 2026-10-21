Attribute VB_Name = "ModFiltroFechaAnterior"
' =========================================================================
' MODULO VBA: FILTRADO INTELIGENTE DE FECHA ANTERIOR (CORTE 9:00 AM)
' Sistema de Gestión y Filtrado de Servicios (SAR)
'
' REGLAS DE NEGOCIO IMPLEMENTADAS:
' 1. Si es LUNES:
'    - Jala información de 3 días antes: DOMINGO, SÁBADO y VIERNES.
' 2. Si es MARTES A VIERNES:
'    - Jala información de 1 día antes (Ayer).
' 3. Si es SÁBADO o DOMINGO:
'    - No jala nada (0 registros).
'
' VALORES ACTUALIZADOS EN LOS REGISTROS COINCIDENTES:
' - Columna "Fecha Registro": Fecha de hoy a las 09:00:00
' - Columna "Mes Registro":   Mes actual (ej. 10)
' - Columna "Hora Registro":  09:00:00
' =========================================================================

Public Sub EjecutarFiltroFechaAnterior()
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim totalRows As Long, totalCols As Long
    Dim r As Long, c As Long, outRow As Long
    Dim colFecha As Long, colMes As Long, colHora As Long
    Dim todayDate As Date, dayOfWeekNum As Integer
    Dim strDom As String, strSab As String, strVie As String, strAyer As String
    Dim rawDateVal As Variant, strDate As String
    Dim isMatch As Boolean, matchCount As Long
    Dim newFechaSerial As Double, newMesVal As Integer, newHoraVal As String
    Dim targetDates As Collection
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set wsSrc = ActiveSheet
    totalRows = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    totalCols = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    
    If totalRows < 2 Then
        MsgBox "La hoja actual no contiene datos suficientes.", vbExclamation, "Aviso"
        GoTo CleanExit
    End If
    
    ' Identificar columnas Fecha Registro, Mes Registro y Hora Registro
    colFecha = 2
    colMes = 3
    colHora = 4
    For c = 1 To Application.Min(totalCols, 10)
        Dim h As String
        h = Trim(CStr(wsSrc.Cells(1, c).Value))
        If InStr(1, h, "Fecha Registro", vbTextCompare) > 0 And colFecha = 2 Then colFecha = c
        If InStr(1, h, "Mes Registro", vbTextCompare) > 0 And colMes = 3 Then colMes = c
        If InStr(1, h, "Hora Registro", vbTextCompare) > 0 And colHora = 4 Then colHora = c
    Next c
    
    ' Fecha de evaluación: Toma siempre la fecha actual del sistema (Hoy) automáticamente
    todayDate = Date
    dayOfWeekNum = Weekday(todayDate, vbMonday) ' 1=Lunes, 2=Martes, ..., 6=Sábado, 7=Domingo
    
    Set targetDates = New Collection
    
    ' REGLAS DE NEGOCIO
    If dayOfWeekNum = 1 Then
        ' LUNES: Jala 3 días antes -> Domingo (-1), Sábado (-2) y Viernes (-3)
        strDom = Format(todayDate - 1, "yyyy-mm-dd")
        strSab = Format(todayDate - 2, "yyyy-mm-dd")
        strVie = Format(todayDate - 3, "yyyy-mm-dd")
        targetDates.Add strDom, strDom
        targetDates.Add strSab, strSab
        targetDates.Add strVie, strVie
    ElseIf dayOfWeekNum >= 2 And dayOfWeekNum <= 5 Then
        ' MARTES A VIERNES: Jala 1 día antes
        strAyer = Format(todayDate - 1, "yyyy-mm-dd")
        targetDates.Add strAyer, strAyer
    Else
        ' SÁBADO (6) O DOMINGO (7): No jala nada
        Application.ScreenUpdating = True
        MsgBox "Hoy es fin de semana (" & Format(todayDate, "dddd") & ")." & vbCrLf & _
               "Según las reglas de negocio, sábado y domingo no se jala información.", vbInformation, "Regla de Negocio"
        GoTo CleanExit
    End If
    
    ' Valores a asignar solicitados
    newFechaSerial = CDbl(todayDate) + TimeSerial(9, 0, 0)
    newMesVal = Month(todayDate)
    newHoraVal = "09:00:00"
    
    ' Crear o limpiar hoja de destino "Filtrado_Corte_9AM"
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("Filtrado_Corte_9AM")
    On Error GoTo ErrorHandler
    
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=wsSrc)
        wsDest.Name = "Filtrado_Corte_9AM"
    Else
        wsDest.Cells.Clear
    End If
    
    ' Copiar cabecera
    wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, totalCols)).Copy wsDest.Cells(1, 1)
    
    outRow = 2
    matchCount = 0
    
    ' Recorrer filas y filtrar
    For r = 2 To totalRows
        rawDateVal = wsSrc.Cells(r, colFecha).Value
        strDate = ""
        
        If IsDate(rawDateVal) Then
            strDate = Format(CDate(rawDateVal), "yyyy-mm-dd")
        ElseIf IsNumeric(rawDateVal) And rawDateVal > 30000 Then
            strDate = Format(CDate(rawDateVal), "yyyy-mm-dd")
        Else
            strDate = Left(Trim(CStr(rawDateVal)), 10)
        End If
        
        ' Verificar si coincide con alguna fecha objetivo
        isMatch = False
        On Error Resume Next
        isMatch = Not IsEmpty(targetDates(strDate))
        On Error GoTo ErrorHandler
        
        If isMatch Then
            matchCount = matchCount + 1
            wsSrc.Range(wsSrc.Cells(r, 1), wsSrc.Cells(r, totalCols)).Copy wsDest.Cells(outRow, 1)
            
            ' Asignar valores solicitados
            wsDest.Cells(outRow, colFecha).Value = newFechaSerial
            wsDest.Cells(outRow, colMes).Value = newMesVal
            wsDest.Cells(outRow, colHora).Value = newHoraVal
            
            outRow = outRow + 1
        End If
    Next r
    
    ' Formatear columnas en hoja destino
    wsDest.Columns(colFecha).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    wsDest.Columns(colMes).NumberFormat = "0"
    wsDest.Columns(colHora).NumberFormat = "hh:mm:ss"
    
    ' Estilo de cabecera
    With wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(1, totalCols))
        .Font.Bold = True
        .Interior.Color = RGB(37, 99, 235) ' Azul Corporativo
        .Font.Color = RGB(255, 255, 255)
    End With
    
    wsDest.Columns.AutoFit
    wsDest.Activate
    
    Application.ScreenUpdating = True
    MsgBox "¡Filtro de Fecha Anterior completado con éxito!" & vbCrLf & vbCrLf & _
           "• Día evaluado: " & Format(todayDate, "dddd dd/mm/yyyy") & vbCrLf & _
           "• Total registros exportados: " & Format(matchCount, "#,##0") & vbCrLf & _
           "• Fecha Registro asignada: " & Format(todayDate, "dd/mm/yyyy") & " 09:00:00" & vbCrLf & _
           "• Mes Registro asignado: " & newMesVal & vbCrLf & _
           "• Hora Registro asignada: " & newHoraVal, vbInformation, "SAR - Corte 9:00 AM"

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Ocurrió un error en la macro: " & Err.Description, vbCritical, "Error"
End Sub
