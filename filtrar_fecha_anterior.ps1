param(
    [Parameter(Position = 0)]
    [string]$ArchivoOrigen = "Pruebas2026.xlsx",
    
    [Parameter(Position = 1)]
    [string]$ArchivoSalida = "",
    
    [Parameter(Position = 2)]
    [string]$FechaReferencia = "" # Si se especifica, usa esta fecha (ej: "2026-10-20")
)

# =========================================================================
# FILTRADO INTELIGENTE DE FECHA ANTERIOR (CORTE 9:00 AM)
# Antigravity - Sistema SAR
#
# Reglas de negocio:
# - Lunes: jala informacion de 3 dias antes (Viernes, Sabado y Domingo).
# - Martes a Viernes: jala informacion de 1 dia antes.
# - Sabado y Domingo: no jala nada (0 registros).
# - Para las filas que coinciden:
#   * Fecha Registro -> Fecha de hoy a las 09:00:00 (OADate 46314.375)
#   * Mes Registro   -> Mes actual (ej: 10)
#   * Hora Registro  -> 09:00:00
# =========================================================================

$ErrorActionPreference = "Stop"
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptDir) { $scriptDir = "C:\Users\PC\Desktop\test" }

$sourcePath = if ([System.IO.Path]::IsPathRooted($ArchivoOrigen)) { $ArchivoOrigen } else { Join-Path $scriptDir $ArchivoOrigen }

if (-not (Test-Path $sourcePath)) {
    Write-Host "[ERROR] No se encontro el archivo de origen: $sourcePath" -ForegroundColor Red
    exit 1
}

$refDate = [DateTime]::Today
if ($FechaReferencia) {
    try {
        $refDate = [DateTime]::Parse($FechaReferencia)
    } catch {
        Write-Host "[WARN] No se pudo interpretar FechaReferencia '$FechaReferencia'. Usando fecha actual." -ForegroundColor Yellow
    }
}

if (-not $ArchivoSalida -or $ArchivoSalida -eq "Pruebas2026_Filtrado_Lunes.xlsx") {
    $diaNombre = switch ($refDate.DayOfWeek) {
        ([DayOfWeek]::Monday)    { "Lunes" }
        ([DayOfWeek]::Tuesday)   { "Martes" }
        ([DayOfWeek]::Wednesday) { "Miercoles" }
        ([DayOfWeek]::Thursday)  { "Jueves" }
        ([DayOfWeek]::Friday)    { "Viernes" }
        ([DayOfWeek]::Saturday)  { "Sabado" }
        ([DayOfWeek]::Sunday)    { "Domingo" }
    }
    $ArchivoSalida = "Pruebas2026_Filtrado_$diaNombre.xlsx"
}
$destPath = if ([System.IO.Path]::IsPathRooted($ArchivoSalida)) { $ArchivoSalida } else { Join-Path $scriptDir $ArchivoSalida }

$dayOfWeek = $refDate.DayOfWeek
$targetDates = @()

Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " FILTRADO DE FECHA ANTERIOR (CORTE 9:00 AM)" -ForegroundColor Yellow
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " Fecha de referencia: $($refDate.ToString('dd/MM/yyyy')) ($dayOfWeek)" -ForegroundColor White

if ($dayOfWeek -eq [DayOfWeek]::Monday) {
    # Lunes: Domingo (-1), Sabado (-2), Viernes (-3)
    $dDom = $refDate.AddDays(-1).ToString("yyyy-MM-dd")
    $dSab = $refDate.AddDays(-2).ToString("yyyy-MM-dd")
    $dVie = $refDate.AddDays(-3).ToString("yyyy-MM-dd")
    $targetDates = @($dDom, $dSab, $dVie)
    Write-Host " -> LUNES detectado: Se filtran 3 dias anteriores:" -ForegroundColor Green
    Write-Host "    * Viernes: $dVie" -ForegroundColor Green
    Write-Host "    * Sabado:  $dSab" -ForegroundColor Green
    Write-Host "    * Domingo: $dDom" -ForegroundColor Green
} elseif ($dayOfWeek -ge [DayOfWeek]::Tuesday -and $dayOfWeek -le [DayOfWeek]::Friday) {
    # Martes a Viernes: 1 dia antes
    $dAyer = $refDate.AddDays(-1).ToString("yyyy-MM-dd")
    $targetDates = @($dAyer)
    Write-Host " -> Dia laboral detectado ($dayOfWeek): Se filtra 1 dia anterior:" -ForegroundColor Green
    Write-Host "    * Ayer: $dAyer" -ForegroundColor Green
} else {
    Write-Host " -> Fin de semana ($dayOfWeek): Segun las reglas de negocio, Sabado y Domingo no jalan informacion." -ForegroundColor Yellow
    Write-Host "    No se procesara ninguna fila." -ForegroundColor Yellow
    exit 0
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Abriendo archivo origen $sourcePath..." -ForegroundColor Gray
    $wbSrc = $excel.Workbooks.Open($sourcePath, $null, $true)
    $wsSrc = $wbSrc.Sheets.Item(1)
    
    $totalRows = $wsSrc.UsedRange.Rows.Count
    $totalCols = $wsSrc.UsedRange.Columns.Count
    Write-Host "Archivo cargado: $totalRows filas, $totalCols columnas." -ForegroundColor Gray
    
    # Leer encabezados (fila 1)
    $headersRange = $wsSrc.Range($wsSrc.Cells.Item(1, 1), $wsSrc.Cells.Item(1, $totalCols)).Value2
    $headers = @()
    for ($c = 1; $c -le $totalCols; $c++) {
        $headers += "$($headersRange[1, $c])"
    }
    
    # Identificar indices de columnas clave
    $idxFecha = 2
    $idxMes = 3
    $idxHora = 4
    for ($c = 1; $c -le $headers.Count; $c++) {
        $h = $headers[$c - 1].Trim()
        if ($h -eq "Fecha Registro" -and $c -le 5) { $idxFecha = $c }
        elseif ($h -eq "Mes Registro" -and $c -le 5) { $idxMes = $c }
        elseif ($h -eq "Hora Registro" -and $c -le 5) { $idxHora = $c }
    }
    
    Write-Host "Columnas clave: Fecha Registro (Col $idxFecha), Mes Registro (Col $idxMes), Hora Registro (Col $idxHora)" -ForegroundColor Gray
    Write-Host "Identificando registros coincidentes..." -ForegroundColor Gray
    
    $datesRange = $wsSrc.Range("B2:B$totalRows").Value2
    $matchingRowIndices = New-Object System.Collections.Generic.List[int]
    
    for ($i = 1; $i -le $datesRange.GetLength(0); $i++) {
        $val = $datesRange[$i, 1]
        $dStr = ""
        if ($val -is [double] -or $val -is [int]) {
            $dt = [DateTime]::FromOADate([double]$val)
            $dStr = $dt.ToString("yyyy-MM-dd")
        } else {
            $raw = "$val".Trim()
            if ($raw.Length -ge 10) { $dStr = $raw.Substring(0, 10) }
        }
        
        if ($targetDates -contains $dStr) {
            $matchingRowIndices.Add($i + 1)
        }
    }
    
    Write-Host "Total filas que cumplen la condicion de fecha: $($matchingRowIndices.Count)" -ForegroundColor Green
    
    if ($matchingRowIndices.Count -eq 0) {
        Write-Host "[AVISO] No se encontraron registros para las fechas especificadas." -ForegroundColor Yellow
        $wbSrc.Close($false)
        return
    }
    
    # Valores de sustitucion requeridos:
    # 1. Fecha Registro -> Fecha de hoy a las 09:00:00 (OADate serial)
    $nuevaFechaObj = [DateTime]::new($refDate.Year, $refDate.Month, $refDate.Day, 9, 0, 0)
    $nuevaFechaOADate = $nuevaFechaObj.ToOADate() # ej: 46314.375
    $nuevaFechaTexto = $nuevaFechaObj.ToString("dd/MM/yyyy 09:00:00")
    
    # 2. Mes Registro -> Mes actual
    $nuevoMes = $refDate.Month
    
    # 3. Hora Registro -> 09:00:00
    $nuevaHoraTexto = "09:00:00"
    
    Write-Host "Preparando nuevo libro Excel con los registros actualizados..." -ForegroundColor Gray
    $wbDest = $excel.Workbooks.Add()
    $wsDest = $wbDest.Sheets.Item(1)
    $wsDest.Name = "Filtrado_Corte_9AM"
    
    $destRowsCount = $matchingRowIndices.Count + 1
    $outArray = [Array]::CreateInstance([object], @($destRowsCount, $totalCols))
    
    # 1. Cabeceras
    for ($c = 0; $c -lt $totalCols; $c++) {
        $outArray.SetValue($headers[$c], 0, $c)
    }
    
    # 2. Cargar datos en bloques para maximo rendimiento
    $minRow = $matchingRowIndices[0]
    $maxRow = $matchingRowIndices[$matchingRowIndices.Count - 1]
    $chunkRange = $wsSrc.Range($wsSrc.Cells.Item($minRow, 1), $wsSrc.Cells.Item($maxRow, $totalCols)).Value2
    
    for ($k = 0; $k -lt $matchingRowIndices.Count; $k++) {
        $realRow = $matchingRowIndices[$k]
        $chunkRowIdx = $realRow - $minRow + 1
        $outRowIdx = $k + 1
        
        for ($c = 1; $c -le $totalCols; $c++) {
            $cellVal = $chunkRange[$chunkRowIdx, $c]
            
            if ($c -eq $idxFecha) {
                # Guardamos como serial de fecha OADate para compatibilidad perfecta con Excel
                $outArray.SetValue($nuevaFechaOADate, $outRowIdx, $c - 1)
            } elseif ($c -eq $idxMes) {
                $outArray.SetValue($nuevoMes, $outRowIdx, $c - 1)
            } elseif ($c -eq $idxHora) {
                $outArray.SetValue($nuevaHoraTexto, $outRowIdx, $c - 1)
            } else {
                $outArray.SetValue($cellVal, $outRowIdx, $c - 1)
            }
        }
    }
    
    Write-Host "Transfiriendo $($destRowsCount) filas y formateando columnas..." -ForegroundColor Gray
    $destRange = $wsDest.Range($wsDest.Cells.Item(1, 1), $wsDest.Cells.Item($destRowsCount, $totalCols))
    $destRange.Value2 = $outArray
    
    # Formato de columna Fecha Registro (Col B)
    $colFechaRange = $wsDest.Range($wsDest.Cells.Item(2, $idxFecha), $wsDest.Cells.Item($destRowsCount, $idxFecha))
    $colFechaRange.NumberFormat = "yyyy-mm-dd hh:mm:ss"
    
    # Formato de columna Mes Registro (Col C)
    $colMesRange = $wsDest.Range($wsDest.Cells.Item(2, $idxMes), $wsDest.Cells.Item($destRowsCount, $idxMes))
    $colMesRange.NumberFormat = "0"
    
    # Formato de columna Hora Registro (Col D)
    $colHoraRange = $wsDest.Range($wsDest.Cells.Item(2, $idxHora), $wsDest.Cells.Item($destRowsCount, $idxHora))
    $colHoraRange.NumberFormat = "hh:mm:ss"
    
    # Estilo visual de cabecera
    $hdrRange = $wsDest.Range($wsDest.Cells.Item(1, 1), $wsDest.Cells.Item(1, $totalCols))
    $hdrRange.Font.Bold = $true
    $hdrRange.Interior.Color = 0x2563EB # Azul elegante
    $hdrRange.Font.Color = 0xFFFFFF
    
    if (Test-Path $destPath) { Remove-Item $destPath -Force }
    $wbDest.SaveAs($destPath, 51)
    $wbDest.Close($false)
    $wbSrc.Close($false)
    
    Write-Host "========================================================================" -ForegroundColor Green
    Write-Host " [EXITO] Archivo generado: $destPath" -ForegroundColor Green
    Write-Host " Registros exportados: $($matchingRowIndices.Count)" -ForegroundColor Green
    Write-Host " Fecha Registro asignada: $nuevaFechaTexto" -ForegroundColor Green
    Write-Host " Mes Registro asignado:   $nuevoMes" -ForegroundColor Green
    Write-Host " Hora Registro asignada:  $nuevaHoraTexto" -ForegroundColor Green
    Write-Host "========================================================================" -ForegroundColor Green
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

$sw.Stop()
Write-Host "Proceso completado en $($sw.ElapsedMilliseconds) ms" -ForegroundColor Cyan
