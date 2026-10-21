param(
    [Parameter(Position = 0)]
    [string]$ArchivoBase = "Base a Cargar prueba.xlsx"
)

# Script de procesamiento y vinculacion de tablas Excel
# Antigravity - Sistema de Filtrado de Servicios con Criterios SAR

$ErrorActionPreference = "Stop"
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptDir) { $scriptDir = "c:\Users\ext_rleon\Desktop\test" }

Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "     SISTEMA DE GESTION Y VINCULACION DE TABLAS EXCEL (SAR)" -ForegroundColor Yellow
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "[1/5] Verificando archivos Excel en la carpeta..." -ForegroundColor Gray

# Determinar archivo base (admite nombre o ruta pasada por parametro)
$fileBaseSAR = $null
if ($ArchivoBase) {
    if ([System.IO.Path]::IsPathRooted($ArchivoBase)) {
        if (Test-Path $ArchivoBase) { $fileBaseSAR = $ArchivoBase }
    }
    else {
        $p = Join-Path $scriptDir $ArchivoBase
        if (Test-Path $p) { $fileBaseSAR = $p }
    }
}

if (-not $fileBaseSAR) {
    $candidatos = @(
        (Join-Path $scriptDir "Base a Cargar prueba.xlsx"),
        (Join-Path $scriptDir "Base SAR V1.xlsx")
    ) + (Get-ChildItem -Path $scriptDir -Filter "Base*.xlsx" | Select-Object -ExpandProperty FullName)
    foreach ($cand in $candidatos) {
        if ($cand -and (Test-Path $cand)) {
            $fileBaseSAR = $cand
            break
        }
    }
}

if (-not $fileBaseSAR -or -not (Test-Path $fileBaseSAR)) {
    Write-Host "[ERROR] No se encontro el archivo base '$ArchivoBase' ni ningun archivo base compatible." -ForegroundColor Red
    exit 1
}

$nombreBase = Split-Path $fileBaseSAR -Leaf
Write-Host "  -> Archivo base seleccionado: $nombreBase" -ForegroundColor Green

$fileTiempos = Get-ChildItem -Path $scriptDir -Filter "*Tiempo*Atenci*n*.xlsx" | Select-Object -First 1 -ExpandProperty FullName
if (-not $fileTiempos) { $fileTiempos = Join-Path $scriptDir "Tiempo de Atencion Servicios.xlsx" }
$fileAgrup = Join-Path $scriptDir "Agrupamiento de Servicios.xlsx"
$fileRegul = Join-Path $scriptDir "Reguladores.xlsx"
$fileRechaz = Join-Path $scriptDir "Servicios Rechazados.xlsx"
$fileEquipo = Join-Path $scriptDir "Equipo Contacto.xlsx"

Write-Host "[2/5] Iniciando motor de lectura de Excel..." -ForegroundColor Gray
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # 1. Cargar Tiempos de Atencion
    Write-Host "  -> Leyendo Tiempo de Atencion Servicios..." -ForegroundColor DarkGray
    $mapTiempos = @{}
    if ($fileTiempos -and (Test-Path $fileTiempos)) {
        $wb = $excel.Workbooks.Open($fileTiempos, $null, $true)
        $vT = $wb.Sheets.Item(1).UsedRange.Value2
        $wb.Close($false)
        if ($vT) {
            for ($r = 2; $r -le $vT.GetLength(0); $r++) {
                $s = "$($vT[$r, 1])".Trim()
                if ($s) { $mapTiempos[$s] = $vT[$r, 2] }
            }
        }
    }

    # 2. Cargar Agrupamiento
    Write-Host "  -> Leyendo Agrupamiento de Servicios..." -ForegroundColor DarkGray
    $mapAgrup = @{}
    if (Test-Path $fileAgrup) {
        $wb = $excel.Workbooks.Open($fileAgrup, $null, $true)
        $vA = $wb.Sheets.Item(1).UsedRange.Value2
        $wb.Close($false)
        if ($vA) {
            for ($r = 2; $r -le $vA.GetLength(0); $r++) {
                $s = "$($vA[$r, 1])".Trim()
                if ($s) { $mapAgrup[$s] = "$($vA[$r, 2])".Trim() }
            }
        }
    }

    # 3. Cargar Reguladores
    Write-Host "  -> Leyendo Reguladores..." -ForegroundColor DarkGray
    $mapRegul = @{}
    if (Test-Path $fileRegul) {
        $wb = $excel.Workbooks.Open($fileRegul, $null, $true)
        $vR = $wb.Sheets.Item(1).UsedRange.Value2
        $wb.Close($false)
        if ($vR) {
            for ($r = 2; $r -le $vR.GetLength(0); $r++) {
                $e = "$($vR[$r, 1])".Trim()
                if ($e) { $mapRegul[$e] = "$($vR[$r, 2])".Trim() }
            }
        }
    }

    # 4. Cargar Servicios Rechazados
    Write-Host "  -> Leyendo Servicios Rechazados..." -ForegroundColor DarkGray
    $setRechazados = @{}
    if (Test-Path $fileRechaz) {
        $wb = $excel.Workbooks.Open($fileRechaz, $null, $true)
        $vRec = $wb.Sheets.Item(1).UsedRange.Value2
        $wb.Close($false)
        if ($vRec) {
            for ($r = 2; $r -le $vRec.GetLength(0); $r++) {
                $s = "$($vRec[$r, 1])".Trim()
                if ($s) { $setRechazados[$s] = $true }
            }
        }
    }

    # 5. Cargar Equipo Contacto
    Write-Host "  -> Leyendo Equipo Contacto..." -ForegroundColor DarkGray
    $setEquipo = @{}
    if (Test-Path $fileEquipo) {
        $wb = $excel.Workbooks.Open($fileEquipo, $null, $true)
        $vEq = $wb.Sheets.Item(1).UsedRange.Value2
        $wb.Close($false)
        if ($vEq) {
            for ($r = 2; $r -le $vEq.GetLength(0); $r++) {
                $n = "$($vEq[$r, 1])".Trim()
                if ($n) { $setEquipo[$n] = $true }
            }
        }
    }

    # 6. Cargar Archivo Base
    Write-Host "[3/5] Leyendo archivo base ($nombreBase)..." -ForegroundColor Gray
    $wb = $excel.Workbooks.Open($fileBaseSAR, $null, $true)
    $vB = $wb.Sheets.Item(1).UsedRange.Value2
    $wb.Close($false)

}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "[4/5] Vinculando tablas, aplicando reglas SAR y formateando datos..." -ForegroundColor Gray

# Mapeo de columnas y recopilacion de todos los encabezados originales
$colsCount = $vB.GetLength(1)
$colsMap = @{}
$origHeaders = @()
for ($c = 1; $c -le $colsCount; $c++) {
    $h = "$($vB[1, $c])".Trim()
    if (-not $h) { $h = "Columna_$c" }
    $colsMap[$h] = $c
    $origHeaders += $h
}

function Find-ExactOrPatternCol($exactName, $pattern, $defaultIdx) {
    if ($colsMap.ContainsKey($exactName)) { return $colsMap[$exactName] }
    foreach ($k in $colsMap.Keys) {
        if ($k -like $pattern) { return $colsMap[$k] }
    }
    return $defaultIdx
}

$cTicket = Find-ExactOrPatternCol "Nro Ticket" "*Ticket*" 1
$cAfiliado = Find-ExactOrPatternCol "Nombres Afiliado" "*Afiliado*" 7
$cContacto = Find-ExactOrPatternCol "Nombre de persona de Contacto" "*Contacto*" 56
$cUsuario = Find-ExactOrPatternCol "Nombre Usuario Registro" "*Usuario*Registro*" 18
$cDeptoUsu = Find-ExactOrPatternCol "Departamento Usuario Registro" "*Departamento*Usuario*" 19
$cServicio = Find-ExactOrPatternCol "Servicio" "*Servicio*" 13
$cDeptoOri = Find-ExactOrPatternCol "Departamento Origen" "*Departamento*Origen*" 20
$cTel = Find-ExactOrPatternCol "Telefono de Contacto" "*fono*" 68
$cFecha = Find-ExactOrPatternCol "Fecha Registro" "*Fecha*Registro*" 2
$cHora = Find-ExactOrPatternCol "Hora Registro" "*Hora*Registro*" 4
$cEstado = Find-ExactOrPatternCol "Estado Servicio" "*Estado*" 43
$cTipoServicio = Find-ExactOrPatternCol "Tipo Servicio" "*Tipo*Servicio*" 15
$cEntidadesReg = Find-ExactOrPatternCol "Entidades Reguladoras" "*Reguladora*" 48
$cCorreo = Find-ExactOrPatternCol "Correo de Contacto" "*Correo*Contacto*" 50
$cObs = Find-ExactOrPatternCol "Observación Inicial" "*Inicial*" 46

$emailRegex = [regex]'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
$phoneRegex = [regex]'(?:(?<=\D)|^)(9\d{8})(?=\D|$)'
$cutoffTime = "17:00:00"

# Fecha actual para evaluar reglas de corte
$today = [DateTime]::Today
$isMonday = ($today.DayOfWeek -eq [DayOfWeek]::Monday)

# =========================================================================
# FUNCIONES SQL SAR: Encontrar_Correo, Encontrar_Nombre, Primera_Palabra
# =========================================================================
function Encontrar_Correo([string]$Texto) {
    if (-not $Texto) { return "Sin Correo" }
    $pOpen = $Texto.IndexOf("<")
    if ($pOpen -ge 0) {
        $pClose = $Texto.IndexOf(">", $pOpen + 1)
        if ($pClose -gt $pOpen) {
            $extracted = $Texto.Substring($pOpen + 1, $pClose - $pOpen - 1).Trim()
            if ($extracted) { return $extracted }
        }
    }
    return "Sin Correo"
}

function Encontrar_Nombre([string]$Texto) {
    if (-not $Texto) { return "SinNombre" }
    $pOpen = $Texto.IndexOf("{")
    if ($pOpen -ge 0) {
        $pClose = $Texto.IndexOf("}", $pOpen + 1)
        if ($pClose -gt $pOpen) {
            $extracted = $Texto.Substring($pOpen + 1, $pClose - $pOpen - 1).Trim()
            if ($extracted) { return $extracted }
        }
    }
    return "SinNombre"
}

function Primera_Palabra([string]$Texto) {
    if (-not $Texto) { return "" }
    $u = $Texto.Trim().ToUpper()
    $idx = $u.IndexOf(" ")
    if ($idx -gt 0) {
        return $u.Substring(0, $idx)
    }
    return $u
}

$rowsCount = $vB.GetLength(0)
$list = New-Object System.Collections.Generic.List[Object]

$totalCumplenSAR = 0
$totalCallCenter = 0
$totalHorario = 0

for ($r = 2; $r -le $rowsCount; $r++) {
    $ticket = "$($vB[$r, $cTicket])".Trim()
    if (-not $ticket) {
        $hasData = $false
        for ($colChk = 1; $colChk -le [Math]::Min(15, $vB.GetLength(1)); $colChk++) {
            if ("$($vB[$r, $colChk])".Trim()) { $hasData = $true; break }
        }
        if (-not $hasData) { continue }
        $ticket = "REG-$r"
    }
    
    $afiliadoRaw = "$($vB[$r, $cAfiliado])".Trim()
    $contactoRaw = "$($vB[$r, $cContacto])".Trim()
    $usuario = "$($vB[$r, $cUsuario])".Trim()
    $deptoUsu = "$($vB[$r, $cDeptoUsu])".Trim()
    $servicio = "$($vB[$r, $cServicio])".Trim()
    $deptoOri = "$($vB[$r, $cDeptoOri])".Trim()
    $tel = "$($vB[$r, $cTel])".Trim()
    $obs = "$($vB[$r, $cObs])"

    # 1. Nombre y Primera Palabra (según IIf exacto de la consulta SQL)
    # Primera_Palabra(IIf(Len(Trim(Nz(Contacto,"")))>0, Contacto, IIf(Len(Trim(Nz(Afiliado,"")))>0, Afiliado, Encontrar_Nombre(Obs))))
    $nombreCandidato = ""
    if ($contactoRaw -and $contactoRaw.Length -gt 0) {
        $nombreCandidato = $contactoRaw
    }
    elseif ($afiliadoRaw -and $afiliadoRaw.Length -gt 0) {
        $nombreCandidato = $afiliadoRaw
    }
    else {
        $nombreCandidato = Encontrar_Nombre $obs
    }
    $primerNombre = Primera_Palabra $nombreCandidato
    $afiliadoDisplay = if ($afiliadoRaw) { $afiliadoRaw } else { $nombreCandidato }

    # 2. Correo Afiliado (usando Encontrar_Correo)
    $correoRaw = "$($vB[$r, $cCorreo])".Trim()
    $correoAfiliado = ""
    if ($correoRaw -and $correoRaw.ToLower() -ne "sin correo" -and $correoRaw -ne "-") {
        $correoAfiliado = $correoRaw
    }
    else {
        $correoAfiliado = Encontrar_Correo $obs
    }

    # 2b. Telefono de Contacto (Logica SQL IIf / Encontrar_Telefono en Observacion Inicial)
    # Si no hay telefono, validar en Observacion Inicial: si empieza con 9 y 8 digitos a la derecha (9XXXXXXXX), rellenarlo
    $telefonoContacto = ""
    if ($tel -and $tel -ne "-" -and $tel -ne "0" -and $tel -ne "Sin teléfono") {
        $telefonoContacto = $tel
    }
    else {
        $mPhone = $phoneRegex.Match($obs)
        if ($mPhone.Success) {
            $telefonoContacto = $mPhone.Groups[1].Value
        }
        else {
            $telefonoContacto = "-"
        }
    }

    # 3. Tiempo de atencion (default 15 si no se encuentra en la tabla de tiempos)
    $tiempo = 15
    $tieneTiempoServ = $false
    if ($mapTiempos.ContainsKey($servicio)) {
        $tiempoVal = $mapTiempos[$servicio]
        if ($tiempoVal -ne $null -and "$tiempoVal" -ne "") {
            $tiempo = [double]$tiempoVal
            $tieneTiempoServ = $true
        }
    }

    # 4. Tablas enlazadas
    $agrupamiento = if ($mapAgrup.ContainsKey($servicio)) { $mapAgrup[$servicio] } else { "(Sin agrupamiento)" }
    $entidad = "$($vB[$r, $cEntidadesReg])".Trim()
    $regulador = if ($mapRegul.ContainsKey($entidad)) { $mapRegul[$entidad] } else { "(Sin regulador)" }
    $esRechazado = if ($setRechazados.ContainsKey($servicio)) { "S" } else { "No" }
    $esEquipo = if ($setEquipo.ContainsKey($usuario)) { "S" } else { "No" }

    $estado = "$($vB[$r, $cEstado])".Trim()
    $tipoServ = "$($vB[$r, $cTipoServicio])".Trim()

    # Formatear fecha (conservar fecha y hora si contiene ambas)
    $rawFecha = $vB[$r, $cFecha]
    $strFecha = ""
    $dateObj = $null
    if ($rawFecha -is [double] -or $rawFecha -is [int]) {
        try {
            $dt = [DateTime]::FromOADate([double]$rawFecha)
            if ($dt.TimeOfDay.TotalSeconds -gt 0) {
                $strFecha = $dt.ToString("dd/MM/yyyy HH:mm:ss")
            }
            else {
                $strFecha = $dt.ToString("dd/MM/yyyy")
            }
            $dateObj = $dt.Date
        }
        catch {
            $strFecha = "$rawFecha"
        }
    }
    else {
        $strFecha = "$rawFecha".Trim()
    }

    # Formatear hora (convertir fraccion de dia OA Date o extraer HH:mm:ss)
    $rawHora = $vB[$r, $cHora]
    $strHora = "00:00:00"
    if ($rawHora -is [double] -or $rawHora -is [int]) {
        try {
            $ts = [TimeSpan]::FromDays([double]$rawHora)
            $strHora = "{0:D2}:{1:D2}:{2:D2}" -f $ts.Hours, $ts.Minutes, $ts.Seconds
        }
        catch {
            $strHora = "$rawHora".Trim()
        }
    }
    else {
        $strHora = "$rawHora".Trim()
        if ($strHora -match '\d{2}:\d{2}:\d{2}') {
            $strHora = $Matches[0]
        }
        elseif ($strHora.Length -eq 5) {
            $strHora += ":00"
        }
    }

    # =========================================================================
    # EVALUACION DE CONDICIONES WHERE EXACTAS DE LA CONSULTA SQL
    # =========================================================================
    # Filtros base comunes:
    $isRegOk = ($regulador -eq "TRANSACCIONES DIARIAS" -or $regulador -eq "(Sin regulador)" -or [string]::IsNullOrEmpty($regulador))
    $isNoRech = ($esRechazado -eq "No")
    $isTipoOk = ($tipoServ.ToLower() -eq "requerimiento" -or $tipoServ.ToLower() -eq "reclamo")
    $isEstadoOk = ($estado.ToLower() -eq "en curso")

    $cumpleBase = ($isRegOk -and $isNoRech -and $isTipoOk -and $isEstadoOk)

    # Rama Call Center / CCG:
    $esCallCenter = (
        $deptoOri -like "*Call Center*" -or 
        $deptoUsu -like "*Call Center*" -or 
        $ticket -like "CCG-*"
    )

    # Rama Horarios / SLA (sin filtrar por correo: se toman también los sin correos según la consulta SQL):
    $isNotInterno = ($deptoOri -notlike "*interno*")
    
    # Evaluacion de corte de fecha/hora:
    # 1. Hoy <= 17:00 (cualquier dia)
    # 2. Si NO es Lunes: Ayer > 17:00
    # 3. Si ES Lunes: Viernes > 17:00, Sabado completo, Domingo completo
    $cumpleCorteHorario = $false
    if ($dateObj) {
        $horaComparar = if ($strHora) { $strHora } else { "00:00:00" }

        # Caso 1: Hoy <= 17:00
        if ($dateObj -eq $today -and $horaComparar -le $cutoffTime) {
            $cumpleCorteHorario = $true
        }
        # Caso 2: Si NO es lunes y fue ayer > 17:00
        elseif (-not $isMonday -and $dateObj -eq $today.AddDays(-1) -and $horaComparar -gt $cutoffTime) {
            $cumpleCorteHorario = $true
        }
        # Caso 3: Si ES lunes
        elseif ($isMonday) {
            if ($dateObj -eq $today.AddDays(-3) -and $horaComparar -gt $cutoffTime) {
                $cumpleCorteHorario = $true
            }
            elseif ($dateObj -eq $today.AddDays(-2) -or $dateObj -eq $today.AddDays(-1)) {
                $cumpleCorteHorario = $true
            }
        }
    }

    # Se toman también los que no tienen correo / 'Sin Correo'
    $esHorarioSLA = ($isNotInterno -and $tieneTiempoServ -and $cumpleCorteHorario)

    # Cumple SAR total (si cumple filtros base Y (Call Center O Horario))
    $cumpleSAR = ($cumpleBase -and ($esCallCenter -or $esHorarioSLA))

    if ($cumpleSAR) { $totalCumplenSAR++ }
    if ($cumpleBase -and $esCallCenter) { $totalCallCenter++ }
    if ($cumpleBase -and $esHorarioSLA) { $totalHorario++ }

    $esMerino = ($usuario -like "*MERINO CARRASCO JEANFRANCO*")
    $esMerinoSinInterno = ($esMerino -and $isNotInterno)

    # Construir registro con TODAS las columnas originales del Excel
    $rowDict = [ordered]@{}
    for ($c = 1; $c -le $colsCount; $c++) {
        $colName = $origHeaders[$c - 1]
        $rawVal = $vB[$r, $c]

        # Formateo si es Fecha Registro
        if ($c -eq $cFecha) {
            $rawVal = $strFecha
        }
        # Formateo si es Hora Registro
        elseif ($c -eq $cHora) {
            $rawVal = $strHora
        }
        # Cualquier otra columna de fecha en el Excel
        elseif ($colName -like "*Fecha*" -and ($rawVal -is [double] -or $rawVal -is [int]) -and $rawVal -gt 30000) {
            try {
                $dtC = [DateTime]::FromOADate([double]$rawVal)
                if ($dtC.TimeOfDay.TotalSeconds -gt 0) {
                    $rawVal = $dtC.ToString("dd/MM/yyyy HH:mm:ss")
                }
                else {
                    $rawVal = $dtC.ToString("dd/MM/yyyy")
                }
            }
            catch {}
        }
        # Cualquier otra columna de hora en el Excel
        elseif ($colName -like "*Hora*" -and ($rawVal -is [double] -or $rawVal -is [int]) -and $rawVal -ge 0 -and $rawVal -lt 1) {
            try {
                $tsC = [TimeSpan]::FromDays([double]$rawVal)
                $rawVal = "{0:D2}:{1:D2}:{2:D2}" -f $tsC.Hours, $tsC.Minutes, $tsC.Seconds
            }
            catch {}
        }
        # Si es la columna de Teléfono de Contacto
        elseif ($c -eq $cTel) {
            $rawVal = if ($telefonoContacto) { $telefonoContacto } else { "-" }
        }

        $rowDict[$colName] = if ($rawVal -ne $null) { "$rawVal".Trim() } else { "" }
    }

    # Agregar campos calculados / complementarios para filtros y lógica SQL
    $rowDict["Tiempo Atención"] = $tiempo
    $rowDict["ticket"] = $ticket
    $rowDict["primer_nombre"] = if ($primerNombre) { $primerNombre } else { "-" }
    $rowDict["tiempo_atencion"] = $tiempo
    $rowDict["correo_afiliado"] = $correoAfiliado
    $rowDict["telefono"] = if ($telefonoContacto) { $telefonoContacto } else { "-" }
    $rowDict["cumple_sar"] = if ($cumpleSAR) { "S" } else { "No" }
    $rowDict["es_merino_sin_int"] = if ($esMerinoSinInterno) { "S" } else { "No" }
    $rowDict["criterio_motivo"] = if ($cumpleSAR) { 
        if ($esCallCenter -and $esHorarioSLA) { "Call Center y Horario SLA" }
        elseif ($esCallCenter) { "Call Center / CCG" }
        else { "Horario y SLA" }
    }
    else { "Fuera de criterio" }

    $list.Add($rowDict)
}

Write-Host "[5/5] Exportando datos consolidados a data.js con TODAS las columnas originales..." -ForegroundColor Gray

# Lista completa de columnas del archivo original + Tiempo Atención
$colTiempoAtencion = [System.Text.Encoding]::UTF8.GetString([byte[]]@(0x54, 0x69, 0x65, 0x6D, 0x70, 0x6F, 0x20, 0x41, 0x74, 0x65, 0x6E, 0x63, 0x69, 0xC3, 0xB3, 0x6E))
$systemColumns = @($origHeaders)
if (-not ($systemColumns -contains $colTiempoAtencion)) {
    $systemColumns += $colTiempoAtencion
}

$refObj = [ordered]@{
    reguladores = $mapRegul
    serviciosRechazados = @($setRechazados.Keys)
    tiemposAtencion = $mapTiempos
}
$refJson = $refObj | ConvertTo-Json -Depth 3 -Compress

$colsJson = $systemColumns | ConvertTo-Json -Compress
$json = $list | ConvertTo-Json -Depth 2 -Compress
$timestamp = (Get-Date).ToString("dd/MM/yyyy HH:mm:ss")

$jsContent = @"
// Archivo generado automaticamente por actualizar_datos.ps1
// Ultima actualizacion: $timestamp
// Archivo base: $nombreBase
window.METADATOS_SISTEMA = {
    fecha_actualizacion: "$timestamp",
    archivo_base: "$nombreBase",
    total_registros: $($list.Count),
    total_cumplen_sar: $totalCumplenSAR,
    total_call_center: $totalCallCenter,
    total_horario_sla: $totalHorario
};
window.TABLAS_REFERENCIA = $refJson;
window.COLUMNAS_SISTEMA = $colsJson;
window.DATOS_SISTEMA = $json;
"@

$dataFile = Join-Path $scriptDir "data.js"
[System.IO.File]::WriteAllText($dataFile, $jsContent, [System.Text.Encoding]::UTF8)

$sw.Stop()
Write-Host "========================================================================" -ForegroundColor Green
Write-Host " [EXITO] Proceso completado en $($sw.ElapsedMilliseconds) ms" -ForegroundColor Green
Write-Host " Archivo base procesado: $nombreBase" -ForegroundColor Green
Write-Host " Total registros procesados: $($list.Count)" -ForegroundColor Green
Write-Host " Registros que cumplen Criterio SAR: $totalCumplenSAR" -ForegroundColor Green
Write-Host "   - Por Call Center / CCG: $totalCallCenter" -ForegroundColor Green
Write-Host "   - Por Horario / SLA: $totalHorario" -ForegroundColor Green
Write-Host " Archivo generado: $dataFile" -ForegroundColor Green
Write-Host "========================================================================" -ForegroundColor Green
