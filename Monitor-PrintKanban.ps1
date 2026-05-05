<#
.SYNOPSIS
    Monitor de impresión Kanban DOOR - Detecta desincronización entre secuencias recibidas y secuencias impresas.

.DESCRIPTION
    Este script compara:
    1. La última secuencia recibida en los archivos H_*.TXT (columna 11, 3 caracteres)
    2. La última secuencia impresa según el log (líneas "Fin impresion Kanban DOOR ... Secuencia NNN")
    
    Si la diferencia supera la tolerancia configurada, el script sugiere (o ejecuta) las acciones correctivas:
    - Detener el servicio ImportTXTTASAsvcHost
    - Reiniciar la consola de impresión
    - Esperar a que se imprima la última secuencia
    - Reiniciar el servicio ImportTXTTASAsvcHost

.PARAMETER Tolerancia
    Número máximo de secuencias de diferencia permitidas antes de actuar. Default: 5

.PARAMETER ModoEjecucion
    "demo"      = Solo muestra qué haría (default en esta primera instancia)
    "preguntar" = Pregunta antes de cada acción
    "auto"      = Ejecuta automáticamente las acciones correctivas

.PARAMETER Servidor
    IP o hostname del servidor donde se encuentran los archivos y el servicio. Default: 172.17.132.153

.PARAMETER RutaLogs
    Ruta UNC a la carpeta de logs de la consola de impresión.

.PARAMETER RutaProcesados
    Ruta UNC a la carpeta de archivos H_ procesados.

.PARAMETER NombreServicio
    Nombre del servicio Windows a controlar. Default: ImportTXTTASAsvcHost

.EXAMPLE
    .\Monitor-PrintKanban.ps1
    # Ejecuta en modo demo con tolerancia 5

.EXAMPLE
    .\Monitor-PrintKanban.ps1 -Tolerancia 3 -ModoEjecucion "preguntar"
    # Tolerancia de 3, pregunta antes de actuar
#>

param(
    [int]$Tolerancia = 5,
    
    [ValidateSet("demo", "preguntar", "auto")]
    [string]$ModoEjecucion = "demo",
    
    [string]$Servidor = "172.17.132.153",
    
    [string]$RutaLogs = "\\172.17.132.153\c$\Program Files\JITMS\prod\print-services-2.4.2-console\logs",
    
    [string]$RutaProcesados = "\\172.17.132.153\c$\Users\rpareja1ta\Desktop\SFTP\SFTP_TASA\Procesados",
    
    [string]$NombreServicio = "ImportTXTTASAsvcHost",

    [int]$PosicionSecuencia = 10,
    
    [int]$LargoSecuencia = 3
)

# ============================================================
# FUNCIONES
# ============================================================

function Write-Banner {
    $banner = @"

╔══════════════════════════════════════════════════════════════╗
║          MONITOR DE IMPRESION KANBAN DOOR                   ║
║          ════════════════════════════════                    ║
╚══════════════════════════════════════════════════════════════╝
"@
    Write-Host $banner -ForegroundColor Cyan
    Write-Host "  Fecha/Hora    : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    Write-Host "  Servidor      : $Servidor" -ForegroundColor Gray
    Write-Host "  Tolerancia    : $Tolerancia secuencias" -ForegroundColor Gray
    Write-Host "  Modo          : $ModoEjecucion" -ForegroundColor Gray
    Write-Host ""
}

function Get-UltimaSecuenciaRecibida {
    <#
    .SYNOPSIS
        Obtiene la secuencia del último archivo H_*.TXT ordenado por fecha.
    #>
    try {
        $ultimoArchivo = Get-ChildItem -Path $RutaProcesados -Filter "H_*.TXT" -ErrorAction Stop |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if (-not $ultimoArchivo) {
            Write-Host "  [ERROR] No se encontraron archivos H_*.TXT en $RutaProcesados" -ForegroundColor Red
            return $null
        }

        $contenido = (Get-Content $ultimoArchivo.FullName -Raw -ErrorAction Stop).Trim()
        
        if ($contenido.Length -lt ($PosicionSecuencia + $LargoSecuencia)) {
            Write-Host "  [ERROR] El archivo $($ultimoArchivo.Name) tiene contenido demasiado corto" -ForegroundColor Red
            return $null
        }

        $secuencia = $contenido.Substring($PosicionSecuencia, $LargoSecuencia).Trim()
        
        return @{
            Secuencia    = [int]$secuencia
            Archivo      = $ultimoArchivo.Name
            FechaArchivo = $ultimoArchivo.LastWriteTime
            RutaCompleta = $ultimoArchivo.FullName
        }
    }
    catch {
        Write-Host "  [ERROR] No se pudo leer la secuencia de archivos H_: $_" -ForegroundColor Red
        return $null
    }
}

function Get-UltimaSecuenciaImpresa {
    <#
    .SYNOPSIS
        Busca en el log más reciente la última línea "Fin impresion Kanban DOOR ... Secuencia NNN"
        y extrae el número de secuencia.
    #>
    try {
        $ultimoLog = Get-ChildItem -Path $RutaLogs -Filter "log-*.txt" -ErrorAction Stop |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if (-not $ultimoLog) {
            Write-Host "  [ERROR] No se encontraron archivos de log en $RutaLogs" -ForegroundColor Red
            return $null
        }

        # Leer las últimas 500 líneas del log para buscar la última impresión
        $lineas = Get-Content $ultimoLog.FullName -Tail 500 -ErrorAction Stop

        # Buscar líneas con "Fin impresion Kanban DOOR" y extraer secuencia
        $patronSecuencia = "Fin impresion Kanban DOOR.*Secuencia\s+(\d+)"
        $ultimaSecuencia = $null
        $ultimaLinea = $null

        foreach ($linea in $lineas) {
            if ($linea -match $patronSecuencia) {
                $ultimaSecuencia = [int]$Matches[1]
                $ultimaLinea = $linea
            }
        }

        if ($null -eq $ultimaSecuencia) {
            Write-Host "  [ADVERTENCIA] No se encontro 'Fin impresion Kanban DOOR' en las ultimas 500 lineas del log" -ForegroundColor Yellow
            return $null
        }

        return @{
            Secuencia   = $ultimaSecuencia
            LineaLog    = $ultimaLinea.Trim()
            ArchivoLog  = $ultimoLog.Name
        }
    }
    catch {
        Write-Host "  [ERROR] No se pudo leer el log: $_" -ForegroundColor Red
        return $null
    }
}

function Get-EstadoServicio {
    try {
        $svc = Get-Service -Name $NombreServicio -ComputerName $Servidor -ErrorAction Stop
        return $svc.Status
    }
    catch {
        Write-Host "  [ERROR] No se pudo consultar el servicio $NombreServicio en $Servidor : $_" -ForegroundColor Red
        return $null
    }
}

function Invoke-AccionCorrectiva {
    param(
        [int]$SecuenciaRecibida,
        [int]$SecuenciaImpresa,
        [int]$Diferencia
    )

    Write-Host ""
    Write-Host "  ┌─────────────────────────────────────────────────────────────┐" -ForegroundColor Red
    Write-Host "  │  ACCION CORRECTIVA REQUERIDA                               │" -ForegroundColor Red
    Write-Host "  │  Diferencia ($Diferencia) supera tolerancia ($Tolerancia)                     │" -ForegroundColor Red
    Write-Host "  └─────────────────────────────────────────────────────────────┘" -ForegroundColor Red
    Write-Host ""

    $pasos = @(
        @{ Orden = 1; Descripcion = "Detener servicio '$NombreServicio' en $Servidor"; Comando = "Stop-Service -Name '$NombreServicio' -Force" },
        @{ Orden = 2; Descripcion = "Reiniciar la consola de impresion (print-services-2.4.2-console)"; Comando = "# Reinicio manual de la consola requerido" },
        @{ Orden = 3; Descripcion = "Esperar a que se imprima la secuencia $SecuenciaRecibida"; Comando = "# Monitorear log hasta ver 'Fin impresion Kanban DOOR ... Secuencia $SecuenciaRecibida'" },
        @{ Orden = 4; Descripcion = "Iniciar servicio '$NombreServicio' en $Servidor"; Comando = "Start-Service -Name '$NombreServicio'" }
    )

    switch ($ModoEjecucion) {
        "demo" {
            Write-Host "  === MODO DEMOSTRACION ===" -ForegroundColor Yellow
            Write-Host "  Las siguientes acciones se ejecutarian si estuviera en modo 'auto' o 'preguntar':" -ForegroundColor Yellow
            Write-Host ""
            foreach ($paso in $pasos) {
                Write-Host "  PASO $($paso.Orden): $($paso.Descripcion)" -ForegroundColor White
                Write-Host "         Comando: $($paso.Comando)" -ForegroundColor DarkGray
                Write-Host ""
            }
            Write-Host "  ─────────────────────────────────────────────────────────────" -ForegroundColor Yellow
            Write-Host "  Para ejecutar las acciones, use:" -ForegroundColor Yellow
            Write-Host "    .\Monitor-PrintKanban.ps1 -ModoEjecucion 'preguntar'" -ForegroundColor Green
            Write-Host "    .\Monitor-PrintKanban.ps1 -ModoEjecucion 'auto'" -ForegroundColor Green
            Write-Host ""
        }

        "preguntar" {
            Write-Host "  === MODO INTERACTIVO ===" -ForegroundColor Magenta
            Write-Host ""
            foreach ($paso in $pasos) {
                Write-Host "  PASO $($paso.Orden): $($paso.Descripcion)" -ForegroundColor White
                $respuesta = Read-Host "  ¿Ejecutar? (S/N)"
                if ($respuesta -eq "S" -or $respuesta -eq "s") {
                    Write-Host "  Ejecutando..." -ForegroundColor Green
                    Invoke-PasoCorrectivo -Paso $paso -SecuenciaRecibida $SecuenciaRecibida
                } else {
                    Write-Host "  Omitido por el usuario." -ForegroundColor Yellow
                }
                Write-Host ""
            }
        }

        "auto" {
            Write-Host "  === MODO AUTOMATICO ===" -ForegroundColor Red
            Write-Host ""
            foreach ($paso in $pasos) {
                Write-Host "  PASO $($paso.Orden): $($paso.Descripcion)" -ForegroundColor White
                Invoke-PasoCorrectivo -Paso $paso -SecuenciaRecibida $SecuenciaRecibida
                Write-Host ""
            }
        }
    }
}

function Invoke-PasoCorrectivo {
    param(
        [hashtable]$Paso,
        [int]$SecuenciaRecibida
    )

    switch ($Paso.Orden) {
        1 {
            # Detener servicio
            try {
                Write-Host "    Deteniendo servicio $NombreServicio en $Servidor..." -ForegroundColor Yellow
                $svc = Get-Service -Name $NombreServicio -ComputerName $Servidor -ErrorAction Stop
                if ($svc.Status -eq "Running") {
                    Set-Service -Name $NombreServicio -ComputerName $Servidor -Status Stopped -ErrorAction Stop
                    # Alternativa con Invoke-Command si Set-Service no funciona remotamente
                    # Invoke-Command -ComputerName $Servidor -ScriptBlock { Stop-Service -Name $using:NombreServicio -Force }
                    Write-Host "    [OK] Servicio detenido." -ForegroundColor Green
                } else {
                    Write-Host "    [INFO] El servicio ya estaba detenido ($($svc.Status))." -ForegroundColor Cyan
                }
            }
            catch {
                Write-Host "    [ERROR] No se pudo detener el servicio: $_" -ForegroundColor Red
                Write-Host "    Intente manualmente: Invoke-Command -ComputerName $Servidor -ScriptBlock { Stop-Service -Name '$NombreServicio' -Force }" -ForegroundColor DarkYellow
            }
        }

        2 {
            # Reiniciar consola - esto típicamente requiere intervención manual o un script específico
            Write-Host "    [MANUAL] Reinicie la consola de impresion manualmente." -ForegroundColor Yellow
            Write-Host "    Ruta: \\$Servidor\c$\Program Files\JITMS\prod\print-services-2.4.2-console" -ForegroundColor DarkGray
            if ($ModoEjecucion -eq "preguntar") {
                Read-Host "    Presione ENTER cuando la consola haya sido reiniciada"
            } elseif ($ModoEjecucion -eq "auto") {
                Write-Host "    [ESPERA] Esperando 30 segundos para que reinicie la consola manualmente..." -ForegroundColor Yellow
                Start-Sleep -Seconds 30
            }
        }

        3 {
            # Esperar impresión de la última secuencia
            Write-Host "    Esperando que se imprima la secuencia $SecuenciaRecibida..." -ForegroundColor Yellow
            $maxIntentos = 60  # 60 intentos x 10 segundos = 10 minutos
            $intento = 0
            $impresa = $false

            while ($intento -lt $maxIntentos -and -not $impresa) {
                $intento++
                $resultado = Get-UltimaSecuenciaImpresa
                if ($null -ne $resultado -and $resultado.Secuencia -ge $SecuenciaRecibida) {
                    Write-Host "    [OK] Secuencia $($resultado.Secuencia) impresa exitosamente!" -ForegroundColor Green
                    $impresa = $true
                } else {
                    $secActual = if ($null -ne $resultado) { $resultado.Secuencia } else { "?" }
                    Write-Host "    [$intento/$maxIntentos] Secuencia actual impresa: $secActual - Esperando $SecuenciaRecibida... (esperando 10s)" -ForegroundColor DarkGray
                    Start-Sleep -Seconds 10
                }
            }

            if (-not $impresa) {
                Write-Host "    [TIMEOUT] La secuencia $SecuenciaRecibida no fue impresa en el tiempo esperado." -ForegroundColor Red
            }
        }

        4 {
            # Iniciar servicio
            try {
                Write-Host "    Iniciando servicio $NombreServicio en $Servidor..." -ForegroundColor Yellow
                $svc = Get-Service -Name $NombreServicio -ComputerName $Servidor -ErrorAction Stop
                if ($svc.Status -ne "Running") {
                    Set-Service -Name $NombreServicio -ComputerName $Servidor -Status Running -ErrorAction Stop
                    Write-Host "    [OK] Servicio iniciado." -ForegroundColor Green
                } else {
                    Write-Host "    [INFO] El servicio ya estaba corriendo." -ForegroundColor Cyan
                }
            }
            catch {
                Write-Host "    [ERROR] No se pudo iniciar el servicio: $_" -ForegroundColor Red
                Write-Host "    Intente manualmente: Invoke-Command -ComputerName $Servidor -ScriptBlock { Start-Service -Name '$NombreServicio' }" -ForegroundColor DarkYellow
            }
        }
    }
}

# ============================================================
# EJECUCION PRINCIPAL
# ============================================================

Write-Banner

# 1. Obtener estado del servicio
Write-Host "  [1/3] Consultando estado del servicio..." -ForegroundColor Cyan
$estadoServicio = Get-EstadoServicio
if ($null -ne $estadoServicio) {
    $colorServicio = if ($estadoServicio -eq "Running") { "Green" } else { "Red" }
    Write-Host "         Servicio '$NombreServicio': $estadoServicio" -ForegroundColor $colorServicio
} else {
    Write-Host "         Servicio '$NombreServicio': NO DISPONIBLE" -ForegroundColor Red
}
Write-Host ""

# 2. Obtener última secuencia recibida (archivos H_)
Write-Host "  [2/3] Leyendo ultima secuencia recibida (archivos H_)..." -ForegroundColor Cyan
$recibida = Get-UltimaSecuenciaRecibida
if ($null -ne $recibida) {
    Write-Host "         Secuencia recibida : $($recibida.Secuencia)" -ForegroundColor White
    Write-Host "         Archivo            : $($recibida.Archivo)" -ForegroundColor DarkGray
    Write-Host "         Fecha archivo      : $($recibida.FechaArchivo)" -ForegroundColor DarkGray
} else {
    Write-Host "         [ERROR] No se pudo obtener la secuencia recibida." -ForegroundColor Red
    exit 1
}
Write-Host ""

# 3. Obtener última secuencia impresa (log)
Write-Host "  [3/3] Leyendo ultima secuencia impresa (log)..." -ForegroundColor Cyan
$impresa = Get-UltimaSecuenciaImpresa
if ($null -ne $impresa) {
    Write-Host "         Secuencia impresa  : $($impresa.Secuencia)" -ForegroundColor White
    Write-Host "         Log                : $($impresa.ArchivoLog)" -ForegroundColor DarkGray
    Write-Host "         Linea              : $($impresa.LineaLog)" -ForegroundColor DarkGray
} else {
    Write-Host "         [ERROR] No se pudo obtener la secuencia impresa." -ForegroundColor Red
    exit 1
}
Write-Host ""

# 4. Calcular diferencia y evaluar
$diferencia = $recibida.Secuencia - $impresa.Secuencia

Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║  RESUMEN                                                    ║" -ForegroundColor Cyan
Write-Host "  ╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
Write-Host ("  ║  Secuencia recibida  : {0,-37}║" -f $recibida.Secuencia) -ForegroundColor White
Write-Host ("  ║  Secuencia impresa   : {0,-37}║" -f $impresa.Secuencia) -ForegroundColor White

if ($diferencia -le 0) {
    Write-Host ("  ║  Diferencia          : {0,-37}║" -f "$diferencia ✅ AL DIA") -ForegroundColor Green
} elseif ($diferencia -le $Tolerancia) {
    Write-Host ("  ║  Diferencia          : {0,-37}║" -f "$diferencia ⚠️  DENTRO DE TOLERANCIA ($Tolerancia)") -ForegroundColor Yellow
} else {
    Write-Host ("  ║  Diferencia          : {0,-37}║" -f "$diferencia ❌ SUPERA TOLERANCIA ($Tolerancia)") -ForegroundColor Red
}

Write-Host ("  ║  Tolerancia          : {0,-37}║" -f $Tolerancia) -ForegroundColor Gray
Write-Host ("  ║  Servicio            : {0,-37}║" -f "$NombreServicio ($estadoServicio)") -ForegroundColor Gray
Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# 5. Actuar si es necesario
if ($diferencia -gt $Tolerancia) {
    Invoke-AccionCorrectiva -SecuenciaRecibida $recibida.Secuencia -SecuenciaImpresa $impresa.Secuencia -Diferencia $diferencia
} elseif ($diferencia -gt 0 -and $diferencia -le $Tolerancia) {
    Write-Host "  [INFO] La diferencia ($diferencia) esta dentro de la tolerancia ($Tolerancia)." -ForegroundColor Yellow
    Write-Host "         No se requiere accion correctiva por el momento." -ForegroundColor Yellow
} else {
    Write-Host "  [OK] Las secuencias estan sincronizadas. No se requiere accion." -ForegroundColor Green
}

Write-Host ""
Write-Host "  ─────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Fin del monitoreo: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor DarkGray
Write-Host ""
