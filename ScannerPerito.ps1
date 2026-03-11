#Requires -Version 5.1
<#
.SYNOPSIS
    ScannerPerito - Generador de Evidencia Digital para Peritaje Informatico
.DESCRIPTION
    Script para recoleccion forense de evidencia digital de un disco o carpeta.
    Genera un listado completo de archivos con hashes, metadatos y un informe
    pericial en formato TXT con toda la informacion de trazabilidad.
.NOTES
    Version: 1.0.0
    Compatible: Windows 10, Windows 11, Windows Server 2019+
    PowerShell: 5.1+
#>

[CmdletBinding()]
param()

# ============================================================================
# VARIABLES CONFIGURABLES
# ============================================================================
$SCRIPT_VERSION          = "1.0.0"
$HASH_ALGORITHM          = "SHA256"
$CSV_DELIMITER           = "|"
$EXCEL_MAX_ROWS          = 1048575
$PROGRESS_INTERVAL       = 100

# ============================================================================
# FUNCIONES UTILITARIAS
# ============================================================================

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes Bytes"
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Warning','Error')]
        [string]$Level = 'Info',
        [System.IO.StreamWriter]$LogWriter = $null
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        'Info'    { Write-Host $logLine -ForegroundColor Green }
        'Warning' { Write-Host $logLine -ForegroundColor Yellow }
        'Error'   { Write-Host $logLine -ForegroundColor Red }
    }

    if ($LogWriter) {
        $LogWriter.WriteLine($logLine)
        $LogWriter.Flush()
    }
}

# ============================================================================
# 1. VALIDACION DE COMPATIBILIDAD DEL SO
# ============================================================================

function Test-OSCompatibility {
    $os = Get-CimInstance Win32_OperatingSystem
    $buildNumber = [int]$os.BuildNumber

    # Windows 10 1809 / Server 2019 = build 17763
    $minBuild = 17763

    if ($buildNumber -lt $minBuild) {
        Write-Host ""
        Write-Host "  ERROR: Sistema operativo no compatible." -ForegroundColor Red
        Write-Host "  Se requiere Windows 10 (1809+), Windows 11 o Windows Server 2019+." -ForegroundColor Red
        Write-Host "  SO detectado: $($os.Caption) (Build $buildNumber)" -ForegroundColor Red
        Write-Host ""
        exit 1
    }

    if ($PSVersionTable.PSVersion.Major -lt 5 -or
       ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1)) {
        Write-Host ""
        Write-Host "  ERROR: Se requiere PowerShell 5.1 o superior." -ForegroundColor Red
        Write-Host "  Version detectada: $($PSVersionTable.PSVersion)" -ForegroundColor Red
        Write-Host ""
        exit 1
    }

    return @{
        OSCaption   = $os.Caption
        OSVersion   = $os.Version
        BuildNumber = $buildNumber
    }
}

# ============================================================================
# 2. LECTURA DE CONFIGURACION
# ============================================================================

function Read-Configuration {
    $configPath = Join-Path $PSScriptRoot "config.json"

    if (-not (Test-Path $configPath)) {
        Write-Host ""
        Write-Host "  AVISO: No se encontro config.json en $PSScriptRoot" -ForegroundColor Yellow
        Write-Host "  Creando plantilla de configuracion..." -ForegroundColor Yellow

        $defaultConfig = @{
            entity_name       = "Nombre de la Entidad / Empresa"
            nit               = "NIT o Identificacion Fiscal"
            address           = "Direccion de la entidad"
            city              = "Ciudad"
            phone             = "+57 XXX XXX XXXX"
            email             = "contacto@entidad.com"
            author            = "Nombre del Perito / Autor del Informe"
            author_id         = "Cedula o identificacion"
            notes             = ""
        }

        $defaultConfig | ConvertTo-Json -Depth 3 | Out-File -FilePath $configPath -Encoding utf8
        Write-Host "  Plantilla creada en: $configPath" -ForegroundColor Yellow
        Write-Host "  Por favor, edite el archivo con los datos reales y vuelva a ejecutar." -ForegroundColor Yellow
        Write-Host ""
    }

    $config = Get-Content -Path $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

    # Validar campos requeridos
    $requiredFields = @('entity_name', 'author')
    foreach ($field in $requiredFields) {
        $value = $config.$field
        if ([string]::IsNullOrWhiteSpace($value) -or $value -match '^Nombre') {
            Write-Host "  AVISO: El campo '$field' en config.json parece no estar configurado." -ForegroundColor Yellow
        }
    }

    return $config
}

# ============================================================================
# 2b. NUMERO DE CASO (SEED AUTO-INCREMENTAL)
# ============================================================================

function Get-NextCaseNumber {
    $seedPath = Join-Path $PSScriptRoot "case_seed.json"

    if (Test-Path $seedPath) {
        $seed = Get-Content -Path $seedPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $nextNumber = [int]$seed.last_case_number + 1
    }
    else {
        $nextNumber = 1
    }

    return [PSCustomObject]@{
        SeedPath   = $seedPath
        CaseNumber = $nextNumber
        CaseId     = "SP-{0:D6}" -f $nextNumber
    }
}

function Save-CaseNumber {
    param(
        [string]$SeedPath,
        [int]$CaseNumber,
        [string]$CaseId,
        [string]$CaseDescription,
        [string]$Timestamp
    )

    # Cargar historial existente o crear nuevo
    if (Test-Path $SeedPath) {
        $seed = Get-Content -Path $SeedPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $history = @()
        if ($seed.history) {
            $history = @($seed.history)
        }
    }
    else {
        $history = @()
    }

    # Agregar entrada al historial
    $entry = @{
        case_id     = $CaseId
        case_number = $CaseNumber
        description = $CaseDescription
        timestamp   = $Timestamp
    }
    $history += $entry

    $seedData = @{
        last_case_number = $CaseNumber
        last_updated     = $Timestamp
        history          = $history
    }

    $seedData | ConvertTo-Json -Depth 5 | Out-File -FilePath $SeedPath -Encoding utf8 -Force
}

# ============================================================================
# 2c. DESCRIPCION DEL CASO (INPUT DEL USUARIO)
# ============================================================================

function Read-CaseDescription {
    param([string]$CaseId)

    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host "    DATOS DEL CASO" -ForegroundColor Cyan
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Numero de caso asignado: $CaseId" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Ingrese la descripcion del caso." -ForegroundColor White
    Write-Host "  Ejemplo: Extraccion de evidencia digital del disco duro del equipo" -ForegroundColor DarkGray
    Write-Host "           de escritorio asignado al area de contabilidad, solicitada" -ForegroundColor DarkGray
    Write-Host "           por la Fiscalia General dentro del proceso 2026-00123." -ForegroundColor DarkGray
    Write-Host ""

    do {
        $description = Read-Host "  Descripcion"
        if ([string]::IsNullOrWhiteSpace($description)) {
            Write-Host "  La descripcion no puede estar vacia. Intente de nuevo." -ForegroundColor Red
        }
    } while ([string]::IsNullOrWhiteSpace($description))

    return $description
}

# ============================================================================
# 3. MENU DE SELECCION DE UNIDAD
# ============================================================================

function Show-DriveSelectionMenu {
    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host "    SELECCION DE UNIDAD A PROCESAR" -ForegroundColor Cyan
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""

    $volumes = Get-Volume | Where-Object {
        $_.DriveLetter -and $_.DriveType -in @('Fixed','Removable')
    } | Sort-Object DriveLetter

    if ($volumes.Count -eq 0) {
        Write-Host "  ERROR: No se encontraron unidades disponibles." -ForegroundColor Red
        exit 1
    }

    $index = 1
    $volumeList = @()
    foreach ($vol in $volumes) {
        $label = if ($vol.FileSystemLabel) { $vol.FileSystemLabel } else { "Sin etiqueta" }
        $sizeStr = Format-FileSize $vol.Size
        $freeStr = Format-FileSize $vol.SizeRemaining
        $fs = if ($vol.FileSystem) { $vol.FileSystem } else { "N/A" }

        Write-Host "    [$index] $($vol.DriveLetter): - $label ($fs, $sizeStr, Libre: $freeStr)" -ForegroundColor White
        $volumeList += $vol
        $index++
    }

    Write-Host ""
    do {
        $selection = Read-Host "  Seleccione el numero de la unidad"
        $selNum = 0
        $valid = [int]::TryParse($selection, [ref]$selNum) -and $selNum -ge 1 -and $selNum -le $volumeList.Count
        if (-not $valid) {
            Write-Host "  Seleccion invalida. Intente de nuevo." -ForegroundColor Red
        }
    } while (-not $valid)

    $selectedVolume = $volumeList[$selNum - 1]
    return $selectedVolume.DriveLetter
}

# ============================================================================
# 4. MENU DE ALCANCE (UNIDAD COMPLETA O CARPETA)
# ============================================================================

function Show-ScopeSelectionMenu {
    param([string]$DriveLetter)

    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host "    ALCANCE DEL ESCANEO" -ForegroundColor Cyan
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    [1] Unidad completa ($($DriveLetter):\)" -ForegroundColor White
    Write-Host "    [2] Carpeta especifica" -ForegroundColor White
    Write-Host ""

    do {
        $selection = Read-Host "  Seleccione una opcion (1 o 2)"
        $valid = $selection -in @('1','2')
        if (-not $valid) {
            Write-Host "  Seleccion invalida. Intente de nuevo." -ForegroundColor Red
        }
    } while (-not $valid)

    if ($selection -eq '1') {
        $scanPath = "$($DriveLetter):\"
    }
    else {
        # Intentar abrir dialogo de seleccion de carpeta
        $folderSelected = $false
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = "Seleccione la carpeta a escanear"
            $dialog.RootFolder = [System.Environment+SpecialFolder]::MyComputer
            $dialog.SelectedPath = "$($DriveLetter):\"
            $dialog.ShowNewFolderButton = $false

            $result = $dialog.ShowDialog()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $scanPath = $dialog.SelectedPath
                $folderSelected = $true
            }
        }
        catch {
            Write-Host "  No se pudo abrir el dialogo grafico." -ForegroundColor Yellow
        }

        if (-not $folderSelected) {
            Write-Host ""
            Write-Host "  Escriba la ruta completa de la carpeta a escanear:" -ForegroundColor White
            do {
                $scanPath = Read-Host "  Ruta"
                if (-not (Test-Path $scanPath -PathType Container)) {
                    Write-Host "  La ruta no existe o no es una carpeta. Intente de nuevo." -ForegroundColor Red
                    $scanPath = $null
                }
            } while (-not $scanPath)
        }
    }

    if (-not (Test-Path $scanPath)) {
        Write-Host "  ERROR: La ruta '$scanPath' no es accesible." -ForegroundColor Red
        exit 1
    }

    return $scanPath
}

# ============================================================================
# 5. RECOLECCION DE INFORMACION DEL EQUIPO
# ============================================================================

function Get-MachineInfo {
    $os   = Get-CimInstance Win32_OperatingSystem
    $cs   = Get-CimInstance Win32_ComputerSystem
    $cpu  = Get-CimInstance Win32_Processor | Select-Object -First 1
    $bios = Get-CimInstance Win32_BIOS
    $ram  = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum

    $networkAdapters = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' |
        ForEach-Object {
            [PSCustomObject]@{
                Description = $_.Description
                MACAddress  = $_.MACAddress
                IPAddress   = ($_.IPAddress -join ', ')
            }
        }

    # Usuario de ejecucion (puede ser admin si se elevo con "Ejecutar como administrador")
    $executionUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )

    # Usuario que inicio sesion en Windows (el real, no el elevado)
    # Obtener del proceso explorer.exe que siempre corre como el usuario de sesion
    $loggedOnUser = ""
    $loggedOnUserFull = ""
    try {
        $explorerProc = Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($explorerProc) {
            $ownerInfo = Invoke-CimMethod -InputObject $explorerProc -MethodName GetOwner -ErrorAction SilentlyContinue
            if ($ownerInfo -and $ownerInfo.User) {
                $loggedOnUser = "$($ownerInfo.Domain)\$($ownerInfo.User)"
                $sessionUsername = $ownerInfo.User
                $sessionDomain = $ownerInfo.Domain
            }
        }
    } catch {}

    # Fallback: usar variables de entorno originales de la sesion
    if ([string]::IsNullOrWhiteSpace($loggedOnUser)) {
        $loggedOnUser = "$env:USERDOMAIN\$env:USERNAME"
        $sessionUsername = $env:USERNAME
        $sessionDomain = $env:USERDOMAIN
    }

    # Nombre completo del usuario de sesion
    try {
        $userObj = Get-CimInstance Win32_UserAccount -Filter "Name='$sessionUsername'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($userObj -and $userObj.FullName) {
            $loggedOnUserFull = $userObj.FullName
        }
        if ([string]::IsNullOrWhiteSpace($loggedOnUserFull)) {
            $loggedOnUserFull = ([adsi]"WinNT://$sessionDomain/$sessionUsername,user").FullName
        }
    } catch {}
    if ([string]::IsNullOrWhiteSpace($loggedOnUserFull)) {
        $loggedOnUserFull = $sessionUsername
    }

    return [PSCustomObject]@{
        Hostname       = $env:COMPUTERNAME
        OSCaption      = $os.Caption
        OSVersion      = $os.Version
        OSBuild        = $os.BuildNumber
        OSArchitecture = $os.OSArchitecture
        Domain         = $cs.Domain
        DomainRole     = $cs.DomainRole
        LoggedOnUser     = $loggedOnUser
        LoggedOnUserFull = $loggedOnUserFull
        ExecutionUser    = $executionUser
        IsAdmin          = $isAdmin
        CPUName        = $cpu.Name
        CPUCores       = $cpu.NumberOfCores
        CPULogical     = $cpu.NumberOfLogicalProcessors
        RAMTotal       = $ram
        RAMFormatted   = Format-FileSize $ram
        BIOSManuf      = $bios.Manufacturer
        BIOSSerial     = $bios.SerialNumber
        BIOSVersion    = $bios.SMBIOSBIOSVersion
        BIOSDate       = $bios.ReleaseDate
        NetworkAdapters = $networkAdapters
        TimeZone       = [System.TimeZoneInfo]::Local.DisplayName
        TimeZoneId     = [System.TimeZoneInfo]::Local.Id
        ScanTimestamp  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss K')
    }
}

# ============================================================================
# 6. RECOLECCION DE INFORMACION DEL DISCO
# ============================================================================

function Get-DiskInfo {
    param([string]$DriveLetter)

    $volume = Get-Volume -DriveLetter $DriveLetter -ErrorAction SilentlyContinue

    # Obtener numero de disco desde la particion
    $partition = Get-Partition -DriveLetter $DriveLetter -ErrorAction SilentlyContinue
    $disk = $null
    $diskDrive = $null

    if ($partition) {
        $disk = Get-Disk -Number $partition.DiskNumber -ErrorAction SilentlyContinue
        $diskDrive = Get-CimInstance Win32_DiskDrive | Where-Object {
            $_.Index -eq $partition.DiskNumber
        } | Select-Object -First 1
    }

    $healthCounter = $null
    if ($disk) {
        try {
            $healthCounter = Get-StorageReliabilityCounter -Disk $disk -ErrorAction SilentlyContinue
        } catch {}
    }

    return [PSCustomObject]@{
        # Disco fisico
        DiskModel        = if ($disk) { $disk.Model } else { "N/A" }
        DiskSerial       = if ($disk) { $disk.SerialNumber } else { "N/A" }
        DiskSize         = if ($disk) { $disk.Size } else { 0 }
        DiskSizeFormatted = if ($disk) { Format-FileSize $disk.Size } else { "N/A" }
        DiskHealth       = if ($disk) { $disk.HealthStatus } else { "N/A" }
        DiskPartStyle    = if ($disk) { $disk.PartitionStyle } else { "N/A" }
        DiskBusType      = if ($disk) { $disk.BusType } else { "N/A" }
        DiskNumber       = if ($partition) { $partition.DiskNumber } else { "N/A" }
        InterfaceType    = if ($diskDrive) { $diskDrive.InterfaceType } else { "N/A" }
        MediaType        = if ($diskDrive) { $diskDrive.MediaType } else { "N/A" }
        Partitions       = if ($disk) { $disk.NumberOfPartitions } else { "N/A" }
        # Particion
        PartitionNumber  = if ($partition) { $partition.PartitionNumber } else { "N/A" }
        PartitionSize    = if ($partition) { Format-FileSize $partition.Size } else { "N/A" }
        PartitionType    = if ($partition) { $partition.Type } else { "N/A" }
        # Volumen
        VolumeLetter     = $DriveLetter
        VolumeLabel      = if ($volume) { $volume.FileSystemLabel } else { "N/A" }
        FileSystem       = if ($volume) { $volume.FileSystem } else { "N/A" }
        VolumeSize       = if ($volume) { $volume.Size } else { 0 }
        VolumeSizeFormatted = if ($volume) { Format-FileSize $volume.Size } else { "N/A" }
        VolumeFree       = if ($volume) { $volume.SizeRemaining } else { 0 }
        VolumeFreeFormatted = if ($volume) { Format-FileSize $volume.SizeRemaining } else { "N/A" }
        DriveType        = if ($volume) { $volume.DriveType } else { "N/A" }
        # SMART
        Temperature      = if ($healthCounter) { $healthCounter.Temperature } else { "N/A" }
        PowerOnHours     = if ($healthCounter) { $healthCounter.PowerOnHours } else { "N/A" }
        ReadErrors       = if ($healthCounter) { $healthCounter.ReadErrorsTotal } else { "N/A" }
    }
}

# ============================================================================
# 7. ESCANEO DE ARCHIVOS (STREAMING)
# ============================================================================

function Start-FileScan {
    param(
        [string]$ScanPath,
        [string]$TempCsvPath,
        [System.IO.StreamWriter]$LogWriter
    )

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $csvWriter = [System.IO.StreamWriter]::new($TempCsvPath, $false, $utf8Bom)

    # Escribir encabezado
    $header = @(
        "RutaCompleta",
        "Nombre",
        "Extension",
        "EsDirectorio",
        "EsOculto",
        "EsSistema",
        "EsSoloLectura",
        "TamanoBytes",
        "HashSHA256",
        "FechaCreacion",
        "FechaModificacion",
        "FechaAcceso",
        "Propietario",
        "Atributos"
    ) -join $CSV_DELIMITER

    $csvWriter.WriteLine($header)

    $stats = @{
        TotalFiles   = 0
        TotalDirs    = 0
        TotalErrors  = 0
        TotalSize    = [long]0
        ProcessedCount = 0
    }

    Write-Host ""
    Write-Host "  Iniciando escaneo de: $ScanPath" -ForegroundColor Cyan
    Write-Host "  (Presione Ctrl+C para cancelar)" -ForegroundColor DarkGray
    Write-Host ""

    $errorList = [System.Collections.Generic.List[string]]::new()

    Get-ChildItem -Path $ScanPath -Recurse -Force -ErrorVariable scanErrors -ErrorAction SilentlyContinue |
        Where-Object {
            # Saltar reparse points para evitar loops
            -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint)
        } |
        ForEach-Object {
            $item = $_
            $stats.ProcessedCount++

            # Progreso
            if ($stats.ProcessedCount % $PROGRESS_INTERVAL -eq 0) {
                Write-Progress -Activity "Escaneando archivos..." `
                    -Status "$($stats.ProcessedCount) elementos procesados | Archivos: $($stats.TotalFiles) | Carpetas: $($stats.TotalDirs)" `
                    -PercentComplete -1
            }

            $isDir    = $item.PSIsContainer
            $isHidden = [bool]($item.Attributes -band [IO.FileAttributes]::Hidden)
            $isSystem = [bool]($item.Attributes -band [IO.FileAttributes]::System)
            $isRO     = [bool]($item.Attributes -band [IO.FileAttributes]::ReadOnly)

            $sizeBytes = ""
            $hashValue = ""
            $owner     = ""

            if ($isDir) {
                $stats.TotalDirs++
                $sizeBytes = ""
                $hashValue = ""
            }
            else {
                $stats.TotalFiles++
                $sizeBytes = $item.Length
                $stats.TotalSize += $item.Length

                # Hash
                try {
                    $hashResult = Get-FileHash -Path $item.FullName -Algorithm $HASH_ALGORITHM -ErrorAction Stop
                    $hashValue = $hashResult.Hash
                }
                catch [System.UnauthorizedAccessException] {
                    $hashValue = "ACCESS_DENIED"
                    $stats.TotalErrors++
                }
                catch {
                    $hashValue = "ERROR"
                    $stats.TotalErrors++
                }
            }

            # Propietario
            try {
                $acl = Get-Acl -Path $item.FullName -ErrorAction Stop
                $owner = $acl.Owner
            }
            catch {
                $owner = "DESCONOCIDO"
            }

            # Fechas
            $creationTime  = $item.CreationTime.ToString('yyyy-MM-dd HH:mm:ss')
            $lastWriteTime = $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            $lastAccessTime = $item.LastAccessTime.ToString('yyyy-MM-dd HH:mm:ss')

            # Escribir linea CSV (escapar pipes en nombres si los hubiera)
            $fullPath = $item.FullName -replace '\|', '_'
            $name     = $item.Name -replace '\|', '_'
            $ext      = $item.Extension
            $attrs    = $item.Attributes.ToString()

            $line = @(
                $fullPath,
                $name,
                $ext,
                $(if ($isDir) { "SI" } else { "NO" }),
                $(if ($isHidden) { "SI" } else { "NO" }),
                $(if ($isSystem) { "SI" } else { "NO" }),
                $(if ($isRO) { "SI" } else { "NO" }),
                $sizeBytes,
                $hashValue,
                $creationTime,
                $lastWriteTime,
                $lastAccessTime,
                $owner,
                $attrs
            ) -join $CSV_DELIMITER

            $csvWriter.WriteLine($line)
        }

    Write-Progress -Activity "Escaneando archivos..." -Completed

    # Registrar errores de acceso a carpetas
    if ($scanErrors) {
        foreach ($err in $scanErrors) {
            $stats.TotalErrors++
            $errorList.Add($err.Exception.Message)
            if ($LogWriter) {
                $LogWriter.WriteLine("[ERROR] $($err.Exception.Message)")
            }
        }
    }

    $csvWriter.Flush()
    $csvWriter.Close()
    $csvWriter.Dispose()

    Write-Log -Message "Escaneo completado. Archivos: $($stats.TotalFiles), Carpetas: $($stats.TotalDirs), Errores: $($stats.TotalErrors)" -LogWriter $LogWriter

    return [PSCustomObject]@{
        TempCsvPath = $TempCsvPath
        TotalFiles  = $stats.TotalFiles
        TotalDirs   = $stats.TotalDirs
        TotalErrors = $stats.TotalErrors
        TotalSize   = $stats.TotalSize
        TotalSizeFormatted = Format-FileSize $stats.TotalSize
        ProcessedCount = $stats.ProcessedCount
        Errors      = $errorList
    }
}

# ============================================================================
# 8. EXPORTACION A EXCEL O CSV
# ============================================================================

function Export-ToExcelOrCsv {
    param(
        [string]$TempCsvPath,
        [string]$OutputDir,
        [string]$Timestamp,
        [PSCustomObject]$MachineInfo,
        [PSCustomObject]$DiskInfo,
        [long]$TotalRows,
        [System.IO.StreamWriter]$LogWriter
    )

    $useExcel = $false
    $outputPath = ""

    # Verificar si excede el limite de Excel
    if ($TotalRows -gt $EXCEL_MAX_ROWS) {
        Write-Log -Message "El numero de registros ($TotalRows) excede el limite de Excel ($EXCEL_MAX_ROWS). Se usara CSV." -Level Warning -LogWriter $LogWriter
    }
    else {
        # Verificar si ImportExcel esta disponible
        $moduleAvailable = Get-Module ImportExcel -ListAvailable -ErrorAction SilentlyContinue

        if (-not $moduleAvailable) {
            Write-Host ""
            Write-Host "  El modulo 'ImportExcel' no esta instalado." -ForegroundColor Yellow
            Write-Host "  [1] Instalar ImportExcel y generar Excel" -ForegroundColor White
            Write-Host "  [2] Continuar con CSV (delimitado por pipes)" -ForegroundColor White
            Write-Host ""
            $choice = Read-Host "  Seleccione una opcion (1 o 2)"

            if ($choice -eq '1') {
                try {
                    Write-Log -Message "Instalando modulo ImportExcel..." -LogWriter $LogWriter
                    Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
                    $moduleAvailable = $true
                    Write-Log -Message "ImportExcel instalado correctamente." -LogWriter $LogWriter
                }
                catch {
                    Write-Log -Message "No se pudo instalar ImportExcel: $($_.Exception.Message)" -Level Warning -LogWriter $LogWriter
                }
            }
        }

        if ($moduleAvailable) {
            $useExcel = $true
        }
    }

    if ($useExcel) {
        $outputPath = Join-Path $OutputDir "ScannerPerito_FileList_$Timestamp.xlsx"
        Write-Log -Message "Generando archivo Excel..." -LogWriter $LogWriter

        try {
            Import-Module ImportExcel -ErrorAction Stop

            $data = Import-Csv -Path $TempCsvPath -Delimiter $CSV_DELIMITER -Encoding UTF8

            $data | Export-Excel -Path $outputPath `
                -WorksheetName 'Archivos' `
                -AutoSize `
                -FreezeTopRow `
                -BoldTopRow `
                -TableStyle Medium2 `
                -ErrorAction Stop

            Write-Log -Message "Excel generado: $outputPath" -LogWriter $LogWriter
        }
        catch {
            Write-Log -Message "Error generando Excel: $($_.Exception.Message). Usando CSV como respaldo." -Level Warning -LogWriter $LogWriter
            $useExcel = $false
        }
    }

    if (-not $useExcel) {
        $outputPath = Join-Path $OutputDir "ScannerPerito_FileList_$Timestamp.csv"
        Write-Log -Message "Generando archivo CSV (delimitado por pipes)..." -LogWriter $LogWriter
        Copy-Item -Path $TempCsvPath -Destination $outputPath -Force
        Write-Log -Message "CSV generado: $outputPath" -LogWriter $LogWriter
    }

    return [PSCustomObject]@{
        OutputPath = $outputPath
        IsExcel    = $useExcel
        FileName   = Split-Path $outputPath -Leaf
    }
}

# ============================================================================
# 9. GENERACION DE INFORME PERICIAL TXT
# ============================================================================

function New-ForensicReport {
    param(
        [string]$OutputDir,
        [string]$Timestamp,
        [PSCustomObject]$Config,
        [PSCustomObject]$MachineInfo,
        [PSCustomObject]$DiskInfo,
        [string]$ScanPath,
        [string]$StartTime,
        [string]$EndTime,
        [PSCustomObject]$ScanStats,
        [PSCustomObject]$ExportResult,
        [string]$DataFileHash,
        [string]$CaseId,
        [string]$CaseDescription
    )

    $reportPath = Join-Path $OutputDir "ScannerPerito_Report_$Timestamp.txt"
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $writer = [System.IO.StreamWriter]::new($reportPath, $false, $utf8Bom)

    $separator = "=" * 80
    $subSeparator = "-" * 80

    # --- CABECERA ---
    $writer.WriteLine($separator)
    $writer.WriteLine("                    INFORME PERICIAL INFORMATICO")
    $writer.WriteLine("                    EVIDENCIA DIGITAL - LISTADO DE ARCHIVOS")
    $writer.WriteLine($separator)
    $writer.WriteLine("")
    $writer.WriteLine("ENTIDAD:             $($Config.entity_name)")
    $writer.WriteLine("NIT:                 $($Config.nit)")
    $writer.WriteLine("DIRECCION:           $($Config.address)")
    $writer.WriteLine("CIUDAD:              $($Config.city)")
    $writer.WriteLine("TELEFONO:            $($Config.phone)")
    $writer.WriteLine("EMAIL:               $($Config.email)")
    $writer.WriteLine("")
    $writer.WriteLine("PERITO / AUTOR:      $($Config.author)")
    $writer.WriteLine("IDENTIFICACION:      $($Config.author_id)")
    $writer.WriteLine("")
    $writer.WriteLine("EJECUTADO POR:")
    $writer.WriteLine("  Sesion Windows:    $($MachineInfo.LoggedOnUser)")
    $writer.WriteLine("  Nombre Completo:   $($MachineInfo.LoggedOnUserFull)")
    $writer.WriteLine("  Cuenta Ejecucion:  $($MachineInfo.ExecutionUser)")
    $writer.WriteLine("  Administrador:     $(if ($MachineInfo.IsAdmin) { 'SI' } else { 'NO' })")
    $writer.WriteLine("")
    $writer.WriteLine("CASO No.:            $CaseId")
    $writer.WriteLine("DESCRIPCION:         $CaseDescription")
    $writer.WriteLine("")

    # --- INFORMACION DEL EQUIPO ---
    $writer.WriteLine($subSeparator)
    $writer.WriteLine("                    INFORMACION DEL EQUIPO EXAMINADO")
    $writer.WriteLine($subSeparator)
    $writer.WriteLine("")
    $writer.WriteLine("Hostname:            $($MachineInfo.Hostname)")
    $writer.WriteLine("Sistema Operativo:   $($MachineInfo.OSCaption)")
    $writer.WriteLine("Version:             $($MachineInfo.OSVersion) (Build $($MachineInfo.OSBuild))")
    $writer.WriteLine("Arquitectura:        $($MachineInfo.OSArchitecture)")
    $writer.WriteLine("Dominio:             $($MachineInfo.Domain)")
    $writer.WriteLine("Usuario Sesion:      $($MachineInfo.LoggedOnUser) ($($MachineInfo.LoggedOnUserFull))")
    $writer.WriteLine("Usuario Ejecucion:   $($MachineInfo.ExecutionUser)")
    $writer.WriteLine("Ejecutado como Admin:$(if ($MachineInfo.IsAdmin) { ' SI' } else { ' NO' })")
    $writer.WriteLine("CPU:                 $($MachineInfo.CPUName)")
    $writer.WriteLine("  Nucleos Fisicos:   $($MachineInfo.CPUCores)")
    $writer.WriteLine("  Nucleos Logicos:   $($MachineInfo.CPULogical)")
    $writer.WriteLine("RAM Total:           $($MachineInfo.RAMFormatted)")
    $writer.WriteLine("BIOS Fabricante:     $($MachineInfo.BIOSManuf)")
    $writer.WriteLine("BIOS Serial:         $($MachineInfo.BIOSSerial)")
    $writer.WriteLine("BIOS Version:        $($MachineInfo.BIOSVersion)")
    $writer.WriteLine("")
    $writer.WriteLine("Interfaces de Red:")
    foreach ($adapter in $MachineInfo.NetworkAdapters) {
        $writer.WriteLine("  - $($adapter.Description)")
        $writer.WriteLine("    MAC: $($adapter.MACAddress)  |  IP: $($adapter.IPAddress)")
    }
    $writer.WriteLine("")
    $writer.WriteLine("Zona Horaria:        $($MachineInfo.TimeZone)")
    $writer.WriteLine("")

    # --- INFORMACION DEL DISCO ---
    $writer.WriteLine($subSeparator)
    $writer.WriteLine("                    INFORMACION DEL DISCO / VOLUMEN")
    $writer.WriteLine($subSeparator)
    $writer.WriteLine("")
    $writer.WriteLine("Disco Fisico:")
    $writer.WriteLine("  Modelo:            $($DiskInfo.DiskModel)")
    $writer.WriteLine("  Numero de Serie:   $($DiskInfo.DiskSerial)")
    $writer.WriteLine("  Tamano:            $($DiskInfo.DiskSizeFormatted)")
    $writer.WriteLine("  Interfaz:          $($DiskInfo.DiskBusType) / $($DiskInfo.InterfaceType)")
    $writer.WriteLine("  Tipo de Medio:     $($DiskInfo.MediaType)")
    $writer.WriteLine("  Estado de Salud:   $($DiskInfo.DiskHealth)")
    $writer.WriteLine("  Estilo Particion:  $($DiskInfo.DiskPartStyle)")
    $writer.WriteLine("  Num. Particiones:  $($DiskInfo.Partitions)")
    $writer.WriteLine("  Numero de Disco:   $($DiskInfo.DiskNumber)")
    $writer.WriteLine("")
    $writer.WriteLine("  Temperatura:       $($DiskInfo.Temperature)")
    $writer.WriteLine("  Horas Encendido:   $($DiskInfo.PowerOnHours)")
    $writer.WriteLine("  Errores Lectura:   $($DiskInfo.ReadErrors)")
    $writer.WriteLine("")
    $writer.WriteLine("Volumen:")
    $writer.WriteLine("  Letra:             $($DiskInfo.VolumeLetter):")
    $writer.WriteLine("  Etiqueta:          $($DiskInfo.VolumeLabel)")
    $writer.WriteLine("  Sistema Archivos:  $($DiskInfo.FileSystem)")
    $writer.WriteLine("  Tamano:            $($DiskInfo.VolumeSizeFormatted)")
    $writer.WriteLine("  Espacio Libre:     $($DiskInfo.VolumeFreeFormatted)")
    $writer.WriteLine("  Tipo Unidad:       $($DiskInfo.DriveType)")
    $writer.WriteLine("")

    # --- PARAMETROS DEL ESCANEO ---
    $writer.WriteLine($subSeparator)
    $writer.WriteLine("                    PARAMETROS DEL ESCANEO")
    $writer.WriteLine($subSeparator)
    $writer.WriteLine("")
    $writer.WriteLine("Ruta Escaneada:      $ScanPath")
    $writer.WriteLine("Fecha/Hora Inicio:   $StartTime")
    $writer.WriteLine("Fecha/Hora Fin:      $EndTime")
    $writer.WriteLine("Zona Horaria:        $($MachineInfo.TimeZoneId)")
    $writer.WriteLine("Algoritmo Hash:      $HASH_ALGORITHM")
    $writer.WriteLine("")
    $writer.WriteLine("Resultados:")
    $writer.WriteLine("  Total Archivos:    $($ScanStats.TotalFiles)")
    $writer.WriteLine("  Total Directorios: $($ScanStats.TotalDirs)")
    $writer.WriteLine("  Total Procesados:  $($ScanStats.ProcessedCount)")
    $writer.WriteLine("  Total Errores:     $($ScanStats.TotalErrors)")
    $writer.WriteLine("  Tamano Total:      $($ScanStats.TotalSizeFormatted)")
    $writer.WriteLine("")

    # --- ARCHIVOS GENERADOS ---
    $writer.WriteLine($subSeparator)
    $writer.WriteLine("                    ARCHIVOS DE EVIDENCIA GENERADOS")
    $writer.WriteLine($subSeparator)
    $writer.WriteLine("")
    $dataFileSize = (Get-Item $ExportResult.OutputPath).Length
    $writer.WriteLine("Archivo de Datos:    $($ExportResult.FileName)")
    $writer.WriteLine("  Tipo:              $(if ($ExportResult.IsExcel) { 'Excel (.xlsx)' } else { 'CSV delimitado por pipes (.csv)' })")
    $writer.WriteLine("  Tamano:            $(Format-FileSize $dataFileSize)")
    $writer.WriteLine("  $($HASH_ALGORITHM):           $DataFileHash")
    $writer.WriteLine("")
    $writer.WriteLine("Informe (este archivo): ScannerPerito_Report_$Timestamp.txt")
    $writer.WriteLine("")

    # --- NOTAS ---
    if (-not [string]::IsNullOrWhiteSpace($Config.notes)) {
        $writer.WriteLine($subSeparator)
        $writer.WriteLine("                    NOTAS")
        $writer.WriteLine($subSeparator)
        $writer.WriteLine("")
        $writer.WriteLine($Config.notes)
        $writer.WriteLine("")
    }

    # --- PIE ---
    $writer.WriteLine($separator)
    $writer.WriteLine("Generado por ScannerPerito v$SCRIPT_VERSION")
    $writer.WriteLine("Fecha de generacion: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss K')")
    $writer.WriteLine($separator)

    $writer.Flush()
    $writer.Close()
    $writer.Dispose()

    # Calcular hash del propio informe y anadirlo al final
    $reportHash = (Get-FileHash -Path $reportPath -Algorithm $HASH_ALGORITHM).Hash
    Add-Content -Path $reportPath -Value "" -Encoding UTF8
    Add-Content -Path $reportPath -Value "$HASH_ALGORITHM de este informe: $reportHash" -Encoding UTF8

    return [PSCustomObject]@{
        ReportPath = $reportPath
        ReportHash = $reportHash
        FileName   = Split-Path $reportPath -Leaf
    }
}

# ============================================================================
# BLOQUE PRINCIPAL (MAIN)
# ============================================================================

# Banner
Clear-Host
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host "    ____                                  ____           _ _        " -ForegroundColor Cyan
Write-Host "   / ___|  ___ __ _ _ __  _ __   ___ _ __|  _ \ ___ _ __(_) |_ ___  " -ForegroundColor Cyan
Write-Host "   \___ \ / __/ _`` | '_ \| '_ \ / _ \ '__| |_) / _ \ '__| | __/ _ \ " -ForegroundColor Cyan
Write-Host "    ___) | (_| (_| | | | | | | |  __/ |  |  __/  __/ |  | | || (_) |" -ForegroundColor Cyan
Write-Host "   |____/ \___\__,_|_| |_|_| |_|\___|_|  |_|   \___|_|  |_|\__\___/ " -ForegroundColor Cyan
Write-Host "" -ForegroundColor Cyan
Write-Host "    Generador de Evidencia Digital para Peritaje Informatico" -ForegroundColor White
Write-Host "    Version $SCRIPT_VERSION" -ForegroundColor DarkGray
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Validar compatibilidad
$osInfo = Test-OSCompatibility
Write-Host "  SO detectado: $($osInfo.OSCaption) (Build $($osInfo.BuildNumber))" -ForegroundColor Green
Write-Host ""

# 2. Cargar configuracion
$config = Read-Configuration

# 2b. Generar numero de caso
$caseInfo = Get-NextCaseNumber

# 2c. Pedir descripcion del caso
$caseDescription = Read-CaseDescription -CaseId $caseInfo.CaseId

# 3. Seleccion de unidad
$driveLetter = Show-DriveSelectionMenu

# 4. Obtener info del disco
$diskInfo = Get-DiskInfo -DriveLetter $driveLetter

# 5. Seleccion de alcance
$scanPath = Show-ScopeSelectionMenu -DriveLetter $driveLetter

# 6. Obtener info del equipo
Write-Host ""
Write-Host "  Recopilando informacion del equipo..." -ForegroundColor Cyan
$machineInfo = Get-MachineInfo

# 7. Confirmar parametros
Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "    RESUMEN ANTES DE INICIAR" -ForegroundColor Cyan
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Caso:       $($caseInfo.CaseId)" -ForegroundColor White
Write-Host "  Descripcion:" -ForegroundColor White
Write-Host "    $caseDescription" -ForegroundColor DarkGray
Write-Host "  Autor:      $($config.author)" -ForegroundColor White
Write-Host "  Equipo:     $($machineInfo.Hostname)" -ForegroundColor White
Write-Host "  Disco:      $($diskInfo.DiskModel) ($($diskInfo.DiskSizeFormatted))" -ForegroundColor White
Write-Host "  Volumen:    $($diskInfo.VolumeLetter): - $($diskInfo.VolumeLabel) ($($diskInfo.FileSystem))" -ForegroundColor White
Write-Host "  Ruta:       $scanPath" -ForegroundColor White
Write-Host "  Algoritmo:  $HASH_ALGORITHM" -ForegroundColor White
Write-Host ""

$confirm = Read-Host "  Desea iniciar el escaneo? (S/N)"
if ($confirm -notin @('S','s','SI','si','Si','Y','y','YES','yes')) {
    Write-Host ""
    Write-Host "  Operacion cancelada por el usuario." -ForegroundColor Yellow
    exit 0
}

# 8. Persistir numero de caso y preparar directorio de salida
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
Save-CaseNumber -SeedPath $caseInfo.SeedPath `
    -CaseNumber $caseInfo.CaseNumber `
    -CaseId $caseInfo.CaseId `
    -CaseDescription $caseDescription `
    -Timestamp (Get-Date -Format 'yyyy-MM-dd HH:mm:ss K')

$outputDir = Join-Path $PSScriptRoot "ScannerPerito_$($caseInfo.CaseId)_$timestamp"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

# Iniciar log
$logPath = Join-Path $outputDir "ScannerPerito_Execution.log"
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$logWriter = [System.IO.StreamWriter]::new($logPath, $false, $utf8Bom)
$logWriter.WriteLine("ScannerPerito v$SCRIPT_VERSION - Log de Ejecucion")
$logWriter.WriteLine("Inicio: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss K')")
$logWriter.WriteLine("")

# Verificar si es admin
if (-not $machineInfo.IsAdmin) {
    Write-Host ""
    Write-Host "  AVISO: No se esta ejecutando como Administrador." -ForegroundColor Yellow
    Write-Host "  Algunos archivos o carpetas podrian no ser accesibles." -ForegroundColor Yellow
    Write-Host ""
    Write-Log -Message "Ejecutando sin privilegios de administrador." -Level Warning -LogWriter $logWriter
}

# 9. Escaneo
$startTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss K'
$startDateTime = Get-Date

Write-Log -Message "Iniciando escaneo de: $scanPath" -LogWriter $logWriter

$tempCsvPath = Join-Path $outputDir "temp_scan.csv"
$scanStats = Start-FileScan -ScanPath $scanPath -TempCsvPath $tempCsvPath -LogWriter $logWriter

$endTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss K'
$endDateTime = Get-Date
$duration = $endDateTime - $startDateTime

Write-Host ""
Write-Host "  Escaneo completado en $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor Green
Write-Host "  Archivos: $($scanStats.TotalFiles) | Carpetas: $($scanStats.TotalDirs) | Errores: $($scanStats.TotalErrors)" -ForegroundColor White
Write-Host "  Tamano total: $($scanStats.TotalSizeFormatted)" -ForegroundColor White
Write-Host ""

Write-Log -Message "Duracion del escaneo: $($duration.ToString('hh\:mm\:ss'))" -LogWriter $logWriter

# 10. Exportar a Excel o CSV
$totalRows = $scanStats.TotalFiles + $scanStats.TotalDirs
$exportResult = Export-ToExcelOrCsv -TempCsvPath $tempCsvPath `
    -OutputDir $outputDir `
    -Timestamp $timestamp `
    -MachineInfo $machineInfo `
    -DiskInfo $diskInfo `
    -TotalRows $totalRows `
    -LogWriter $logWriter

# 11. Hash del archivo de datos
Write-Host "  Calculando hash del archivo de datos..." -ForegroundColor Cyan
$dataFileHash = (Get-FileHash -Path $exportResult.OutputPath -Algorithm $HASH_ALGORITHM).Hash
Write-Log -Message "Hash del archivo de datos ($HASH_ALGORITHM): $dataFileHash" -LogWriter $logWriter

# 12. Generar informe TXT
Write-Host "  Generando informe pericial..." -ForegroundColor Cyan
$reportResult = New-ForensicReport `
    -OutputDir $outputDir `
    -Timestamp $timestamp `
    -Config $config `
    -MachineInfo $machineInfo `
    -DiskInfo $diskInfo `
    -ScanPath $scanPath `
    -StartTime $startTime `
    -EndTime $endTime `
    -ScanStats $scanStats `
    -ExportResult $exportResult `
    -DataFileHash $dataFileHash `
    -CaseId $caseInfo.CaseId `
    -CaseDescription $caseDescription

Write-Log -Message "Informe generado: $($reportResult.ReportPath)" -LogWriter $logWriter
Write-Log -Message "Hash del informe ($HASH_ALGORITHM): $($reportResult.ReportHash)" -LogWriter $logWriter

# 13. Limpiar archivo temporal
Remove-Item -Path $tempCsvPath -Force -ErrorAction SilentlyContinue

# 14. Cerrar log
$logWriter.WriteLine("")
$logWriter.WriteLine("Fin: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss K')")
$logWriter.Flush()
$logWriter.Close()
$logWriter.Dispose()

# 15. Resumen final
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor Green
Write-Host "    PROCESO COMPLETADO EXITOSAMENTE" -ForegroundColor Green
Write-Host "  ================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Archivos generados en:" -ForegroundColor White
Write-Host "  $outputDir" -ForegroundColor Yellow
Write-Host ""
Write-Host "  - $($exportResult.FileName)" -ForegroundColor White
Write-Host "    $HASH_ALGORITHM : $dataFileHash" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  - $($reportResult.FileName)" -ForegroundColor White
Write-Host "    $HASH_ALGORITHM : $($reportResult.ReportHash)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  - ScannerPerito_Execution.log" -ForegroundColor White
Write-Host ""
Write-Host "  Duracion total: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host "  ================================================================" -ForegroundColor Green
Write-Host ""

# Abrir carpeta de salida
$openFolder = Read-Host "  Desea abrir la carpeta de resultados? (S/N)"
if ($openFolder -in @('S','s','SI','si','Si','Y','y')) {
    Start-Process explorer.exe -ArgumentList $outputDir
}
