[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Test', 'Install')]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$ComputerList,

    [string]$SourceFolder,
    [string]$RemoteSourceFolder,
    [string]$Installer,
    [string]$SilentArgs,
    [string]$PsExecPath,
    [string]$DestinationName = 'PacoteInstalacao',

    [ValidateRange(1, 5)]
    [int]$ThrottleLimit = 4,

    [ValidateSet('ExitCode', 'File', 'Service')]
    [string]$ValidationType = 'ExitCode',
    [string]$ValidationValue = ''
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

function Write-MainStatus {
    param([string]$Computer, [string]$Stage, [string]$Result, [string]$Details = '')
    $line = '[{0}] [{1}] {2}: {3}' -f (Get-Date -Format 'HH:mm:ss'), $Computer, $Stage, $Result
    if ($Details) { $line += ' - ' + $Details }
    Write-Output $line
}

function Get-ComputerNames {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Lista nao encontrada: $Path" }

    if ([IO.Path]::GetExtension($Path) -ieq '.csv') {
        $header = Get-Content -LiteralPath $Path -TotalCount 1
        if ([string]::IsNullOrWhiteSpace($header)) { return @() }
        $semicolonCount = ($header.ToCharArray() | Where-Object { $_ -eq ';' }).Count
        $commaCount = ($header.ToCharArray() | Where-Object { $_ -eq ',' }).Count
        $delimiter = if ($semicolonCount -gt $commaCount) { ';' } else { ',' }
        $rows = @(Import-Csv -LiteralPath $Path -Delimiter $delimiter)
        if ($rows.Count -eq 0) { return @() }
        $properties = @($rows[0].PSObject.Properties.Name)
        $preferred = @('Patrimonio', 'Patrimônio', 'Computer', 'Computador', 'Hostname', 'Host', 'Name', 'Nome')
        $column = $preferred | Where-Object { $properties -contains $_ } | Select-Object -First 1
        if (-not $column) { $column = $properties[0] }
        $values = $rows | ForEach-Object { [string]$_.$column }
    }
    else {
        $values = Get-Content -LiteralPath $Path
    }

    $valid = New-Object System.Collections.Generic.List[string]
    foreach ($value in $values) {
        $name = ([string]$value).Trim().ToUpperInvariant()
        if (-not $name) { continue }
        if ($name -match '^\d+$') { $name = 'W' + $name }
        if ($name -notmatch '^W\d+$') {
            Write-MainStatus $name 'VALIDACAO' 'IGNORADO' 'Use W seguido de numeros'
            continue
        }
        if (-not $valid.Contains($name)) { $valid.Add($name) }
    }
    return $valid.ToArray()
}

$worker = {
    param(
        [string]$Computer,
        [string]$PsExec,
        [string]$RemoteSource,
        [string]$RemoteLocal,
        [string]$RemoteShare,
        [string]$InstallerRelative,
        [string]$InstallerExtension,
        [string]$Arguments,
        [string]$CheckType,
        [string]$CheckValue,
        [string]$OperationMode
    )

    $ErrorActionPreference = 'Stop'
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $result = [ordered]@{
        RecordType = 'Result'; Computer = $Computer; Access = 'Pendente'; Copy = 'Nao executada'
        Install = 'Nao executada'; ExitCode = ''; Validation = 'Nao executada'
        Cleanup = 'Nao executada'; DurationSeconds = 0; ErrorReason = ''; Overall = 'Falha'
    }

    function Status([string]$Stage, [string]$State, [string]$Details = '') {
        $line = '[{0}] [{1}] {2}: {3}' -f (Get-Date -Format 'HH:mm:ss'), $Computer, $Stage, $State
        if ($Details) { $line += ' - ' + $Details }
        Write-Output $line
    }

    function CompleteResult {
        $timer.Stop()
        $result.DurationSeconds = [math]::Round($timer.Elapsed.TotalSeconds, 1)
        return [pscustomobject]$result
    }

    function Invoke-PsExecSafe {
        param([string[]]$CommandArguments)

        $savedPreference = $ErrorActionPreference
        $nativeOutput = @()
        $nativeExitCode = -1
        try {
            # O PsExec escreve mensagens normais de conexao no stderr. Com
            # ErrorActionPreference=Stop, o PowerShell pode trata-las como
            # excecao antes de o processo remoto terminar.
            $ErrorActionPreference = 'Continue'
            $nativeOutput = @(& $PsExec @CommandArguments 2>&1)
            $nativeExitCode = $LASTEXITCODE
        }
        catch {
            $nativeOutput += $_.Exception.Message
        }
        finally {
            $ErrorActionPreference = $savedPreference
        }

        return [pscustomobject]@{
            ExitCode = $nativeExitCode
            Output = $nativeOutput
        }
    }

    function Get-NativeOutputSummary {
        param([object[]]$NativeOutput)
        return (@(
            $NativeOutput |
                ForEach-Object { ([string]$_).Trim() } |
                Where-Object { $_ } |
                Select-Object -Last 5
        ) -join ' | ')
    }

    try {
        if (-not (Test-Path -LiteralPath "\\$Computer\C$" -ErrorAction SilentlyContinue)) {
            $result.Access = 'Indisponivel'
            $result.ErrorReason = 'Computador offline, sem permissao administrativa ou C$ bloqueado'
            Status 'ACESSO' 'ERRO' $result.ErrorReason
            return (CompleteResult)
        }
        $result.Access = 'OK'
        Status 'ACESSO' 'OK'

        if ($OperationMode -eq 'Test') {
            $result.Overall = 'Sucesso'
            return (CompleteResult)
        }

        try {
            New-Item -ItemType Directory -Path $RemoteShare -Force | Out-Null
            Status 'COPIA' 'INICIANDO' "$RemoteSource -> $RemoteLocal"
            $copyRun = Invoke-PsExecSafe -CommandArguments @(
                "\\$Computer", '-accepteula', '-nobanner', '-h', '-s',
                'robocopy.exe', $RemoteSource, $RemoteLocal,
                '/E', '/Z', '/R:2', '/W:2', '/NP', '/NFL', '/NDL'
            )
            $copyCode = $copyRun.ExitCode
            $copyDetails = Get-NativeOutputSummary -NativeOutput $copyRun.Output
            $copiedInstallerShare = Join-Path $RemoteShare $InstallerRelative
            $copiedInstaller = $null
            for ($copyCheckAttempt = 1; $copyCheckAttempt -le 5 -and -not $copiedInstaller; $copyCheckAttempt++) {
                $copiedInstaller = Get-Item -LiteralPath $copiedInstallerShare -ErrorAction SilentlyContinue
                if (-not $copiedInstaller -and $copyCheckAttempt -lt 5) {
                    Start-Sleep -Milliseconds 500
                }
            }
            if (-not $copiedInstaller -or $copiedInstaller.PSIsContainer -or $copiedInstaller.Length -le 0) {
                throw "O instalador nao foi encontrado no destino apos a copia: $RemoteLocal. Retorno $copyCode. $copyDetails"
            }

            $copiedSizeMB = [math]::Round($copiedInstaller.Length / 1MB, 2)
            if ($copyCode -ge 0 -and $copyCode -le 7) {
                $result.Copy = "OK ($copyCode)"
                Status 'COPIA' 'OK' "Instalador confirmado no destino; $copiedSizeMB MB; Robocopy $copyCode"
            }
            else {
                # O arquivo no destino e a fonte de verdade. Alguns retornos do
                # PsExec nao representam falha do Robocopy.
                $result.Copy = 'OK (validado)'
                Status 'COPIA' 'OK' "Instalador confirmado no destino; $copiedSizeMB MB; retorno $copyCode ignorado"
            }
        }
        catch {
            $result.Copy = 'Erro'
            $result.ErrorReason = $_.Exception.Message
            Status 'COPIA' 'ERRO' $result.ErrorReason
            return (CompleteResult)
        }

        $remoteInstaller = Join-Path $RemoteLocal $InstallerRelative
        $remoteCmdShare = Join-Path $RemoteShare '__instalacao_remota.cmd'
        if ($InstallerExtension -eq '.msi') {
            $commandLine = 'msiexec.exe /i "{0}" {1}' -f $remoteInstaller, $Arguments
        }
        else {
            $commandLine = '"{0}" {1}' -f $remoteInstaller, $Arguments
        }
        [IO.File]::WriteAllText(
            $remoteCmdShare,
            "@echo off`r`n$commandLine`r`nexit /b %errorlevel%`r`n",
            [Text.Encoding]::Default
        )

        Status 'INSTALACAO' 'INICIANDO' ([IO.Path]::GetFileName($remoteInstaller))
        $installRun = Invoke-PsExecSafe -CommandArguments @(
            "\\$Computer", '-accepteula', '-nobanner', '-h', '-s',
            'cmd.exe', '/d', '/c', "$RemoteLocal\__instalacao_remota.cmd"
        )
        $installCode = $installRun.ExitCode
        $result.ExitCode = $installCode
        if ($installCode -eq 0) {
            $result.Install = 'Sucesso'
            Status 'INSTALACAO' 'SUCESSO' 'Codigo 0'
        }
        elseif ($installCode -in @(1641, 3010)) {
            $result.Install = 'Sucesso; reinicio necessario'
            Status 'INSTALACAO' 'SUCESSO' "Codigo $installCode; reinicio necessario"
        }
        else {
            $result.Install = 'Erro'
            $installDetails = Get-NativeOutputSummary -NativeOutput $installRun.Output
            $result.ErrorReason = "Instalador/PsExec retornou codigo $installCode. $installDetails"
            Status 'INSTALACAO' 'ERRO' $result.ErrorReason
            $result.Cleanup = 'Mantida para diagnostico'
            Status 'LIMPEZA' 'MANTIDA' 'Arquivos preservados para diagnostico'
            return (CompleteResult)
        }

        if ($CheckType -eq 'File') {
            if ([string]::IsNullOrWhiteSpace($CheckValue) -or $CheckValue -notmatch '^[Cc]:\\') {
                $result.Validation = 'Erro de configuracao'
                $result.ErrorReason = 'Informe um caminho iniciado por C:\ para validar por arquivo'
            }
            else {
                $checkShare = "\\$Computer\C$\" + $CheckValue.Substring(3)
                if (Test-Path -LiteralPath $checkShare -PathType Leaf -ErrorAction SilentlyContinue) {
                    $result.Validation = 'OK'
                }
                else {
                    $result.Validation = 'Falha'
                    $result.ErrorReason = "Arquivo de validacao nao encontrado: $CheckValue"
                }
            }
        }
        elseif ($CheckType -eq 'Service') {
            if ([string]::IsNullOrWhiteSpace($CheckValue)) {
                $result.Validation = 'Erro de configuracao'
                $result.ErrorReason = 'Informe o nome interno do servico para validacao'
            }
            else {
                $serviceRun = Invoke-PsExecSafe -CommandArguments @(
                    "\\$Computer", '-accepteula', '-nobanner', '-h', '-s',
                    'sc.exe', 'query', $CheckValue
                )
                if ($serviceRun.ExitCode -eq 0) { $result.Validation = 'OK' }
                else {
                    $result.Validation = 'Falha'
                    $serviceDetails = Get-NativeOutputSummary -NativeOutput $serviceRun.Output
                    $result.ErrorReason = "Servico nao encontrado: $CheckValue. Codigo $($serviceRun.ExitCode). $serviceDetails"
                }
            }
        }
        else {
            $result.Validation = 'Codigo de saida OK'
        }

        if ($result.Validation -in @('OK', 'Codigo de saida OK')) {
            Status 'VALIDACAO' 'OK' $result.Validation
            try {
                Remove-Item -LiteralPath $RemoteShare -Recurse -Force
                $result.Cleanup = 'OK'
                Status 'LIMPEZA' 'OK' "$RemoteLocal removida"
            }
            catch {
                $result.Cleanup = 'Erro'
                if (-not $result.ErrorReason) { $result.ErrorReason = $_.Exception.Message }
                Status 'LIMPEZA' 'ERRO' $_.Exception.Message
            }
            $result.Overall = 'Sucesso'
        }
        else {
            Status 'VALIDACAO' 'ERRO' $result.ErrorReason
            $result.Cleanup = 'Mantida para diagnostico'
            Status 'LIMPEZA' 'MANTIDA' 'Validacao falhou; arquivos preservados'
        }
    }
    catch {
        if (-not $result.ErrorReason) { $result.ErrorReason = $_.Exception.Message }
        Status 'GERAL' 'ERRO' $result.ErrorReason
    }
    return (CompleteResult)
}

try {
    $computers = @(Get-ComputerNames -Path $ComputerList)
    if ($computers.Count -eq 0) { throw 'A lista nao possui patrimonios validos.' }
    Write-Output ('Patrimonios validos: {0}; simultaneos: {1}' -f $computers.Count, $ThrottleLimit)

    $sourceFull = ''
    $remoteSourceFull = ''
    $installerRelative = ''
    $extension = ''
    $safeDestination = $DestinationName.Trim()

    if ($Mode -eq 'Install') {
        foreach ($required in @($SourceFolder, $RemoteSourceFolder, $Installer, $PsExecPath)) {
            if ([string]::IsNullOrWhiteSpace($required)) { throw 'Preencha pasta, origem remota, instalador e PsExec.' }
        }
        if (-not (Test-Path -LiteralPath $SourceFolder -PathType Container)) { throw "Pasta de origem nao encontrada: $SourceFolder" }
        if (-not (Test-Path -LiteralPath $Installer -PathType Leaf)) { throw "Instalador nao encontrado: $Installer" }
        if (-not (Test-Path -LiteralPath $PsExecPath -PathType Leaf)) { throw "PsExec nao encontrado: $PsExecPath" }
        if ([IO.Path]::GetFileName($PsExecPath) -notmatch '^PsExec(64)?\.exe$') { throw 'Selecione PsExec.exe ou PsExec64.exe.' }
        if (-not $RemoteSourceFolder.StartsWith('\\')) { throw 'A origem remota precisa ser um caminho UNC.' }

        $sourceFull = [IO.Path]::GetFullPath($SourceFolder).TrimEnd('\')
        $remoteSourceFull = $RemoteSourceFolder.TrimEnd('\')
        $installerFull = [IO.Path]::GetFullPath($Installer)
        if (-not $installerFull.StartsWith($sourceFull + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'O instalador precisa estar dentro da pasta de origem.'
        }
        $installerRelative = $installerFull.Substring($sourceFull.Length + 1)
        $extension = [IO.Path]::GetExtension($installerFull).ToLowerInvariant()
        if ($extension -notin @('.exe', '.msi')) { throw 'O instalador deve possuir extensao .exe ou .msi.' }
        if (-not $safeDestination) { $safeDestination = [IO.Path]::GetFileName($sourceFull) }
        if ($safeDestination.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $safeDestination.Contains('\') -or $safeDestination.Contains('/')) {
            throw 'O nome da pasta de destino possui caracteres invalidos.'
        }
        if ($extension -eq '.msi' -and [string]::IsNullOrWhiteSpace($SilentArgs)) { $SilentArgs = '/qn /norestart' }
    }

    $queue = New-Object System.Collections.Generic.Queue[string]
    foreach ($computer in $computers) { $queue.Enqueue($computer) }
    $active = New-Object System.Collections.ArrayList
    $summary = New-Object System.Collections.Generic.List[object]
    $receivedResults = New-Object System.Collections.Generic.HashSet[string]

    while ($queue.Count -gt 0 -or $active.Count -gt 0) {
        while ($queue.Count -gt 0 -and $active.Count -lt $ThrottleLimit) {
            $computer = $queue.Dequeue()
            Write-MainStatus $computer 'FILA' 'INICIANDO'
            $remoteShare = "\\$computer\C$\Temp\$safeDestination"
            $remoteLocal = "C:\Temp\$safeDestination"
            $job = Start-Job -Name $computer -ScriptBlock $worker -ArgumentList @(
                $computer, $PsExecPath, $remoteSourceFull, $remoteLocal, $remoteShare,
                $installerRelative, $extension, $SilentArgs, $ValidationType,
                $ValidationValue, $Mode
            )
            [void]$active.Add($job)
        }

        foreach ($job in @($active)) {
            $items = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
            foreach ($item in $items) {
                if ($item.PSObject.Properties['RecordType'] -and $item.RecordType -eq 'Result') {
                    $summary.Add($item)
                    [void]$receivedResults.Add([string]$item.Computer)
                    $finalDetails = 'Tempo {0}s' -f $item.DurationSeconds
                    if ($item.ErrorReason) { $finalDetails += '; ' + $item.ErrorReason }
                    Write-MainStatus $item.Computer 'FINAL' $item.Overall $finalDetails
                }
                else { Write-Output ([string]$item) }
            }

            if ($job.State -in @('Completed', 'Failed', 'Stopped')) {
                $remaining = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
                foreach ($item in $remaining) {
                    if ($item.PSObject.Properties['RecordType'] -and $item.RecordType -eq 'Result') {
                        $summary.Add($item)
                        [void]$receivedResults.Add([string]$item.Computer)
                        $finalDetails = 'Tempo {0}s' -f $item.DurationSeconds
                        if ($item.ErrorReason) { $finalDetails += '; ' + $item.ErrorReason }
                        Write-MainStatus $item.Computer 'FINAL' $item.Overall $finalDetails
                    }
                    else { Write-Output ([string]$item) }
                }
                if (-not $receivedResults.Contains($job.Name)) {
                    $reason = if ($job.JobStateInfo.Reason) { $job.JobStateInfo.Reason.Message } else { 'Job remoto terminou sem resultado' }
                    $summary.Add([pscustomobject]@{
                        Computer = $job.Name; Access = 'Erro'; Copy = 'Nao executada'; Install = 'Nao executada'
                        ExitCode = ''; Validation = 'Nao executada'; Cleanup = 'Nao executada'
                        DurationSeconds = 0; ErrorReason = $reason; Overall = 'Falha'
                    })
                    Write-MainStatus $job.Name 'GERAL' 'ERRO' $reason
                }
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                [void]$active.Remove($job)
            }
        }
        if ($active.Count -gt 0) { Start-Sleep -Milliseconds 200 }
    }

    exit 0
}
catch {
    Write-Output ('ERRO GERAL: ' + $_.Exception.Message)
    exit 1
}
