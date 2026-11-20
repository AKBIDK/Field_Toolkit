#requires -Version 5.1
param([switch]$Demo, [switch]$TesteRelatorio, [switch]$Snapshot)
$global:ScriptPath = $PSCommandPath

$ErrorActionPreference = 'Stop'

$engine = {
    param($state, $order, $customer, $technician, $scan, $admin, $demo)

    $ErrorActionPreference = 'Stop'
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $samples = New-Object 'System.Collections.Generic.List[object]'

    $stability = $null
    $today = (Get-Date).Date
    $startDate = $today.AddDays(-29)

    function Add-Result($section, $item, $value, $status = 'Info') {
        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            $value = 'Não informado pelo Windows'
            $status = 'Indisponível'
        }
        $rows.Add([pscustomobject]@{ Section=$section; Item=$item; Value=[string]$value; Status=$status })
    }

    function Step($percent, $message) {
        $state.Progress = $percent
        $state.Messages.Enqueue($message)
    }

    function Try-Check($section, $item, [scriptblock]$action) {
        try { & $action }
        catch { Add-Result $section $item $_.Exception.Message 'Indisponível' }
    }

    function GiB($bytes) { '{0:N1} GiB' -f ($bytes / 1GB) }

    Step 5 'Identificando equipamento e Windows'
    if ($demo) {
        Add-Result 'Equipamento' 'Modelo' 'Computador de demonstração'
        Add-Result 'Equipamento' 'Windows' 'Windows 11 (dados fictícios)'
        Add-Result 'Equipamento' 'Número de série / Service Tag' 'DEMO-SN-0001'
        Add-Result 'Equipamento' 'BIOS' 'Fabricante fictício · versão 1.0 · 01/01/2024'
        Add-Result 'Rede' 'MAC · Ethernet' '00:00:00:00:00:01 · físico · ativo'
        Add-Result 'Rede' 'Low MAC (menor MAC entre adaptadores físicos)' '00:00:00:00:00:01 · Ethernet'
        Add-Result 'Drivers' 'Dispositivos com problema' 'Nenhum dispositivo com erro no Gerenciador de Dispositivos' 'OK'
        Add-Result 'Armazenamento' 'Disco 0' 'SSD de 512 GB (dados fictícios)'
        Add-Result 'Armazenamento' 'Temperatura' '39 °C'
        Add-Result 'Armazenamento' 'Espaço livre' '83,2 GiB'
        Add-Result 'Eventos do Windows' 'Exemplo' 'Evento crítico de energia nos últimos 7 dias' 'Atenção'
        $stability = @(
            [pscustomobject]@{ Time=$today.AddDays(-6).AddHours(10); Kind='Tela azul ou reinício por erro'; Id=1001 }
            [pscustomobject]@{ Time=$today.AddDays(-3).AddHours(14); Kind='Desligamento inesperado'; Id=41 }
        )
        foreach ($i in 0..127) {
            $sample = [pscustomobject]@{ Index=$i; Offset=[long]$i*1MB; Ms=4; Status='OK'; Detail='Demonstração' }
            if ($i -eq 62) { $sample.Status='Erro'; $sample.Detail='Falha simulada' }
            $samples.Add($sample)
            $state.Cells.Enqueue($sample)
        }
    } else {
        Try-Check 'Equipamento' 'Sistema' {
            $pc = Get-CimInstance Win32_ComputerSystem -OperationTimeoutSec 12
            $os = Get-CimInstance Win32_OperatingSystem -OperationTimeoutSec 12
            Add-Result 'Equipamento' 'Modelo' ("{0} {1}" -f $pc.Manufacturer, $pc.Model).Trim()
            Add-Result 'Equipamento' 'Windows' ("{0} · build {1}" -f $os.Caption, $os.BuildNumber)
            Add-Result 'Equipamento' 'Nome do computador' $pc.Name
        }

        Step 20 'Coletando processador, memória e vídeo'
        Try-Check 'Equipamento' 'Processador' {
            foreach ($cpu in (Get-CimInstance Win32_Processor -OperationTimeoutSec 12)) {
                Add-Result 'Equipamento' 'Processador' ("{0} · {1} núcleos · {2} threads" -f $cpu.Name.Trim(), $cpu.NumberOfCores, $cpu.NumberOfLogicalProcessors)
                if ($null -ne $cpu.LoadPercentage) { Add-Result 'Equipamento' 'Carga da CPU no momento da coleta' "$($cpu.LoadPercentage)%" }
            }
            $sha = [Security.Cryptography.SHA256]::Create()
            try {
                $bytes = [Text.Encoding]::ASCII.GetBytes('abc')
                $expected = 'BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD'
                $clock = [Diagnostics.Stopwatch]::StartNew()
                $iterations = 0
                $valid = $true
                do {
                    $digest = [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '')
                    $valid = $digest -ceq $expected
                    $iterations++
                } while ($valid -and $clock.ElapsedMilliseconds -lt 1500)
                if ($valid) { Add-Result 'Verificações rápidas' 'CPU · cálculo SHA-256' "$iterations cálculos corretos em 1,5 s (uma thread)" 'OK' }
                else { Add-Result 'Verificações rápidas' 'CPU · cálculo SHA-256' 'Resultado divergente do valor esperado' 'Atenção' }
            } finally { $sha.Dispose() }
        }
        Try-Check 'Equipamento' 'Memória' {
            $os = Get-CimInstance Win32_OperatingSystem -OperationTimeoutSec 12
            Add-Result 'Equipamento' 'RAM visível ao Windows' (GiB ($os.TotalVisibleMemorySize * 1KB))
            Add-Result 'Equipamento' 'RAM disponível na coleta' (GiB ($os.FreePhysicalMemory * 1KB))
            $modules = @(Get-CimInstance Win32_PhysicalMemory -OperationTimeoutSec 12)
            if ($modules.Count -eq 0) { Add-Result 'Equipamento' 'Módulos de RAM' '' }
            $arrays = @(Get-CimInstance Win32_PhysicalMemoryArray -Filter 'Use=3' -OperationTimeoutSec 12)
            if ($arrays.Count -eq 1 -and $arrays[0].MemoryDevices -ge $modules.Count -and $arrays[0].MemoryDevices -gt 0) {
                $slots = [int]$arrays[0].MemoryDevices
                Add-Result 'Equipamento' 'Encaixes de RAM segundo o firmware' "$slots informados · $($slots - $modules.Count) livres estimados"
                $maxKb = [long]$arrays[0].MaxCapacityEx
                if ($maxKb -le 0 -and $arrays[0].MaxCapacity -lt 2147483648) { $maxKb = $arrays[0].MaxCapacity }
                if ($maxKb -gt 0) { Add-Result 'Equipamento' 'Capacidade máxima de RAM informada pelo firmware' (GiB ($maxKb * 1KB)) }
            }
            foreach ($module in $modules) {
                Add-Result 'Equipamento' "RAM · $($module.DeviceLocator)" (("{0} · {1} {2}" -f (GiB $module.Capacity), $module.Manufacturer, $module.PartNumber).Trim())
            }
        }
        Try-Check 'Equipamento' 'Vídeo' {
            foreach ($gpu in (Get-CimInstance Win32_VideoController -OperationTimeoutSec 12)) {
                Add-Result 'Equipamento' 'Vídeo' ("{0} · driver {1}" -f $gpu.Name, $gpu.DriverVersion)
            }
        }
        Try-Check 'Equipamento' 'Bateria' {
            $batteries = @(Get-CimInstance Win32_Battery -OperationTimeoutSec 12)
            if ($batteries.Count -eq 0) { Add-Result 'Equipamento' 'Bateria' 'Nenhuma bateria identificada' }
            foreach ($battery in $batteries) {
                if ($null -eq $battery.EstimatedChargeRemaining) { Add-Result 'Equipamento' 'Carga da bateria' '' }
                else { Add-Result 'Equipamento' 'Carga da bateria' "$($battery.EstimatedChargeRemaining)%" }
            }
            if ($batteries.Count -gt 0) {
                Try-Check 'Equipamento' 'Saúde da bateria' {
                    $found = 0
                    function Add-Health($design, $full, $cycles) {
                        if (-not $design -or -not $full -or $design -in 0, 4294967295 -or $full -in 0, 4294967295) { return $false }
                        $percent = [Math]::Round(100 * $full / $design, 0)
                        $status = if ($percent -lt 70) { 'Atenção' } else { 'Info' }
                        Add-Result 'Equipamento' 'Saúde da bateria (capacidade atual / original)' "$percent% · $full de $design mWh" $status
                        if ($null -ne $cycles -and $cycles -gt 0) { Add-Result 'Equipamento' 'Ciclos de carga da bateria' "$cycles" }
                        return $true
                    }
                    # Fonte 1: classes WMI da bateria (podem exigir administrador)
                    try {
                        $design = @(Get-CimInstance -Namespace root/wmi -ClassName BatteryStaticData -OperationTimeoutSec 10 | Where-Object Active)
                        $current = @(Get-CimInstance -Namespace root/wmi -ClassName BatteryFullChargedCapacity -OperationTimeoutSec 10 | Where-Object Active)
                        $cc = (@(Get-CimInstance -Namespace root/wmi -ClassName BatteryCycleCount -OperationTimeoutSec 10 -ErrorAction SilentlyContinue) | Select-Object -First 1).CycleCount
                        foreach ($d in $design) {
                            $m = @($current | Where-Object InstanceName -eq $d.InstanceName)
                            if ($m.Count -eq 0 -and $design.Count -eq 1 -and $current.Count -eq 1) { $m = $current }
                            if ($m.Count -ge 1 -and (Add-Health $d.DesignedCapacity $m[0].FullChargedCapacity $cc)) { $found++ }
                        }
                    } catch { }
                    # Fonte 2: relatório de bateria do próprio Windows (funciona sem administrador)
                    if ($found -eq 0) {
                        $xmlPath = Join-Path ([IO.Path]::GetTempPath()) ('bateria-' + [Guid]::NewGuid().ToString('N') + '.xml')
                        try {
                            & powercfg.exe /batteryreport /xml /output $xmlPath 2>&1 | Out-Null
                            if (Test-Path -LiteralPath $xmlPath) {
                                $xml = New-Object Xml.XmlDocument; $xml.Load($xmlPath)
                                foreach ($bat in $xml.SelectNodes("//*[local-name()='Battery']")) {
                                    $dc = $bat.SelectSingleNode("*[local-name()='DesignCapacity']").InnerText
                                    $fc = $bat.SelectSingleNode("*[local-name()='FullChargeCapacity']").InnerText
                                    $cy = $bat.SelectSingleNode("*[local-name()='CycleCount']")
                                    $cyc = if ($cy) { [int]$cy.InnerText } else { $null }
                                    if (Add-Health ([long]$dc) ([long]$fc) $cyc) { $found++ }
                                }
                            }
                        } catch { } finally { Remove-Item -LiteralPath $xmlPath -Force -ErrorAction SilentlyContinue }
                    }
                    # Fonte 3: Win32_Battery (poucos fabricantes preenchem)
                    if ($found -eq 0) {
                        foreach ($w in @(Get-CimInstance Win32_Battery -OperationTimeoutSec 10)) { if (Add-Health $w.DesignCapacity $w.FullChargeCapacity $null) { $found++ } }
                    }
                    if ($found -eq 0) { Add-Result 'Equipamento' 'Saúde da bateria' 'Firmware ou driver não informou a capacidade original e a atual. Use Ferramentas > Testes > Relatório de bateria.' 'Indisponível' }
                }
            }
        }

        Step 40 'Verificando espaço, saúde e telemetria dos discos'
        Try-Check 'Armazenamento' 'Volumes locais' {
            foreach ($volume in (Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -OperationTimeoutSec 12)) {
                if ($volume.Size -gt 0) {
                    $free = 100 * $volume.FreeSpace / $volume.Size
                    $status = if ($free -lt 10) { 'Atenção' } else { 'OK' }
                    Add-Result 'Armazenamento' "Volume $($volume.DeviceID)" ('{0} livres de {1} ({2:N1}% livres)' -f (GiB $volume.FreeSpace), (GiB $volume.Size), $free) $status
                }
            }
        }
        $systemDisk = $null
        try {
            $letter = $env:SystemDrive.TrimEnd(':')
            $systemDisk = Get-Partition -DriveLetter $letter | Get-Disk
        } catch { Add-Result 'Armazenamento' 'Disco do Windows' $_.Exception.Message 'Indisponível' }
        Try-Check 'Armazenamento' 'Discos físicos' {
            $disks = @(Get-Disk)
            if ($disks.Count -eq 0) { Add-Result 'Armazenamento' 'Discos físicos' '' }
            foreach ($disk in $disks) {
                $name = "Disco $($disk.Number)"
                $status = switch ([string]$disk.HealthStatus) {
                    'Healthy' { 'OK' }
                    { $_ -in 'Warning','Unhealthy' } { 'Atenção' }
                    default { 'Indisponível' }
                }
                Add-Result 'Armazenamento' $name ("{0} · {1} · {2} · saúde reportada: {3}" -f $disk.FriendlyName, (GiB $disk.Size), $disk.BusType, $disk.HealthStatus) $status
                Try-Check 'Armazenamento' "$name · contadores de confiabilidade" {
                    $counter = Get-StorageReliabilityCounter -Disk $disk
                    if ($null -eq $counter) { Add-Result 'Armazenamento' "$name · contadores de confiabilidade" ''; return }
                    $metrics = [ordered]@{
                        Temperature = @('Temperatura',' °C')
                        Wear = @('Desgaste reportado','%')
                        PowerOnHours = @('Horas ligado',' h')
                        ReadErrorsUncorrected = @('Erros de leitura não corrigidos','')
                        WriteErrorsUncorrected = @('Erros de gravação não corrigidos','')
                    }
                    foreach ($metric in $metrics.GetEnumerator()) {
                        $value = $counter.($metric.Key)
                        $label = "$name · $($metric.Value[0])"
                        if ($null -eq $value -or ($metric.Key -eq 'Temperature' -and $value -eq 0)) {
                            Add-Result 'Armazenamento' $label ''
                            continue
                        }
                        $status = if (($metric.Key -like '*Uncorrected' -and $value -gt 0) -or ($metric.Key -eq 'Wear' -and $value -ge 90)) { 'Atenção' } else { 'Info' }
                        Add-Result 'Armazenamento' $label "$value$($metric.Value[1])" $status
                    }
                }
            }
        }

        Step 55 'Consultando eventos do Windows e adaptadores de rede'
        Try-Check 'Eventos do Windows' 'Registro Sistema' {
            $events = @()
            try {
                $events = @(Get-WinEvent -FilterHashtable @{LogName='System'; Level=1,2; StartTime=(Get-Date).AddDays(-7)} -MaxEvents 20)
            } catch {
                if ($_.FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*') { throw }
            }
            if ($events.Count -eq 0) { Add-Result 'Eventos do Windows' 'Últimos 7 dias' 'Sem eventos críticos ou de erro no registro Sistema' 'OK' }
            foreach ($event in $events) {
                Add-Result 'Eventos do Windows' ("{0:dd/MM HH:mm} · {1} · ID {2}" -f $event.TimeCreated, $event.ProviderName, $event.Id) ([string]$event.Message -replace '\s+', ' ') 'Revisar'
            }
        }
        try {
            $powerEvents = @()
            try { $powerEvents = @(Get-WinEvent -FilterHashtable @{LogName='System'; Id=41,1001,6008; StartTime=$startDate}) }
            catch { if ($_.FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*') { throw } }
            $stability = @(
                foreach ($event in $powerEvents) {
                    $kind = switch ($event.Id) {
                        41 { if ($event.ProviderName -eq 'Microsoft-Windows-Kernel-Power') { 'Desligamento inesperado' } }
                        6008 { if ($event.ProviderName -eq 'EventLog') { 'Desligamento inesperado' } }
                        1001 { if ($event.ProviderName -in 'Microsoft-Windows-WER-SystemErrorReporting','BugCheck') { 'Tela azul ou reinício por erro' } }
                    }
                    if ($kind) { [pscustomobject]@{ Time=$event.TimeCreated; Kind=$kind; Id=$event.Id } }
                }
            )
            if ($stability.Count) {
                Add-Result 'Eventos do Windows' 'Estabilidade em 30 dias' "$($stability.Count) registros relevantes. Um incidente pode gerar mais de um evento." 'Revisar'
            }
        } catch { Add-Result 'Eventos do Windows' 'Estabilidade em 30 dias' $_.Exception.Message 'Indisponível' }
        Try-Check 'Rede' 'Adaptadores ativos'  {
            $adapters = @(Get-NetAdapter | Where-Object Status -eq 'Up')
            if ($adapters.Count -eq 0) { Add-Result 'Rede' 'Adaptadores ativos' 'Nenhum adaptador reportado como ativo' }
            foreach ($adapter in $adapters) {
                Add-Result 'Rede' $adapter.Name ("{0} · enlace {1}" -f $adapter.InterfaceDescription, $adapter.LinkSpeed)
            }
        }

        Step 60 'Lendo BIOS, MAC e drivers'
        Try-Check 'Equipamento' 'BIOS' {
            $bios = Get-CimInstance Win32_BIOS -OperationTimeoutSec 12
            $bb = Get-CimInstance Win32_BaseBoard -OperationTimeoutSec 12
            $prod = Get-CimInstance Win32_ComputerSystemProduct -OperationTimeoutSec 12
            Add-Result 'Equipamento' 'Número de série / Service Tag' $bios.SerialNumber
            Add-Result 'Equipamento' 'UUID' $prod.UUID
            Add-Result 'Equipamento' 'BIOS' ("{0} · versão {1} · {2:dd/MM/yyyy}" -f $bios.Manufacturer, $bios.SMBIOSBIOSVersion, $bios.ReleaseDate)
            Add-Result 'Equipamento' 'Placa-mãe' ("{0} {1}" -f $bb.Manufacturer, $bb.Product).Trim()
            $sb = try { if (Confirm-SecureBootUEFI) { 'Ativado' } else { 'Desativado' } } catch { 'BIOS legado ou não suportado' }
            Add-Result 'Equipamento' 'Secure Boot' $sb
            $tpm = try { $t = Get-Tpm; if ($t.TpmPresent) { "Presente · pronto: $($t.TpmReady)" } else { 'Não presente' } } catch { '' }
            Add-Result 'Equipamento' 'TPM' $tpm
        }
        Try-Check 'Rede' 'Endereços MAC e IP' {
            $macs = @(Get-CimInstance Win32_NetworkAdapter -OperationTimeoutSec 12 | Where-Object { $_.MACAddress } | Sort-Object MACAddress)
            foreach ($a in $macs) {
                $tipo = if ($a.PhysicalAdapter) { 'físico' } else { 'virtual/lógico' }
                $ativo = if ($a.NetEnabled) { 'ativo' } else { 'inativo ou desativado' }
                Add-Result 'Rede' "MAC · $($a.Name)" ("{0} · {1} · {2}" -f $a.MACAddress, $tipo, $ativo)
            }
            $low = @($macs | Where-Object PhysicalAdapter | Select-Object -First 1)
            if ($low.Count) { Add-Result 'Rede' 'Low MAC (menor MAC entre adaptadores físicos)' ("{0} · {1}" -f $low[0].MACAddress, $low[0].Name) }
            foreach ($ip in @(Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike '169.254*' -and $_.IPAddress -ne '127.0.0.1' })) { Add-Result 'Rede' "IPv4 · $($ip.InterfaceAlias)" $ip.IPAddress }
        }
        Try-Check 'Drivers' 'Dispositivos com problema' {
            $bad = @(Get-PnpDevice -PresentOnly | Where-Object { $_.Status -in 'Error', 'Degraded' })
            if ($bad.Count -eq 0) { Add-Result 'Drivers' 'Dispositivos com problema' 'Nenhum dispositivo com erro no Gerenciador de Dispositivos' 'OK' }
            foreach ($d in $bad) { $n = if ($d.FriendlyName) { $d.FriendlyName } else { $d.InstanceId }; Add-Result 'Drivers' $n ("{0} · {1}" -f $d.Class, $d.Status) 'Atenção' }
        }
        if (!$scan) { Add-Result 'Mapa de leitura' 'Amostragem' 'Desmarcada na abertura da ordem de serviço' 'Não executado' }
        elseif (!$admin -or $null -eq $systemDisk) {
            Add-Result 'Mapa de leitura' 'Amostragem' 'Requer PowerShell elevado e identificação do disco do Windows' 'Indisponível'
        } else {
            Step 65 "Amostrando somente leitura no disco $($systemDisk.Number)"
            Try-Check 'Mapa de leitura' 'Amostragem' {
                if (-not ('Bancada.RawDisk' -as [type])) {
                    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace Bancada {
    public static class RawDisk {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string path, uint access, uint share,
            IntPtr security, uint disposition, uint flags, IntPtr template);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetFilePointerEx(SafeFileHandle handle, long offset,
            out long position, uint origin);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool ReadFile(SafeFileHandle handle, [Out] byte[] bytes,
            uint count, out uint read, IntPtr overlapped);
        public static SafeFileHandle Open(int disk) {
            SafeFileHandle handle = CreateFile(@"\\.\PhysicalDrive" + disk, 0x80000000u, 3u,
                IntPtr.Zero, 3u, 0u, IntPtr.Zero);
            if (handle.IsInvalid) {
                int error = Marshal.GetLastWin32Error();
                handle.Dispose();
                throw new Win32Exception(error);
            }
            return handle;
        }
        public static void Read(SafeFileHandle handle, long offset, byte[] buffer) {
            long position;
            if (!SetFilePointerEx(handle, offset, out position, 0u))
                throw new Win32Exception(Marshal.GetLastWin32Error());
            uint length;
            if (!ReadFile(handle, buffer, (uint)buffer.Length, out length, IntPtr.Zero))
                throw new Win32Exception(Marshal.GetLastWin32Error());
            if (length != buffer.Length)
                throw new System.IO.EndOfStreamException("Leitura incompleta.");
        }
    }
}
'@
                }
                $block = [Math]::Max(4096, [int]$systemDisk.LogicalSectorSize)
                $last = [long][Math]::Floor($systemDisk.Size / $block) - 1
                if ($last -lt 127) { throw 'Capacidade insuficiente para 128 amostras distintas.' }
                $buffer = New-Object byte[] $block
                $handle = [Bancada.RawDisk]::Open($systemDisk.Number)
                $failed = 0
                $slow = 0
                try {
                    foreach ($i in 0..127) {
                        $offset = [long][Math]::Floor(($last * [double]$i) / 127) * $block
                        $clock = [Diagnostics.Stopwatch]::StartNew()
                        $status = 'OK'
                        $detail = 'Bloco lido'
                        try {
                            [Bancada.RawDisk]::Read($handle, $offset, $buffer)
                            if ($clock.ElapsedMilliseconds -ge 100) { $status='Lento'; $detail='Leitura de 100 ms ou mais'; $slow++ }
                        } catch {
                            $status = 'Erro'
                            $detail = $_.Exception.GetBaseException().Message
                            $failed++
                        }
                        $sample = [pscustomobject]@{ Index=$i; Offset=$offset; Ms=[Math]::Round($clock.Elapsed.TotalMilliseconds,1); Status=$status; Detail=$detail }
                        $samples.Add($sample)
                        $state.Cells.Enqueue($sample)
                        $state.Progress = 65 + [int](30 * ($i + 1) / 128)
                        # Evita prosseguir quando o dispositivo deixa de responder repetidamente.
                        if ($failed -ge 8) { break }
                    }
                } finally { $handle.Dispose() }
                $status = if ($failed -gt 0 -or $slow -gt 0) { 'Atenção' } else { 'OK' }
                Add-Result 'Mapa de leitura' "Disco $($systemDisk.Number)" "$($samples.Count) de 128 blocos verificados · $failed falhas · $slow lentos · $block bytes por bloco" $status
            }
        }
    }

    Step 100 'Coleta concluída'
    [pscustomobject]@{
        Order=$order; Customer=$customer; Technician=$technician; Demo=$demo
        Date=(Get-Date); Rows=$rows.ToArray(); Samples=$samples.ToArray(); Stability=$stability; StartDate=$startDate
    }
}

function Save-Report($result, [string]$folder) {
    function ConvertTo-BancadaHtmlText($value) { [Net.WebUtility]::HtmlEncode([string]$value) }

    [void][IO.Directory]::CreateDirectory($folder)
    $hostname = [Environment]::MachineName
    $safeHost = $hostname -replace '[^a-zA-Z0-9_-]', '_'
    $createdAt = Get-Date
    $attention = @($result.Rows | Where-Object Status -in 'Atenção','Revisar').Count
    $missing = @($result.Rows | Where-Object Status -eq 'Indisponível').Count
    $headline = if ($attention) { 'Há registros para revisão' } else { 'Nenhum alerta nas verificações executadas' }
    $demoBanner = if ($result.Demo) { '<p class="demo">DEMONSTRAÇÃO · Dados fictícios</p>' } else { '' }
    $sections = foreach ($group in ($result.Rows | Group-Object Section)) {
        $items = foreach ($row in $group.Group) {
            $class = switch ($row.Status) {
                'OK' { 'ok' }
                { $_ -in 'Atenção','Revisar' } { 'warning' }
                default { 'neutral' }
            }
            "<tr><th scope='row'>$(ConvertTo-BancadaHtmlText $row.Item)</th><td>$(ConvertTo-BancadaHtmlText $row.Value)</td><td><span class='$class'>$(ConvertTo-BancadaHtmlText $row.Status)</span></td></tr>"
        }
        "<section><h2>$(ConvertTo-BancadaHtmlText $group.Name)</h2><table><thead><tr><th>Item</th><th>Resultado</th><th>Situação</th></tr></thead><tbody>$($items -join '')</tbody></table></section>"
    }
    $map = ''
    if ($result.Samples.Count -gt 0) {
        $cells = foreach ($i in 0..127) {
            if ($i -ge $result.Samples.Count) { "<span class='neutral' title='Não testado'></span>"; continue }
            $cell = $result.Samples[$i]
            $style = switch ($cell.Status) { 'OK' { 'ok' }; 'Lento' { 'warning' }; default { 'error' } }
            $title = ConvertTo-BancadaHtmlText ("Amostra {0} · offset {1} bytes · {2} ms · {3}" -f ($i+1), $cell.Offset, $cell.Ms, $cell.Detail)
            "<span class='$style' title='$title'></span>"
        }
        $failures = foreach ($cell in ($result.Samples | Where-Object Status -ne 'OK')) {
            "<li>Amostra $($cell.Index+1) · offset $(ConvertTo-BancadaHtmlText $cell.Offset) bytes · $(ConvertTo-BancadaHtmlText $cell.Ms) ms · $(ConvertTo-BancadaHtmlText $cell.Status) · $(ConvertTo-BancadaHtmlText $cell.Detail)</li>"
        }
        $map = "<section><h2>Mapa de amostras de leitura</h2><div class='grid'>$($cells -join '')</div><p>Verde = lido · Amarelo = lento · Vermelho = erro · Cinza = não testado</p><ul>$($failures -join '')</ul></section>"
    }

    $calendar = '<section><h2>Estabilidade em 30 dias</h2><p>Registro Sistema indisponível.</p></section>'
    if ($null -ne $result.Stability) {
        $dayCells = foreach ($i in 0..29) {
            $day = $result.StartDate.AddDays($i)
            $events = @($result.Stability | Where-Object { $_.Time.Date -eq $day })
            $kinds = @($events | Select-Object -ExpandProperty Kind -Unique)
            $class = if ($kinds -contains 'Tela azul ou reinício por erro') { 'error' }
                elseif ($events.Count) { 'warning' } else { 'neutral' }
            $label = if ($kinds.Count) { $kinds -join ' / ' } else { 'Sem evento registrado' }
            "<li class='$class'><strong>$($day.ToString('dd/MM'))</strong><br>$(ConvertTo-BancadaHtmlText $label)</li>"
        }
        $eventDetails = foreach ($event in $result.Stability) {
            "<li>$(ConvertTo-BancadaHtmlText ($event.Time.ToString('dd/MM HH:mm'))) · $(ConvertTo-BancadaHtmlText $event.Kind) · evento $($event.Id)</li>"
        }
        $calendar = "<section><h2>Estabilidade em 30 dias</h2><ul class='calendar'>$($dayCells -join '')</ul><ul>$($eventDetails -join '')</ul><p>Datas de registro no Windows. Um mesmo incidente pode gerar mais de um evento. O calendário depende da retenção do log.</p></section>"
    }
    $shopName = 'Bancada de diagnóstico'
    $shopContact = ''
    $config = Join-Path $PSScriptRoot 'Loja.json'
    if (Test-Path -LiteralPath $config) {
        $shop = Get-Content -LiteralPath $config -Raw -Encoding UTF8 | ConvertFrom-Json
        if (![string]::IsNullOrWhiteSpace($shop.Nome)) { $shopName = $shop.Nome }
        $shopContact = $shop.Contato
    }
    $quote = @'
<section id="budget">
<h2>Orçamento de serviço</h2>
<p>Informe os valores e marque os serviços aprovados. Imprima o laudo ou salve em PDF para registrar a seleção.</p>
<table><thead><tr><th>Serviço ou peça</th><th>Valor em R$</th><th>Aprovado</th></tr></thead><tbody id="quote-lines"></tbody></table>
<p class="quote-actions"><button type="button" id="add-line">Adicionar item</button> <button type="button" id="print-report">Imprimir / PDF</button></p>
<p><strong>Total aprovado</strong> <output id="quote-total" aria-live="polite">R$ 0,00</output></p>
</section>
<script>
(function () {
    var lines = document.getElementById('quote-lines');
    var total = document.getElementById('quote-total');
    var brl = new Intl.NumberFormat('pt-BR', {style:'currency', currency:'BRL'});
    function cents(text) {
        var match = /^(\d{1,7})(?:[,.](\d{1,2}))?$/.exec(text.trim());
        if (!match) return null;
        var decimal = match[2] || '';
        return Number(match[1]) * 100 + Number((decimal + '00').slice(0, 2));
    }
    function update() {
        var sum = 0;
        Array.prototype.forEach.call(lines.rows, function (row) {
            var amount = row.querySelector('.price');
            var value = cents(amount.value);
            var valid = !amount.value.trim() || value !== null;
            amount.setAttribute('aria-invalid', valid ? 'false' : 'true');
            amount.setCustomValidity(valid ? '' : 'Digite o valor como 150 ou 150,50.');
            if (row.querySelector('.approved').checked && value !== null) sum += value;
        });
        total.textContent = brl.format(sum / 100);
    }
    function add() {
        var line = lines.insertRow();
        line.innerHTML = '<td><input aria-label="Serviço ou peça" maxlength="120"></td><td><input class="price" inputmode="decimal" aria-label="Valor em reais" maxlength="10"></td><td><input type="checkbox" class="approved" aria-label="Item aprovado"></td>';
        return line;
    }
    add();
    document.getElementById('add-line').addEventListener('click', function () {
        add().querySelector('input').focus();
    });
    lines.addEventListener('input', update);
    lines.addEventListener('change', update);
    document.getElementById('print-report').addEventListener('click', function () {
        update();
        var invalid = lines.querySelector('[aria-invalid="true"]');
        if (invalid) { invalid.reportValidity(); return; }
        window.print();
    });
}());
</script>
'@
    $html = @"
<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Laudo da OS $(ConvertTo-BancadaHtmlText $result.Order) · Bancada</title>
<style>
*{box-sizing:border-box}body{margin:0;background:#ececec;color:#2b2b2b;font:16px/1.5 'Segoe UI',sans-serif}main{max-width:1080px;margin:auto;padding:32px 18px}header{padding:28px;background:#3a3a3a;color:white;border-radius:14px}h1{margin:4px 0 12px;font-size:2rem}h2{margin:0 0 12px;font-size:1.2rem}header p{margin:6px 0}section{background:white;margin-top:16px;padding:20px;border-radius:12px;overflow:auto}table{width:100%;border-collapse:collapse;text-align:left}th,td{border-bottom:1px solid #d9d9d9;padding:9px;vertical-align:top;overflow-wrap:anywhere}th[scope=row]{width:30%;font-weight:600}thead{background:#f0f0f0}td span{display:inline-block;padding:2px 8px;border-radius:6px;font-size:.8rem;white-space:nowrap}.ok{background:#c4f1dc;color:#164b34}.warning{background:#ffe099;color:#664400}.error{background:#f68187;color:#65171a}.neutral{background:#e4e4e4;color:#3a3a3a}.grid{display:grid;grid-template-columns:repeat(16,minmax(5px,1fr));gap:4px;max-width:640px}.grid span{height:14px;border-radius:3px}.grid .ok{background:#20a575}.grid .warning{background:#f4bc46}.grid .error{background:#dc4252}.demo{color:#ffd489;font-weight:bold}footer{margin-top:20px;font-size:.9rem}@media print{body{background:white}main{padding:0}header{background:white;color:#2b2b2b;padding:0}.demo{color:#664400}section{padding:12px 0}tr{break-inside:avoid}.grid{print-color-adjust:exact;-webkit-print-color-adjust:exact}}
.calendar{display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:5px;list-style:none;padding:0}.calendar li{font-size:.8rem;padding:5px;border-radius:5px}.calendar li.error,.calendar li.warning{print-color-adjust:exact}#budget input:not([type=checkbox]){width:100%;font:inherit;padding:5px;border:1px solid #b5b5b5}.price[aria-invalid=true]{outline:2px solid #b31b34}.quote-actions button{background:#e4e4e4;border:0;border-radius:6px;padding:8px;cursor:pointer}@media print{.quote-actions{display:none}}
$(Get-WatermarkCss)</style></head><body><main>
<header><p><strong>$(ConvertTo-BancadaHtmlText $shopName)</strong> · $(ConvertTo-BancadaHtmlText $shopContact)</p><h1>$(ConvertTo-BancadaHtmlText $headline)</h1>$demoBanner
<p><strong>OS $(ConvertTo-BancadaHtmlText $result.Order)</strong> · Computador $(ConvertTo-BancadaHtmlText $hostname)</p>
<p>Laudo criado em $(ConvertTo-BancadaHtmlText ($createdAt.ToString('dd/MM/yyyy HH:mm:ss zzz')))</p>
<p>Cliente $(ConvertTo-BancadaHtmlText $result.Customer) · Técnico $(ConvertTo-BancadaHtmlText $result.Technician)</p>
<p>$attention itens para revisão · $missing métricas ou verificações indisponíveis</p></header>
$map
$calendar
$($sections -join [Environment]::NewLine)
$quote
<footer>A amostragem verifica apenas os endereços lidos. A coleta de RAM é inventário e a verificação da CPU é um teste rápido de cálculo. Eventos do Windows requerem interpretação do técnico. A análise não substitui testes completos de hardware.</footer>
</main></body></html>
"@
    $prefix = if ($result.Demo) { 'DEMO-' } else { '' }
    $stem = "{0}{1}_{2:yyyy-MM-dd_HH-mm-ss-fff}" -f $prefix,$safeHost,$createdAt
    $path = Join-Path $folder "$stem.html"
    $id = [Guid]::NewGuid().ToString('N')
    $temporaryPath = Join-Path $folder ".$id.tmp"
    try {
        [IO.File]::WriteAllText($temporaryPath, $html, [Text.UTF8Encoding]::new($false))
        try { [IO.File]::Move($temporaryPath, $path) }
        catch [IO.IOException] {
            if (![IO.File]::Exists($path)) { throw }
            $path = Join-Path $folder "${stem}_$id.html"
            [IO.File]::Move($temporaryPath, $path)
        }
    } finally {
        if ([IO.File]::Exists($temporaryPath)) { [IO.File]::Delete($temporaryPath) }
    }
    $path
}

function New-ScanState {
    [hashtable]::Synchronized(@{
        Progress=0
        Messages=[System.Collections.Concurrent.ConcurrentQueue[object]]::new()
        Cells=[System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    })
}

function global:Log($t) { $global:log.AppendText("$t`r`n"); [Windows.Forms.Application]::DoEvents() }

function global:Run($title, [scriptblock]$sb, $a = @()) {
    $ErrorActionPreference = 'Continue'
    Log "`r`n=== $title  ($(Get-Date -Format 'HH:mm:ss')) ==="
    $global:form.Cursor = 'WaitCursor'
    try { $o = (& $sb @a 2>&1 | Out-String -Width 150).TrimEnd(); if ($o) { Log $o } } catch { Log "ERRO: $_" }
    $global:form.Cursor = 'Default'
    Log '--- fim ---'
}

function global:Clean($p) {
    if (Test-Path $p) {
        $b = (Get-ChildItem $p -Recurse -Force -EA 0 | Measure-Object Length -Sum).Sum
        Remove-Item "$p\*" -Recurse -Force -EA 0
        $a = (Get-ChildItem $p -Recurse -Force -EA 0 | Measure-Object Length -Sum).Sum
        '{0}: liberado {1:N1} MB' -f $p, (($b - $a) / 1MB)
    }
}

$global:HD = "$env:ProgramData\FieldToolkit\history"

function global:Get-ReportDir {
    $it = 'C:\IT'; $d = "$it\Diagnostico"
    if (-not (Test-Path $it)) { New-Item $it -ItemType Directory -Force | Out-Null }
    $i = Get-Item $it -Force
    if (-not ($i.Attributes -band [IO.FileAttributes]::Hidden)) { $i.Attributes = $i.Attributes -bor [IO.FileAttributes]::Hidden }
    if (-not (Test-Path $d)) { New-Item $d -ItemType Directory -Force | Out-Null }
    $d
}

function global:Save-Snapshot {
    $ErrorActionPreference = 'Continue'
    New-Item $global:HD -ItemType Directory -Force | Out-Null
    $o = Get-CimInstance Win32_OperatingSystem; $v = Get-Volume -DriveLetter C
    $s = [pscustomobject]@{
        data         = (Get-Date -Format 'yyyy-MM-dd')
        uptimeH      = [math]::Round(((Get-Date) - $o.LastBootUpTime).TotalHours, 1)
        discoLivreGB = [math]::Round($v.SizeRemaining / 1GB, 1)
        discoTotalGB = [math]::Round($v.Size / 1GB, 1)
        ramUsoPct    = [math]::Round(100 * (1 - $o.FreePhysicalMemory / $o.TotalVisibleMemorySize))
    }
    $s | ConvertTo-Json | Out-File "$global:HD\$($s.data).json" -Encoding UTF8
    Get-ChildItem $global:HD -Filter *.json | Where-Object LastWriteTime -lt (Get-Date).AddDays(-30) | Remove-Item -Force
    "Snapshot salvo: $($s.data)"
}
if ($Snapshot) { Save-Snapshot | Out-Null; exit }

function global:Ev($log, $ids, $prov) {
    $f = @{ LogName = $log; Id = $ids; StartTime = (Get-Date).AddDays(-30) }
    if ($prov) { $f.ProviderName = $prov }
    @(Get-WinEvent -FilterHashtable $f -EA 0)
}

function global:Bars($ev, $color) {
    $c = @(0) * 30; $now = (Get-Date).Date
    foreach ($e in $ev) { $i = 29 - [int]($now - $e.TimeCreated.Date).TotalDays; if ($i -ge 0 -and $i -lt 30) { $c[$i]++ } }
    $m = [math]::Max(1, ($c | Measure-Object -Maximum).Maximum); $s = ''
    for ($i = 0; $i -lt 30; $i++) {
        $h = [int](90 * $c[$i] / $m); $d = $now.AddDays($i - 29).ToString('dd/MM')
        $s += "<rect x='$($i * 30 + 2)' y='$(100 - $h)' width='24' height='$h' rx='3' fill='$color'><title>${d}: $($c[$i])</title></rect>"
    }
    "<svg viewBox='0 0 900 120' class='chart'>$s<text x='2' y='116' class='ax'>$($now.AddDays(-29).ToString('dd/MM'))</text><text x='898' y='116' text-anchor='end' class='ax'>hoje</text></svg>"
}

function global:Rows($ev, $k, $cl) {
    foreach ($e in $ev) {
        $m = "$(($e.Message -split "`n")[0])"; if ($m.Length -gt 120) { $m = $m.Substring(0, 120) }
        [pscustomobject]@{ T = $e.TimeCreated; K = $k; C = $cl; M = [Net.WebUtility]::HtmlEncode($m) }
    }
}

function global:New-Dashboard {
    $ErrorActionPreference = 'Continue'
    Log 'Coletando eventos dos últimos 30 dias (pode levar um minuto)...'
    Save-Snapshot | Out-Null
    $b = Get-CimInstance Win32_BIOS; $c = Get-CimInstance Win32_ComputerSystem; $o = Get-CimInstance Win32_OperatingSystem
    $shut = Ev 'System' @(41, 6008); $bsod = Ev 'System' @(1001) 'Microsoft-Windows-WER-SystemErrorReporting'
    $app = Ev 'Application' @(1000, 1002); $boots = Ev 'System' @(6005)
    $upd = Ev 'System' @(19) 'Microsoft-Windows-WindowsUpdateClient'; $inst = Ev 'Application' @(11707) 'MsiInstaller'
    $errs = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 1, 2; StartTime = (Get-Date).AddDays(-30) } -EA 0)
    $hot = @(Get-HotFix -EA 0 | Where-Object { $_.InstalledOn -ge (Get-Date).AddDays(-30) })
    $snaps = @(Get-ChildItem $global:HD -Filter *.json -EA 0 | Sort-Object Name | ForEach-Object { Get-Content $_.FullName -Raw | ConvertFrom-Json })
    $disk = '<p class="mut">Ainda não há snapshots suficientes. Use "Instalar coleta diária" para acompanhar a tendência.</p>'
    if ($snaps.Count -ge 2) {
        $max = ($snaps | Measure-Object discoTotalGB -Maximum).Maximum; $n = $snaps.Count
        $pts = for ($i = 0; $i -lt $n; $i++) { '{0},{1}' -f [int](10 + 880 * $i / ($n - 1)), [int](110 - 100 * $snaps[$i].discoLivreGB / $max) }
        $disk = "<svg viewBox='0 0 900 120' class='chart'><polyline points='$($pts -join ' ')' fill='none' stroke='#6b6b6b' stroke-width='3'/></svg><p class='mut'>Espaço livre em C: de $($snaps[0].discoLivreGB) GB para $($snaps[-1].discoLivreGB) GB ($n dias coletados)</p>"
    }
    $all = @(Rows $shut 'Desligamento inesperado' 'bad') + @(Rows $bsod 'Tela azul (BSOD)' 'bad') + @(Rows ($app | Select-Object -First 25) 'Falha de aplicativo' 'warn') +
           @(Rows $upd 'Atualização do Windows' 'ok') + @(Rows $inst 'Programa instalado' 'ok') +
           @($hot | ForEach-Object { [pscustomobject]@{ T = [datetime]$_.InstalledOn; K = 'Hotfix'; C = 'ok'; M = $_.HotFixID } })
    $tl = ($all | Sort-Object T -Descending | Select-Object -First 80 | ForEach-Object { "<tr><td>$($_.T.ToString('dd/MM HH:mm'))</td><td><span class='tag $($_.C)'>$($_.K)</span></td><td>$($_.M)</td></tr>" }) -join ''
    if (-not $tl) { $tl = '<tr><td colspan="3">Nenhum evento relevante nos últimos 30 dias.</td></tr>' }
    if ($bsod.Count -gt 0 -or $shut.Count -gt 2) { $st = 'Atenção'; $sc = 'bad' } elseif ($app.Count -gt 20 -or $errs.Count -gt 200) { $st = 'Observar'; $sc = 'warn' } else { $st = 'Saudável'; $sc = 'ok' }
    $v = Get-Volume -DriveLetter C; $up = [math]::Round(((Get-Date) - $o.LastBootUpTime).TotalHours, 1)
    $html = @"
<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Field Toolkit - $env:COMPUTERNAME</title><style>
:root{--bg:#efefef;--card:#fff;--tx:#262626;--mut:#707070;--bd:#dcdcdc;--ok:#1f9d63;--warn:#d98a00;--bad:#d64545}
@media(prefers-color-scheme:dark){:root{--bg:#1b1b1b;--card:#262626;--tx:#ececec;--mut:#9a9a9a;--bd:#3a3a3a}}
*{box-sizing:border-box}body{margin:0;padding:24px;background:var(--bg);color:var(--tx);font-family:Segoe UI,system-ui,sans-serif}
.wrap{max-width:1000px;margin:auto}h1{margin:0;font-size:26px}h2{font-size:16px;margin:0 0 10px}.mut{color:var(--mut);font-size:13px}
.head{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;margin-bottom:18px}
.card{background:var(--card);border:1px solid var(--bd);border-radius:14px;padding:16px;margin-bottom:16px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:16px}
.kpi{background:var(--card);border:1px solid var(--bd);border-radius:14px;padding:14px}.kpi b{display:block;font-size:28px}
.info{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:8px 20px;font-size:14px}.info span{color:var(--mut)}
.chart{width:100%;height:auto}.ax{fill:var(--mut);font-size:11px}
table{width:100%;border-collapse:collapse;font-size:13px}td{padding:7px 6px;border-top:1px solid var(--bd);vertical-align:top}
.tag{padding:2px 9px;border-radius:99px;font-size:12px;color:#fff;white-space:nowrap}.tag.ok,.pill.ok{background:var(--ok)}.tag.warn,.pill.warn{background:var(--warn)}.tag.bad,.pill.bad{background:var(--bad)}
.pill{padding:6px 16px;border-radius:99px;color:#fff;font-weight:600}.tbl{overflow-x:auto}
$(Get-WatermarkCss)</style></head><body><div class="wrap">
<div class="head"><div><h1>$env:COMPUTERNAME</h1><div class="mut">Histórico de 30 dias • gerado em $(Get-Date -Format 'dd/MM/yyyy HH:mm')</div></div><span class="pill $sc">$st</span></div>
<div class="card"><h2>Identificação</h2><div class="info">
<div><span>Modelo</span><br>$($c.Manufacturer) $($c.Model)</div><div><span>Serial / Service Tag</span><br>$($b.SerialNumber)</div>
<div><span>BIOS</span><br>$($b.SMBIOSBIOSVersion) ($($b.ReleaseDate.ToString('dd/MM/yyyy')))</div><div><span>Windows</span><br>$($o.Caption) b$($o.BuildNumber)</div>
<div><span>Ligada há</span><br>$up h</div><div><span>Disco C: livre</span><br>$([math]::Round($v.SizeRemaining/1GB,1)) de $([math]::Round($v.Size/1GB,1)) GB</div></div></div>
<div class="grid"><div class="kpi"><b>$($bsod.Count)</b><span class="mut">Telas azuis</span></div><div class="kpi"><b>$($shut.Count)</b><span class="mut">Desligamentos inesperados</span></div>
<div class="kpi"><b>$($app.Count)</b><span class="mut">Falhas de apps</span></div><div class="kpi"><b>$($errs.Count)</b><span class="mut">Erros do sistema</span></div>
<div class="kpi"><b>$($boots.Count)</b><span class="mut">Inicializações</span></div><div class="kpi"><b>$($upd.Count + $hot.Count)</b><span class="mut">Atualizações</span></div></div>
<div class="card"><h2>Erros do sistema por dia</h2>$(Bars $errs '#d64545')</div>
<div class="card"><h2>Falhas de aplicativos por dia</h2>$(Bars $app '#d98a00')</div>
<div class="card"><h2>Espaço livre em C: (snapshots)</h2>$disk</div>
<div class="card tbl"><h2>Linha do tempo</h2><table>$tl</table></div></div></body></html>
"@
    $f = "$(Get-ReportDir)\FieldToolkit_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd').html"
    $html | Out-File $f -Encoding UTF8; Start-Process $f; "Dashboard salvo e aberto: $f"
}

function global:Get-LogoDataUri {
    $f = Get-ChildItem 'C:\IT' -Filter 'logo.*' -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp' } | Select-Object -First 1
    if (-not $f -or $f.Length -gt 3MB) { return '' }
    $mime = switch ($f.Extension.ToLower()) { '.png' { 'image/png' } '.jpg' { 'image/jpeg' } '.jpeg' { 'image/jpeg' } '.gif' { 'image/gif' } '.svg' { 'image/svg+xml' } '.webp' { 'image/webp' } }
    'data:' + $mime + ';base64,' + [Convert]::ToBase64String([IO.File]::ReadAllBytes($f.FullName))
}

function global:Get-WatermarkCss {
    $u = Get-LogoDataUri
    if (-not $u) { return '' }
    "body::before{content:'';position:fixed;inset:0;background:url('$u') center/45% no-repeat;opacity:.07;pointer-events:none;z-index:9999;print-color-adjust:exact;-webkit-print-color-adjust:exact}"
}

function global:Set-ReportLogo {
    $atual = @(Get-ChildItem 'C:\IT' -Filter 'logo.*' -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp' })
    if ($atual.Count) {
        $r = [Windows.Forms.MessageBox]::Show("Já existe um logo configurado.`n`nSim = trocar`nNão = remover`nCancelar = manter", 'Logo do laudo', 'YesNoCancel', 'Question')
        if ($r -eq 'Cancel') { return 'Logo mantido.' }
        if ($r -eq 'No') { $atual | Remove-Item -Force; return 'Logo removido. Os próximos laudos saem sem marca d''água.' }
    }
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Title = 'Escolha a imagem do logo'
    $dlg.Filter = 'Imagens (png, jpg, svg, gif, webp)|*.png;*.jpg;*.jpeg;*.svg;*.gif;*.webp'
    if ($dlg.ShowDialog() -ne 'OK') { return 'Nenhum arquivo escolhido.' }
    $f = Get-Item -LiteralPath $dlg.FileName
    if ($f.Length -gt 3MB) { return 'Imagem acima de 3 MB. Escolha um arquivo menor.' }
    [void](Get-ReportDir)
    $atual | Remove-Item -Force -ErrorAction SilentlyContinue
    Copy-Item -LiteralPath $f.FullName -Destination ('C:\IT\logo' + $f.Extension.ToLower()) -Force
    'Logo salvo. Será a marca d''água dos próximos laudos.'
}

if ($TesteRelatorio) {
    $test = & $engine (New-ScanState) 'TESTE-001' '<Cliente>' 'Técnico' $false $false $true
    if ($test.Samples.Count -ne 128) { throw 'Quantidade inesperada de amostras.' }
    $testFolder = Join-Path ([IO.Path]::GetTempPath()) ('Bancada-Teste-' + [Guid]::NewGuid().ToString('N'))
    $testPath = Save-Report $test $testFolder
    $testHost = [regex]::Escape(([Environment]::MachineName -replace '[^a-zA-Z0-9_-]', '_'))
    if ([IO.Path]::GetFileName($testPath) -notmatch "^DEMO-${testHost}_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}-\d{3}\.html$") {
        throw 'Nome inesperado no laudo HTML.'
    }
    $testHtml = [IO.File]::ReadAllText($testPath)
    if ($testHtml -notmatch '&lt;Cliente&gt;' -or $testHtml -notmatch 'Falha simulada' -or $testHtml -notmatch 'Estabilidade em 30 dias' -or $testHtml -notmatch 'quote-total') { throw 'Falha na validação do laudo.' }
    Write-Output "Teste do HTML concluído: $testPath"
    return
}

if ($env:OS -ne 'Windows_NT') { throw 'Abra este aplicativo no PowerShell do Windows.' }
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    throw 'Inicie pelo Abrir-Bancada.cmd ou use powershell.exe -STA.'
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
$isAdmin = ([Security.Principal.WindowsPrincipal]::new(
    [Security.Principal.WindowsIdentity]::GetCurrent()
)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$history = Get-ReportDir

function Color($hex) { [Drawing.ColorTranslator]::FromHtml($hex) }

$cfg = [ordered]@{
    'Histórico 30 dias' = [ordered]@{
        'Dashboard HTML (30 dias)' = { New-Dashboard }
        'Salvar snapshot agora' = { Save-Snapshot }
        'Instalar coleta diária' = {
            New-Item "$env:ProgramData\FieldToolkit" -ItemType Directory -Force | Out-Null
            $dest = "$env:ProgramData\FieldToolkit\FieldToolkit.ps1"; Copy-Item $global:ScriptPath $dest -Force
            $a = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$dest`" -Snapshot"
            Register-ScheduledTask -TaskName 'FieldToolkit-Snapshot' -Action $a -Trigger (New-ScheduledTaskTrigger -Daily -At 12:00) -Settings (New-ScheduledTaskSettingsSet -StartWhenAvailable) -User 'SYSTEM' -RunLevel Highest -Force | Out-Null
            'Coleta diária instalada (12:00, como SYSTEM; roda ao ligar se perder o horário).'
        }
        'Remover coleta diária' = { Unregister-ScheduledTask -TaskName 'FieldToolkit-Snapshot' -Confirm:$false; 'Coleta removida.' }
        'Abrir pasta do histórico' = { New-Item $global:HD -ItemType Directory -Force | Out-Null; Start-Process explorer.exe $global:HD; 'Pasta aberta.' }
    }
    'Sistema / BIOS' = [ordered]@{
        'Resumo BIOS e Placa' = {
            $b = Get-CimInstance Win32_BIOS; $c = Get-CimInstance Win32_ComputerSystem
            $p = Get-CimInstance Win32_ComputerSystemProduct; $m = Get-CimInstance Win32_BaseBoard
            $o = Get-CimInstance Win32_OperatingSystem
            "Host: $env:COMPUTERNAME | Usuário: $env:USERNAME"
            "Fabricante/Modelo: $($c.Manufacturer) $($c.Model)"
            "Serial / Service Tag: $($b.SerialNumber)"
            "UUID: $($p.UUID)"
            "BIOS: $($b.Manufacturer) v$($b.SMBIOSBIOSVersion) - $($b.ReleaseDate.ToString('dd/MM/yyyy'))"
            "Placa-mãe: $($m.Manufacturer) $($m.Product)"
            "Windows: $($o.Caption) build $($o.BuildNumber)"
            $sb = try { Confirm-SecureBootUEFI } catch { 'Desativado / BIOS legado' }
            "Secure Boot: $sb"
            $t = try { $x = Get-Tpm; 'presente=' + $x.TpmPresent + ' pronto=' + $x.TpmReady } catch { 'não detectado' }
            "TPM: $t"
        }
        'Rede e MAC' = {
            $n = Get-CimInstance Win32_NetworkAdapter | Where-Object MACAddress | Sort-Object MACAddress
            $n | Select-Object Name, MACAddress, PhysicalAdapter, NetEnabled | Format-Table -AutoSize
            $p = $n | Where-Object PhysicalAdapter | Select-Object -First 1
            if ($p) { "Low MAC (menor MAC entre adaptadores físicos): $($p.MACAddress) - $($p.Name)" }
            Get-NetIPAddress -AddressFamily IPv4 | Select-Object InterfaceAlias, IPAddress | Format-Table -AutoSize
            'Wake-on-LAN:'
            Get-NetAdapterPowerManagement -EA 0 | Select-Object Name, WakeOnMagicPacket, WakeOnPattern | Format-Table -AutoSize
        }
        'Hardware (CPU/RAM/Disco/GPU)' = {
            Get-CimInstance Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors
            Get-CimInstance Win32_PhysicalMemory | Select-Object BankLabel, @{n = 'GB'; e = { $_.Capacity / 1GB } }, Speed, Manufacturer
            Get-PhysicalDisk | Select-Object FriendlyName, MediaType, @{n = 'GB'; e = { [math]::Round($_.Size / 1GB) } }, HealthStatus
            Get-Volume | Where-Object DriveLetter | Select-Object DriveLetter, @{n = 'LivreGB'; e = { [math]::Round($_.SizeRemaining / 1GB, 1) } }, @{n = 'TotalGB'; e = { [math]::Round($_.Size / 1GB, 1) } }
            Get-CimInstance Win32_VideoController | Select-Object Name, DriverVersion
        }
    }
    'Limpeza' = [ordered]@{
        'Temporários (usuário + Windows)' = { Clean $env:TEMP; Clean 'C:\Windows\Temp'; Clean "$env:LOCALAPPDATA\Microsoft\Windows\INetCache" }
        'Esvaziar Lixeira' = { Clear-RecycleBin -Force -EA 0; 'Lixeira esvaziada.' }
        'Cache do Windows Update' = {
            Stop-Service wuauserv -Force -EA 0; Stop-Service bits -Force -EA 0
            Clean 'C:\Windows\SoftwareDistribution\Download'
            Start-Service bits -EA 0; Start-Service wuauserv -EA 0
        }
        'Limpar cache DNS' = { Clear-DnsClientCache; 'Cache DNS limpo.' }
        'Espaço livre em C:' = { $v = Get-Volume -DriveLetter C; 'Livre em C: {0:N1} GB de {1:N1} GB' -f ($v.SizeRemaining / 1GB), ($v.Size / 1GB) }
    }
    'Otimização' = [ordered]@{
        'Plano de energia: Alto desempenho' = {
            powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
            if ($LASTEXITCODE -ne 0) { powercfg /duplicatescheme 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c | Out-Null; 'Plano criado; selecione em Opções de Energia.' }
            powercfg /getactivescheme
        }
        'SFC (verificar arquivos)' = { Start-Process cmd.exe -ArgumentList '/k sfc /scannow'; 'SFC aberto em janela separada.' }
        'DISM RestoreHealth' = { Start-Process cmd.exe -ArgumentList '/k DISM /Online /Cleanup-Image /RestoreHealth'; 'DISM aberto em janela separada.' }
        'TRIM do SSD (C:)' = { Optimize-Volume -DriveLetter C -ReTrim -Verbose 4>&1 }
        'Programas na inicialização' = { Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location | Format-List }
        'Top 10 processos por memória' = { Get-Process | Sort-Object WS -Descending | Select-Object -First 10 Name, Id, @{n = 'MB'; e = { [math]::Round($_.WS / 1MB) } } }
    }
    'Testes' = [ordered]@{
        'Rede / Internet' = {
            Test-Connection 8.8.8.8 -Count 4 | Select-Object Address, ResponseTime
            Test-NetConnection google.com -Port 443 | Select-Object ComputerName, TcpTestSucceeded
            Resolve-DnsName google.com -EA 0 | Select-Object -First 2 Name, IPAddress
        }
        'Saúde do disco (SMART)' = { Get-PhysicalDisk | Get-StorageReliabilityCounter | Select-Object DeviceId, Temperature, Wear, ReadErrorsTotal, WriteErrorsTotal, PowerOnHours }
        'Relatório de bateria' = { $f = "$(Get-ReportDir)\bateria_$env:COMPUTERNAME.html"; powercfg /batteryreport /output $f | Out-Null; "Salvo em $f" }
        'Diagnóstico de memória (reinicia)' = { Start-Process mdsched.exe; 'Abrindo o Diagnóstico de Memória do Windows.' }
        'Erros críticos (últimos 3 dias)' = { Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 1, 2; StartTime = (Get-Date).AddDays(-3) } -MaxEvents 15 -EA 0 | Select-Object TimeCreated, ProviderName, Id | Format-Table -AutoSize }
    }
    'Drivers' = [ordered]@{
        'Dispositivos com problema' = {
            $d = Get-PnpDevice -PresentOnly | Where-Object { $_.Status -in 'Error', 'Unknown', 'Degraded' }
            if ($d) { $d | Select-Object FriendlyName, Class, Status, InstanceId | Format-List } else { 'Nenhum dispositivo com erro.' }
        }
        'Lista de drivers (mais antigos primeiro)' = { Get-CimInstance Win32_PnPSignedDriver | Where-Object DeviceName | Sort-Object DriverDate | Select-Object -First 40 DeviceName, Manufacturer, DriverVersion, @{n = 'Data'; e = { $_.DriverDate.ToString('dd/MM/yyyy') } } | Format-Table -AutoSize }
        'Reescanear dispositivos' = { pnputil /scan-devices }
        'Abrir Windows Update' = { Start-Process 'ms-settings:windowsupdate'; 'Aberto.' }
    }
}

function global:Read-LogTail($p, $n) {
    try {
        $fs = [IO.File]::Open($p, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        try {
            $start = [Math]::Max(0, $fs.Length - 16384); if ($start % 2) { $start-- }
            [void]$fs.Seek($start, [IO.SeekOrigin]::Begin)
            $t = (New-Object IO.StreamReader($fs, [Text.Encoding]::Unicode)).ReadToEnd()
        } finally { $fs.Dispose() }
        (($t -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -Last $n) -join "`r`n"
    } catch { '' }
}

function global:Start-Backup($src, $dst, $sub, $prot, $mir) {
    $src = "$src".Trim().Trim('"'); $dst = "$dst".Trim().Trim('"')
    if (-not $src -or -not $dst) { throw 'Informe a origem e o destino.' }
    if (-not (Test-Path -LiteralPath $src)) { throw "Origem não encontrada: $src" }
    $isFile = Test-Path -LiteralPath $src -PathType Leaf
    $srcFull = (Resolve-Path -LiteralPath $src).ProviderPath.TrimEnd('\')
    $dstFull = [IO.Path]::GetFullPath($dst).TrimEnd('\')
    $nome = if ($isFile) { [IO.Path]::GetFileNameWithoutExtension($srcFull) } else { Split-Path $srcFull -Leaf }
    if (-not $nome) { $nome = 'Disco_' + $srcFull.Substring(0, 1) }
    if ($sub) { $dstFull = Join-Path $dstFull ('{0}_{1:yyyy-MM-dd_HH-mm}' -f $nome, (Get-Date)) }
    if (-not $isFile -and ($dstFull + '\').StartsWith($srcFull + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'O destino está dentro da origem. Escolha outra pasta de destino.' }
    New-Item -ItemType Directory -Path $dstFull -Force -ErrorAction Stop | Out-Null
    $fix = { param($x) if ($x -match '^[A-Za-z]:\\?$') { $x.Substring(0, 2) + '\.' } else { $x.TrimEnd('\') } }
    $q = { param($x) '"' + $x + '"' }
    $log = Join-Path (Get-ReportDir) ('Backup_{0}_{1:yyyyMMdd_HHmmss}.log' -f $env:COMPUTERNAME, (Get-Date))
    $srcDir = if ($isFile) { Split-Path $srcFull -Parent } else { $srcFull }
    $a = @((& $q (& $fix $srcDir)), (& $q (& $fix $dstFull)))
    if ($isFile) { $a += (& $q (Split-Path $srcFull -Leaf)) } else { $a += $(if ($mir) { '/MIR' } else { '/E' }) }
    $a += '/COPY:DAT', '/DCOPY:DAT', '/R:2', '/W:3', '/XJ', '/MT:8', '/NP', '/XD', '"$RECYCLE.BIN"', '"System Volume Information"', ('/UNILOG:' + (& $q $log))
    if ($prot) { $a += '/B' }
    $proc = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" -ArgumentList ($a -join ' ') -PassThru -WindowStyle Hidden
    $null = $proc.Handle
    @{ Proc = $proc; Log = $log; Src = $srcFull; Dest = $dstFull; Cancelled = $false }
}

function global:Find-AdUser($q, $resumo) {
    $ErrorActionPreference = 'Continue'
    $q = "$q".Trim()
    if ($q.Length -lt 2) { 'Informe ao menos 2 caracteres (e-mail, matrícula, telefone ou login).'; return }
    if (-not (Get-CimInstance Win32_ComputerSystem).PartOfDomain) { 'Este computador não está ingressado em um domínio. A consulta ao AD exige domínio e rede/VPN corporativa.'; return }
    $esc = $q -replace '\\', '\5c' -replace '\*', '\2a' -replace '\(', '\28' -replace '\)', '\29'
    try {
        $s = [adsisearcher]"(&(objectCategory=person)(objectClass=user)(|(mail=$esc)(userPrincipalName=$esc)(sAMAccountName=$esc)(employeeID=$esc)(employeeNumber=$esc)(proxyAddresses=smtp:$esc)))"
        $s.PageSize = 50; $s.SizeLimit = 20
        $res = @($s.FindAll())
        $porNome = $false; $porTel = $false
        $dg = $q -replace '\D', ''
        if ($res.Count -eq 0 -and $q -match '^[\d\s()+.\-]+$' -and $dg.Length -ge 4) {
            $key = $dg.Substring($dg.Length - [Math]::Min($dg.Length, 9))
            $k8 = $dg.Substring($dg.Length - [Math]::Min($dg.Length, 8))
            $pat = '*' + (($k8.ToCharArray()) -join '*') + '*'
            $attrs = 'telephoneNumber', 'mobile', 'homePhone', 'ipPhone', 'otherTelephone', 'otherMobile', 'otherIpPhone'
            $f = '(|' + (($attrs | ForEach-Object { "($_=$pat)" }) -join '') + ')'
            $s3 = [adsisearcher]"(&(objectCategory=person)(objectClass=user)$f)"
            $s3.PageSize = 200; $s3.SizeLimit = 0
            $res = @($s3.FindAll() | Where-Object {
                $x = $_; $ok = $false
                foreach ($a in $attrs) { foreach ($t in $x.Properties[$a.ToLower()]) { if (("$t" -replace '\D', '').Contains($key)) { $ok = $true } } }
                $ok
            })
            $porTel = $res.Count -gt 0
        }
        if ($res.Count -eq 0) {
            $s2 = [adsisearcher]"(&(objectCategory=person)(objectClass=user)(|(displayName=$esc*)(cn=$esc*)))"
            $s2.PageSize = 50; $s2.SizeLimit = 20
            $res = @($s2.FindAll()); $porNome = $true
        }
    } catch {
        "Não foi possível consultar o AD: $($_.Exception.Message)"
        'Confirme se o computador está no domínio e conectado à rede/VPN corporativa.'
        return
    }
    if ($res.Count -eq 0) { "Nenhum usuário encontrado para '$q'. Procurado em e-mail, UPN, login, matrícula (employeeID/employeeNumber), telefone e nome."; return }
    if ($res.Count -gt 1) {
        "$($res.Count) usuários encontrados" + $(if ($porNome) { ' pelo nome' } elseif ($porTel) { ' pelo telefone' } else { '' }) + '. Refine a busca com e-mail, matrícula ou login:'
        foreach ($x in $res) {
            $ph = (@('telephonenumber', 'mobile', 'ipphone') | ForEach-Object { $x.Properties[$_] } | ForEach-Object { "$_" } | Where-Object { $_ }) -join ' / '
            '  {0,-18} {1,-30} {2,-32} {3}' -f "$($x.Properties['samaccountname'])", "$($x.Properties['displayname'])", "$($x.Properties['mail'])", $ph
        }
        return
    }
    $r = $res[0]; $p = $r.Properties
    $ft = 'lastlogon', 'lastlogontimestamp', 'pwdlastset', 'accountexpires', 'badpasswordtime', 'lockouttime', 'lastlogoff'
    function Fmt($n, $v) {
        if ($v -is [byte[]]) {
            if ($n -eq 'objectsid') { return (New-Object Security.Principal.SecurityIdentifier($v, 0)).Value }
            if ($n -eq 'objectguid') { return ([Guid]::new($v)).ToString() }
            return "(binário, $($v.Length) bytes)"
        }
        if ($n -in $ft -and $v -is [int64]) {
            if ($v -le 0 -or $v -ge 9223372036854775807) { return '(nunca / não definido)' }
            return [DateTime]::FromFileTime($v).ToString('dd/MM/yyyy HH:mm:ss')
        }
        if ($v -is [datetime]) { return $v.ToString('dd/MM/yyyy HH:mm:ss') }
        "$v"
    }
    function V($n) { if ($p[$n].Count) { Fmt $n $p[$n][0] } else { '-' } }
    $uac = 0; if ($p['useraccountcontrol'].Count) { $uac = [int]$p['useraccountcontrol'][0] }
    $bloq = 'não verificado'; $expirada = ''
    try {
        $e = $r.GetDirectoryEntry(); $e.RefreshCache(@('msDS-User-Account-Control-Computed'))
        $comp = [int]$e.Properties['msDS-User-Account-Control-Computed'].Value
        $bloq = if ($comp -band 0x10) { 'SIM' } else { 'Não' }
        if ($comp -band 0x800000) { $expirada = '  (SENHA EXPIRADA)' }
    } catch { }
    $gerente = '-'; if ($p['manager'].Count) { $gerente = ([string]$p['manager'][0] -split ',')[0] -replace '^CN=', '' }
    $senha = if ($uac -band 0x10000) { 'nunca expira' } else { 'segue a política do domínio' }
    $L = { param($k, $v) '{0,-30} {1}' -f $k, $v }
    '=== Resumo ==='
    & $L 'Nome' (V 'displayname'); & $L 'Login (sAMAccountName)' (V 'samaccountname'); & $L 'UPN' (V 'userprincipalname')
    & $L 'E-mail' (V 'mail'); & $L 'Matrícula (employeeID)' (V 'employeeid'); & $L 'Número (employeeNumber)' (V 'employeenumber')
    & $L 'Cargo' (V 'title'); & $L 'Departamento' (V 'department'); & $L 'Empresa' (V 'company'); & $L 'Gerente' $gerente
    & $L 'Telefone' (V 'telephonenumber'); & $L 'Celular' (V 'mobile'); & $L 'Ramal (ipPhone)' (V 'ipphone'); & $L 'Telefone residencial' (V 'homephone'); & $L 'Escritório' (V 'physicaldeliveryofficename')
    & $L 'Conta' $(if ($uac -band 2) { 'DESATIVADA' } else { 'Ativa' }); & $L 'Bloqueada' $bloq
    & $L 'Último logon (aprox.)' (V 'lastlogontimestamp'); & $L 'Última troca de senha' ((V 'pwdlastset') + $expirada); & $L 'Senha' $senha
    & $L 'Conta expira em' (V 'accountexpires'); & $L 'Criada em' (V 'whencreated'); & $L 'Alterada em' (V 'whenchanged')
    & $L 'Grupos (memberOf)' "$($p['memberof'].Count)"; & $L 'DN' (V 'distinguishedname')
    if ($resumo) { return }
    ''; '=== Todas as propriedades (' + $p.PropertyNames.Count + ') ==='
    foreach ($n in ($p.PropertyNames | Sort-Object)) {
        $vals = @($p[$n] | ForEach-Object { Fmt $n $_ })
        if ($vals.Count -le 1) { '{0,-32} {1}' -f $n, ($vals -join '') }
        else { "$n ($($vals.Count))"; foreach ($v in $vals) { '    ' + $v } }
    }
}

function Show-Tools {
    $tf = [Windows.Forms.Form]::new()
    $tf.Text = 'Ferramentas do técnico'; $tf.Size = [Drawing.Size]::new(940, 700); $tf.StartPosition = 'CenterParent'
    $tf.BackColor = Color '#102034'; $tf.ForeColor = Color '#E8F1FF'; $tf.Font = [Drawing.Font]::new('Segoe UI', 10)
    $tl = [Windows.Forms.TextBox]::new()
    $tl.Multiline = $true; $tl.ReadOnly = $true; $tl.ScrollBars = 'Vertical'; $tl.Dock = 'Fill'
    $tl.Font = [Drawing.Font]::new('Consolas', 9); $tl.BackColor = Color '#192F47'; $tl.ForeColor = Color '#E8F1FF'
    $global:form = $tf; $global:log = $tl
    if (-not $isAdmin) { $tl.Text = "ATENÇÃO: sem administrador, limpeza, SFC/DISM, TRIM e a coleta diária podem falhar.`r`nFeche e abra com botão direito > Executar como administrador.`r`n" }
    $tabs = [Windows.Forms.TabControl]::new(); $tabs.Dock = 'Top'; $tabs.Height = 270
    foreach ($tn in $cfg.Keys) {
        $tp = [Windows.Forms.TabPage]::new($tn); $tp.BackColor = Color '#102034'
        $fp = [Windows.Forms.FlowLayoutPanel]::new(); $fp.Dock = 'Fill'; $fp.Padding = [Windows.Forms.Padding]::new(8)
        foreach ($bn in $cfg[$tn].Keys) {
            $t = $bn; $s = $cfg[$tn][$bn]
            $btn = [Windows.Forms.Button]::new(); $btn.Text = $bn; $btn.Width = 215; $btn.Height = 42
            $btn.FlatStyle = 'Flat'; $btn.FlatAppearance.BorderSize = 0; $btn.BackColor = Color '#284763'; $btn.ForeColor = Color '#E8F1FF'
            $btn.Add_Click({ Run $t $s }.GetNewClosure())
            $fp.Controls.Add($btn)
        }
        $tp.Controls.Add($fp); [void]$tabs.TabPages.Add($tp)
    }
    $bkCfg = "$env:ProgramData\FieldToolkit\backup.json"
    $last = try { Get-Content $bkCfg -Raw -ErrorAction Stop | ConvertFrom-Json } catch { $null }
    $tk = [Windows.Forms.TabPage]::new('Backup'); $tk.BackColor = Color '#102034'
    $mk = { param($t, $x, $y) $l = [Windows.Forms.Label]::new(); $l.Text = $t; $l.AutoSize = $true; $l.Location = [Drawing.Point]::new($x, $y); $l }
    $l1 = & $mk 'Origem (pasta ou arquivo a copiar):' 14 8
    $l2 = & $mk 'Destino (pasta, unidade ou \\servidor\compartilhamento):' 14 60
    $tbS = [Windows.Forms.TextBox]::new(); $tbS.Location = [Drawing.Point]::new(14, 30); $tbS.Width = 560
    $tbD = [Windows.Forms.TextBox]::new(); $tbD.Location = [Drawing.Point]::new(14, 82); $tbD.Width = 560
    foreach ($t in @($tbS, $tbD)) { $t.BackColor = Color '#192F47'; $t.ForeColor = Color '#E8F1FF'; $t.BorderStyle = 'FixedSingle' }
    if ($last) { $tbS.Text = "$($last.origem)"; $tbD.Text = "$($last.destino)" }
    $mb = { param($t, $x, $y, $w, $h) $b = [Windows.Forms.Button]::new(); $b.Text = $t; $b.Width = $w; $b.Height = $h; $b.Location = [Drawing.Point]::new($x, $y); $b.FlatStyle = 'Flat'; $b.FlatAppearance.BorderSize = 0; $b.BackColor = Color '#284763'; $b.ForeColor = Color '#E8F1FF'; $b }
    $bpS = & $mb 'Procurar...' 584 27 120 30
    $bpD = & $mb 'Procurar...' 584 79 120 30
    $bpS.Add_Click({ $fb = New-Object Windows.Forms.FolderBrowserDialog; if ($fb.ShowDialog() -eq 'OK') { $tbS.Text = $fb.SelectedPath } }.GetNewClosure())
    $bpD.Add_Click({ $fb = New-Object Windows.Forms.FolderBrowserDialog; $fb.ShowNewFolderButton = $true; if ($fb.ShowDialog() -eq 'OK') { $tbD.Text = $fb.SelectedPath } }.GetNewClosure())
    $mc = { param($t, $x, $y, $on) $c = [Windows.Forms.CheckBox]::new(); $c.Text = $t; $c.AutoSize = $true; $c.Location = [Drawing.Point]::new($x, $y); $c.Checked = $on; $c }
    $ckS = & $mc 'Criar subpasta com nome da origem e data/hora no destino' 14 116 $true
    $ckP = & $mc 'Incluir arquivos protegidos (modo /B, exige administrador)' 14 140 ([bool]$isAdmin)
    $ckP.Enabled = [bool]$isAdmin
    $ckM = & $mc 'Espelhar (APAGA no destino o que não existe mais na origem)' 14 164 $false
    $bgo = & $mb 'Fazer backup' 14 192 170 38
    $bcn = & $mb 'Cancelar' 194 192 110 38; $bcn.Enabled = $false
    $bop = & $mb 'Abrir destino' 314 192 130 38
    $bst = [Windows.Forms.Label]::new(); $bst.Location = [Drawing.Point]::new(460, 190); $bst.Size = [Drawing.Size]::new(440, 32); $bst.ForeColor = Color '#9FB6CF'
    $pb = [Windows.Forms.ProgressBar]::new(); $pb.Style = 'Marquee'; $pb.MarqueeAnimationSpeed = 30; $pb.Location = [Drawing.Point]::new(460, 224); $pb.Size = [Drawing.Size]::new(440, 8); $pb.Visible = $false
    $tm = New-Object Windows.Forms.Timer; $tm.Interval = 1000
    $tm.Add_Tick({
        $job = $global:bk
        if (-not $job) { $tm.Stop(); return }
        $ln = (Read-LogTail $job.Log 1) -replace '\s+', ' '
        if ($ln.Length -gt 95) { $ln = $ln.Substring(0, 95) + '...' }
        if ($ln) { $bst.Text = $ln.Trim() }
        if ($job.Proc.HasExited) {
            $tm.Stop(); $pb.Visible = $false; $bgo.Enabled = $true; $bcn.Enabled = $false
            $code = $job.Proc.ExitCode
            $msg = if ($job.Cancelled) { 'CANCELADO (a cópia parcial permanece no destino)' } elseif ($code -ge 16) { 'ERRO FATAL (verifique origem, destino e permissões)' } elseif ($code -ge 8) { 'concluído COM FALHAS em alguns arquivos' } else { 'concluído com sucesso' }
            $bst.Text = "Backup $msg"
            Log "`r`n=== Backup $msg (código $code, $(Get-Date -Format 'HH:mm:ss')) ==="
            Log (Read-LogTail $job.Log 14)
            Log "Origem : $($job.Src)`r`nDestino: $($job.Dest)`r`nLog completo: $($job.Log)"
            $global:bk = $null
        }
    }.GetNewClosure())
    $bgo.Add_Click({
        try {
            if ($global:bk) { Log 'Já existe um backup em andamento.'; return }
            if ($ckM.Checked) {
                $r = [Windows.Forms.MessageBox]::Show("O modo espelho APAGA no destino os arquivos que não existem mais na origem.`n`nContinuar?", 'Backup - modo espelho', 'YesNo', 'Warning')
                if ($r -ne 'Yes') { return }
            }
            $job = Start-Backup $tbS.Text $tbD.Text $ckS.Checked $ckP.Checked $ckM.Checked
            $global:bk = $job; $global:bkDest = $job.Dest
            try { New-Item (Split-Path $bkCfg) -ItemType Directory -Force | Out-Null; @{ origem = $tbS.Text; destino = $tbD.Text } | ConvertTo-Json | Out-File $bkCfg -Encoding UTF8 } catch { }
            Log "`r`n=== Backup iniciado ($(Get-Date -Format 'HH:mm:ss')) ===`r`nOrigem : $($job.Src)`r`nDestino: $($job.Dest)"
            $bst.Text = 'Copiando...'; $bgo.Enabled = $false; $bcn.Enabled = $true; $pb.Visible = $true; $tm.Start()
        } catch { Log "ERRO: $($_.Exception.Message)" }
    }.GetNewClosure())
    $bcn.Add_Click({ $job = $global:bk; if ($job -and -not $job.Proc.HasExited) { $job.Cancelled = $true; try { $job.Proc.Kill() } catch { } } }.GetNewClosure())
    $bop.Add_Click({
        $d = if ($global:bkDest) { $global:bkDest } else { $tbD.Text.Trim().Trim('"') }
        if ($d -and (Test-Path -LiteralPath $d)) { Start-Process explorer.exe $d } else { Log 'O destino ainda não existe.' }
    }.GetNewClosure())
    $tk.Controls.AddRange(@($l1, $tbS, $bpS, $l2, $tbD, $bpD, $ckS, $ckP, $ckM, $bgo, $bcn, $bop, $bst, $pb)); [void]$tabs.TabPages.Add($tk)
    $ta = [Windows.Forms.TabPage]::new('Active Directory'); $ta.BackColor = Color '#102034'
    $lb = [Windows.Forms.Label]::new(); $lb.Text = 'E-mail, matrícula, telefone ou login do usuário:'; $lb.AutoSize = $true; $lb.Location = [Drawing.Point]::new(14, 16)
    $tb = [Windows.Forms.TextBox]::new(); $tb.Location = [Drawing.Point]::new(14, 44); $tb.Width = 430
    $tb.BackColor = Color '#192F47'; $tb.ForeColor = Color '#E8F1FF'; $tb.BorderStyle = 'FixedSingle'
    $ck = [Windows.Forms.CheckBox]::new(); $ck.Text = 'Mostrar só o resumo (sem todas as propriedades)'; $ck.AutoSize = $true; $ck.Location = [Drawing.Point]::new(14, 82)
    $nt = [Windows.Forms.Label]::new(); $nt.Text = 'Consulta somente leitura, com a conta que está logada. Exige computador no domínio e rede/VPN.'; $nt.AutoSize = $true; $nt.Location = [Drawing.Point]::new(14, 118); $nt.ForeColor = Color '#9FB6CF'
    $bs = [Windows.Forms.Button]::new(); $bs.Text = 'Buscar usuário'; $bs.Width = 190; $bs.Height = 42; $bs.Location = [Drawing.Point]::new(460, 38)
    $bs.FlatStyle = 'Flat'; $bs.FlatAppearance.BorderSize = 0; $bs.BackColor = Color '#284763'; $bs.ForeColor = Color '#E8F1FF'
    $bs.Add_Click({
        $q = $tb.Text.Trim()
        if (-not $q) { Log 'Digite um e-mail, matrícula, telefone ou login.'; return }
        Run "Active Directory: $q" { param($x, $y) Find-AdUser $x $y } @($q, $ck.Checked)
    }.GetNewClosure())
    $tb.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { $bs.PerformClick(); $_.SuppressKeyPress = $true } }.GetNewClosure())
    $ta.Controls.AddRange(@($lb, $tb, $ck, $nt, $bs)); [void]$tabs.TabPages.Add($ta)
    $bar = [Windows.Forms.FlowLayoutPanel]::new(); $bar.Dock = 'Bottom'; $bar.Height = 46; $bar.Padding = [Windows.Forms.Padding]::new(5)
    $bHtml = [Windows.Forms.Button]::new(); $bHtml.Text = 'Exportar tela (HTML)'; $bHtml.Width = 180; $bHtml.Height = 34
    $bHtml.Add_Click({
        $f = "$(Get-ReportDir)\Tela_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmm').html"
        $t = [Net.WebUtility]::HtmlEncode($global:log.Text)
        "<!DOCTYPE html><html lang='pt-BR'><head><meta charset='utf-8'><title>$env:COMPUTERNAME</title><style>body{margin:0;padding:24px;background:#1b1b1b;color:#ececec;font-family:Segoe UI,sans-serif}pre{background:#262626;border:1px solid #3a3a3a;border-radius:14px;padding:16px;font:13px Consolas,monospace;white-space:pre-wrap}</style></head><body><h1>$env:COMPUTERNAME</h1><div style='color:#9a9a9a'>Gerado em $(Get-Date -Format 'dd/MM/yyyy HH:mm')</div><pre>$t</pre></body></html>" | Out-File $f -Encoding UTF8
        Start-Process $f; Log "Salvo em $f"
    })
    $bClr = [Windows.Forms.Button]::new(); $bClr.Text = 'Limpar tela'; $bClr.Width = 120; $bClr.Height = 34
    $bClr.Add_Click({ $global:log.Clear() })
    $bar.Controls.AddRange(@($bHtml, $bClr))
    $tf.Controls.Add($tl); $tf.Controls.Add($bar); $tf.Controls.Add($tabs)
    [void]$tf.ShowDialog(); $tf.Dispose()
}

$form = [Windows.Forms.Form]::new()
$form.Text = 'Field Toolkit | Diagnóstico e manutenção'
$form.StartPosition = 'CenterScreen'
$form.ClientSize = [Drawing.Size]::new(820,650)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.AutoScaleMode = 'Dpi'
$form.AutoScroll = $true
$form.BackColor = Color '#102034'
$form.ForeColor = Color '#E8F1FF'
$form.Font = [Drawing.Font]::new('Segoe UI',10)

function Add-Control($control, $caption, $x, $y, $width, $height) {
    $control.Text = $caption
    $control.SetBounds($x,$y,$width,$height)
    $form.Controls.Add($control)
    $control
}

function Label($caption, $x, $y, $width, $height) {
    Add-Control ([Windows.Forms.Label]::new()) $caption $x $y $width $height
}

function Button($caption, $x, $y, $width, $background, $foreground) {
    $button = Add-Control ([Windows.Forms.Button]::new()) $caption $x $y $width 39
    $button.FlatStyle = 'Flat'
    $button.FlatAppearance.BorderSize = 0
    $button.BackColor = Color $background
    $button.ForeColor = Color $foreground
    $button
}

$title = Label 'Field Toolkit' 24 19 770 42
$title.Font = [Drawing.Font]::new('Segoe UI',24,[Drawing.FontStyle]::Bold)
$subtitle = if ($Demo) { 'Modo demonstração · dados fictícios' } else { 'Equipe Field  /  Diagnóstico, limpeza e manutenção do Windows' }
[void](Label $subtitle 27 69 760 24)
[void](Label 'Ordem de serviço *' 27 113 212 25)
[void](Label 'Cliente' 245 113 270 25)
[void](Label 'Técnico' 523 113 270 25)
$orderBox = Add-Control ([Windows.Forms.TextBox]::new()) '' 27 139 203 28
$customerBox = Add-Control ([Windows.Forms.TextBox]::new()) '' 245 139 260 28
$techBox = Add-Control ([Windows.Forms.TextBox]::new()) '' 523 139 271 28
foreach ($box in @($orderBox,$customerBox,$techBox)) {
    $box.BackColor = Color '#213953'
    $box.ForeColor = Color '#E8F1FF'
    $box.BorderStyle = 'FixedSingle'
    $box.MaxLength = 120
}
if ($Demo) { $orderBox.Text = 'DEMO-001' }
$sampleBox = Add-Control ([Windows.Forms.CheckBox]::new()) 'Ler 128 blocos do disco do Windows (requer administrador)' 27 179 600 25
$sampleBox.Enabled = $isAdmin -and !$Demo
$sampleBox.Checked = $sampleBox.Enabled
$start = Button 'Iniciar diagnóstico' 27 219 181 '#73E2CB' '#102034'
$reportButton = Button 'Abrir último laudo' 221 219 182 '#284763' '#E8F1FF'
$historyButton = Button 'Histórico de atendimentos' 416 219 218 '#284763' '#E8F1FF'
$toolsButton = Button 'Ferramentas' 647 219 147 '#284763' '#E8F1FF'
$logoButton = Button 'Logo do laudo' 640 172 154 '#284763' '#E8F1FF'
$reportButton.Enabled = $false
$progress = Add-Control ([Windows.Forms.ProgressBar]::new()) '' 27 280 767 13
$statusLabel = Label 'Aguardando início' 27 301 767 25
$log = Add-Control ([Windows.Forms.TextBox]::new()) '' 27 334 767 139
$log.Multiline = $true
$log.ReadOnly = $true
$log.ScrollBars = 'Vertical'
$log.BackColor = Color '#192F47'
$log.ForeColor = Color '#E8F1FF'
[void](Label 'Mapa de amostras do disco do Windows' 27 489 767 23)
$tip = [Windows.Forms.ToolTip]::new()
$tiles = foreach ($i in 0..127) {
    $x = 27 + ($i % 32) * 24
    $y = 520 + [int][Math]::Floor($i / 32) * 21
    $tile = Add-Control ([Windows.Forms.Panel]::new()) '' $x $y 20 17
    $tile.BackColor = Color '#32475C'
    $tip.SetToolTip($tile,'Não testado')
    $tile
}
[void](Label 'Verde = lido    Amarelo = lento    Vermelho = erro    Cinza = não testado' 27 609 767 25)
$session = @{ State=$null; Pipeline=$null; Async=$null; Report=$null; Result=$null }

$timer = [Windows.Forms.Timer]::new()
$timer.Interval = 120
$timer.Add_Tick({
    $state = $session.State
    $progress.Value = $state.Progress
    $message = $null
    while ($state.Messages.TryDequeue([ref]$message)) {
        $statusLabel.Text = [string]$message
        $log.AppendText("$message`r`n")
    }
    $cell = $null
    while ($state.Cells.TryDequeue([ref]$cell)) {
        $tile = $tiles[$cell.Index]
        $tile.BackColor = switch ($cell.Status) {
            'OK' { Color '#22B486' }
            'Lento' { Color '#F3B94C' }
            default { Color '#EF6771' }
        }
        $tip.SetToolTip($tile, ("Amostra {0} · {1} ms · {2}" -f ($cell.Index+1),$cell.Ms,$cell.Detail))
    }
    if (!$session.Async.IsCompleted -or !$state.Cells.IsEmpty -or !$state.Messages.IsEmpty) { return }
    $timer.Stop()
    try {
        $output = $session.Pipeline.EndInvoke($session.Async)
        if ($output.Count -ne 1) { throw 'A coleta retornou um resultado inesperado. Consulte o registro abaixo.' }
        $session.Result = $output[0]
        $session.Report = Save-Report $session.Result $history
        $statusLabel.Text = 'HTML salvo em C:\IT\Diagnostico.'
        try { [void](Save-Snapshot) } catch { }
        $reportButton.Enabled = $true
        try { Start-Process -FilePath $session.Report }
        catch {
            $statusLabel.Text = 'HTML salvo. Use Abrir último laudo para tentar abrir o navegador.'
            $log.AppendText("$($_.Exception.Message)`r`n")
        }
    } catch {
        if ($null -ne $session.Result) {
            $statusLabel.Text = 'Falha ao gerar o laudo. Veja o detalhe abaixo e use Salvar laudo para tentar novamente.'
            $reportButton.Text = 'Salvar laudo'
            $reportButton.Enabled = $true
        } else { $statusLabel.Text = 'Diagnóstico interrompido' }
        $log.AppendText("$($_.Exception.Message)`r`n")
    } finally {
        $session.Pipeline.Dispose()
        $session.Pipeline = $null
        $session.Async = $null
        $start.Enabled = $true
    }
})

$start.Add_Click({
    if ([string]::IsNullOrWhiteSpace($orderBox.Text)) {
        $statusLabel.Text = 'Preencha o número da ordem de serviço.'
        [void]$orderBox.Focus()
        return
    }
    $start.Enabled = $false
    $reportButton.Enabled = $false
    $reportButton.Text = 'Abrir último laudo'
    $session.Result = $null
    $session.Report = $null
    $log.Clear()
    $progress.Value = 0
    foreach ($tile in $tiles) {
        $tile.BackColor = Color '#32475C'
        $tip.SetToolTip($tile,'Não testado')
    }
    $statusLabel.Text = 'Iniciando coleta...'
    $session.State = New-ScanState
    try {
        $session.Pipeline = [PowerShell]::Create()
        $pipeline = $session.Pipeline
        [void]$pipeline.AddScript($engine.ToString())
        foreach ($argument in @($session.State, $orderBox.Text.Trim(), $customerBox.Text.Trim(),
                $techBox.Text.Trim(), $sampleBox.Checked, $isAdmin, $Demo.IsPresent)) {
            [void]$pipeline.AddArgument($argument)
        }
        $session.Async = $pipeline.BeginInvoke()
        $timer.Start()
    } catch {
        if ($null -ne $session.Pipeline) { $session.Pipeline.Dispose(); $session.Pipeline = $null }
        $statusLabel.Text = 'Falha ao iniciar'
        $log.AppendText($_.Exception.Message)
        $start.Enabled = $true
    }
})

$reportButton.Add_Click({
    try {
        if ($null -eq $session.Report) {
            $session.Report = Save-Report $session.Result $history
            $reportButton.Text = 'Abrir último laudo'
            $statusLabel.Text = 'Laudo salvo.'
        }
        Start-Process -FilePath $session.Report
    } catch { $log.AppendText("$($_.Exception.Message)`r`n") }
})

$historyButton.Add_Click({
    try {
        [void][IO.Directory]::CreateDirectory($history)
        Start-Process -FilePath explorer.exe -ArgumentList ('"' + $history + '"')
    } catch { $log.AppendText("$($_.Exception.Message)`r`n") }
})

$toolsButton.Add_Click({ try { Show-Tools } catch { $log.AppendText("$($_.Exception.Message)`r`n") } })

$logoButton.Add_Click({ try { $statusLabel.Text = [string](Set-ReportLogo) } catch { $statusLabel.Text = 'Falha no logo: ' + $_.Exception.Message } })

$form.Add_FormClosing({
    param($sender,$eventArgs)
    if ($null -ne $session.Pipeline) {
        $eventArgs.Cancel = $true
        $statusLabel.Text = 'Aguarde o término da coleta para fechar.'
    }
})

try { [void]$form.ShowDialog() }
finally { $timer.Dispose(); $tip.Dispose(); $form.Dispose() }
