
function Write-InfoToHost([string]$title, [string]$value, [string]$unit, [int]$indentCount, [int]$tabCount) {
    $indents = " " * $indentCount
    $tabs = "`t" * $tabCount
    $u = $unit
    if (!($unit -eq $null -or $unit -eq "")) {
        $u = " $($u)"
    }

    Write-Host "$($indents)$($title)$($tabs): $($value)$($u)"
}

function Get-PropertyFromNetAdapter([string]$prop, [string]$jaName, [string]$enName) {
    $value = $null

    $rows = Get-NetAdapter | Select-Object Name, $prop

    foreach ($row in $rows) {
        if ($row.Name -eq $jaName -or $row.Name -eq $enName) {
            $value = $row.$prop
            break
        }
    }

    return $value
}

function Show-PcInfo() {
    $system = Get-CimInstance Win32_ComputerSystem
    $serial = (Get-CimInstance Win32_BIOS).SerialNumber

    Write-Host "■ PC情報"
    Write-InfoToHost "PC名" $system.Name "" 1 4
    Write-InfoToHost "ドメイン" $system.Domain "" 1 3
    Write-InfoToHost "メーカー" $system.Manufacturer "" 1 3
    Write-InfoToHost "製品名" $system.Model "" 1 4
    Write-InfoToHost "S/N" $serial "" 1 4
    Write-Host
}

function Show-OsInfo() {
    $comp = Get-ComputerInfo
    $osVersion = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'

    Write-Host "■ OS情報"
    Write-InfoToHost "製品名" $comp.WindowsProductName "" 1 4
    Write-InfoToHost "表示用バージョン" $osVersion.DisplayVersion "" 1 2
    Write-InfoToHost "バージョン" $comp.OsVersion "" 1 3
    Write-InfoToHost "ビルド番号" $comp.OsBuildNumber "" 1 3
    Write-InfoToHost "アーキテクチャ" $comp.OsArchitecture "" 1 3
    Write-InfoToHost "言語" $comp.OsLanguage "" 1 4
    Write-InfoToHost "キーボードレイアウト" $comp.KeyboardLayout "" 1 2
    Write-InfoToHost "タイムゾーン" $comp.TimeZone "" 1 3
    Write-Host
}

function Show-EthernetInfo($ipAddress) {
    $ipv4Ethernet = ($ipAddress | Where-Object { $_.AddressFamily -eq "IPv4" -and ($_.InterfaceAlias -eq "イーサネット" -or $_.InterfaceAlias -eq "Ethernet") }).IPAddress
    $ipv6Ethernet = ($ipAddress | Where-Object { $_.AddressFamily -eq "IPv6" -and ($_.InterfaceAlias -eq "イーサネット" -or $_.InterfaceAlias -eq "Ethernet") }).IPAddress
    $ethernetInterface = Get-PropertyFromNetAdapter "InterfaceDescription" "イーサネット" "Ethernet"
    $ethernetMacAddress = Get-PropertyFromNetAdapter "MacAddress" "イーサネット" "Ethernet"

    Write-Host "■ ネットワーク情報"
    Write-Host "【イーサネット】"
    Write-InfoToHost "製品名" $ethernetInterface "" 2 3
    Write-InfoToHost "IPv4アドレス" $ipv4Ethernet "" 2 3
    Write-InfoToHost "IPv6アドレス" $ipv6Ethernet "" 2 3
    Write-InfoToHost "MACアドレス" $ethernetMacAddress "" 2 3
    Write-Host
}

function Show-WifiInfo($ipAddress) {
    $ipv4Wifi = ($ipAddress | Where-Object { $_.AddressFamily -eq "IPv4" -and $_.InterfaceAlias -eq "Wi-Fi" }).IPAddress
    $ipv6Wifi = ($ipAddress | Where-Object { $_.AddressFamily -eq "IPv6" -and $_.InterfaceAlias -eq "Wi-Fi" }).IPAddress
    $wifiInterface = Get-PropertyFromNetAdapter "InterfaceDescription" "Wi-Fi" "Wi-Fi"
    $wifiMacAddress = Get-PropertyFromNetAdapter "MacAddress" "Wi-Fi" "Wi-Fi"

    Write-Host "【Wi-Fi】"
    Write-InfoToHost "製品名" $wifiInterface "" 2 3
    Write-InfoToHost "IPv4アドレス" $ipv4Wifi "" 2 3
    Write-InfoToHost "IPv6アドレス" $ipv6Wifi "" 2 3
    Write-InfoToHost "MACアドレス" $wifiMacAddress "" 2 3
    Write-Host
}

function Show-NetworkInfo() {
    $ipAddress = Get-NetIPAddress

    Show-EthernetInfo $ipAddress
    Show-WifiInfo $ipAddress
}

function Show-CpuInfo() {
    $cpu = Get-CimInstance Win32_Processor
    $clockSpeed = "{0:f2}" -f ([double]($cpu.MaxClockSpeed / 1000))

    Write-Host "■ CPU情報"
    Write-InfoToHost "製品名" $cpu.Name "" 1 4
    Write-InfoToHost "コア数" $cpu.NumberOfCores "" 1 4
    Write-InfoToHost "論理プロセッサ数" $cpu.NumberOfLogicalProcessors "" 1 2
    Write-InfoToHost "最大クロック数" $clockSpeed "GHz" 1 3
    Write-Host
}

function Show-GpuInfo() {
    $gpu = (Get-CimInstance Win32_VideoController).Name
    $seq = 1

    Write-Host "■ GPU情報"

    if ($gpu -is [array]) {
        foreach ($g in $gpu) {
            Write-Host "【GPU $($seq)】"
            Write-InfoToHost "製品名" $g "" 2 3
            Write-InfoToHost "VRAM" "取得不可" "" 2 4
            Write-Host

            $seq++
        }
    }
    else {
        Write-Host "【GPU $($seq)】"
        Write-InfoToHost "製品名" $gpu "" 2 3
        Write-InfoToHost "VRAM" "取得不可" "" 2 4
        Write-Host
    }
}

function Show-RamInfo() {
    $ram = "{0:f2}" -f ([double](((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory) / 1GB))

    Write-Host "■ メモリ情報"
    Write-InfoToHost "RAM" $ram "GB" 1 4
    Write-Host
}

function Show-EachDiskInfo($d, [int]$seq) {
    $size = "{0:f2}" -f ([double]($d.AllocatedSize / 1GB))

    Write-Host "【ディスク $($seq)】"
    Write-InfoToHost "製品名" $d.FriendlyName "" 2 3
    Write-InfoToHost "ディスク種別" $d.MediaType "" 2 3
    Write-InfoToHost "割り当てサイズ" $size "GB" 2 2
    Write-Host
}

function Show-DiskInfo() {
    $disk = Get-PhysicalDisk | Select-Object FriendlyName, MediaType, AllocatedSize
    $seq = 1

    Write-Host "■ ディスク情報"
 
    if ($disk -is [array]) {
        foreach ($d in $disk) {
            Show-EachDiskInfo $d $seq
            $seq++
        }
    }
    else {
        Show-EachDiskInfo $disk $seq
    }
}

Show-PcInfo
Show-OsInfo
Show-NetworkInfo
Show-CpuInfo
Show-GpuInfo
Show-RamInfo
Show-DiskInfo
