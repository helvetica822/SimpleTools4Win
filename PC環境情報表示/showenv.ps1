
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

    Write-Host "�� PC���"
    Write-InfoToHost "PC��" $system.Name "" 1 4
    Write-InfoToHost "�h���C��" $system.Domain "" 1 3
    Write-InfoToHost "���[�J�[" $system.Manufacturer "" 1 3
    Write-InfoToHost "���i��" $system.Model "" 1 4
    Write-InfoToHost "S/N" $serial "" 1 4
    Write-Host
}

function Show-OsInfo() {
    $comp = Get-ComputerInfo
    $osVersion = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'

    Write-Host "�� OS���"
    Write-InfoToHost "���i��" $comp.WindowsProductName "" 1 4
    Write-InfoToHost "�\���p�o�[�W����" $osVersion.DisplayVersion "" 1 2
    Write-InfoToHost "�o�[�W����" $comp.OsVersion "" 1 3
    Write-InfoToHost "�r���h�ԍ�" $comp.OsBuildNumber "" 1 3
    Write-InfoToHost "�A�[�L�e�N�`��" $comp.OsArchitecture "" 1 3
    Write-InfoToHost "����" $comp.OsLanguage "" 1 4
    Write-InfoToHost "�L�[�{�[�h���C�A�E�g" $comp.KeyboardLayout "" 1 2
    Write-InfoToHost "�^�C���]�[��" $comp.TimeZone "" 1 3
    Write-Host
}

function Show-EthernetInfo($ipAddress) {
    $ipv4Ethernet = ($ipAddress | Where-Object { $_.AddressFamily -eq "IPv4" -and ($_.InterfaceAlias -eq "�C�[�T�l�b�g" -or $_.InterfaceAlias -eq "Ethernet") }).IPAddress
    $ipv6Ethernet = ($ipAddress | Where-Object { $_.AddressFamily -eq "IPv6" -and ($_.InterfaceAlias -eq "�C�[�T�l�b�g" -or $_.InterfaceAlias -eq "Ethernet") }).IPAddress
    $ethernetInterface = Get-PropertyFromNetAdapter "InterfaceDescription" "�C�[�T�l�b�g" "Ethernet"
    $ethernetMacAddress = Get-PropertyFromNetAdapter "MacAddress" "�C�[�T�l�b�g" "Ethernet"

    Write-Host "�� �l�b�g���[�N���"
    Write-Host "�y�C�[�T�l�b�g�z"
    Write-InfoToHost "���i��" $ethernetInterface "" 2 3
    Write-InfoToHost "IPv4�A�h���X" $ipv4Ethernet "" 2 3
    Write-InfoToHost "IPv6�A�h���X" $ipv6Ethernet "" 2 3
    Write-InfoToHost "MAC�A�h���X" $ethernetMacAddress "" 2 3
    Write-Host
}

function Show-WifiInfo($ipAddress) {
    $ipv4Wifi = ($ipAddress | Where-Object { $_.AddressFamily -eq "IPv4" -and $_.InterfaceAlias -eq "Wi-Fi" }).IPAddress
    $ipv6Wifi = ($ipAddress | Where-Object { $_.AddressFamily -eq "IPv6" -and $_.InterfaceAlias -eq "Wi-Fi" }).IPAddress
    $wifiInterface = Get-PropertyFromNetAdapter "InterfaceDescription" "Wi-Fi" "Wi-Fi"
    $wifiMacAddress = Get-PropertyFromNetAdapter "MacAddress" "Wi-Fi" "Wi-Fi"

    Write-Host "�yWi-Fi�z"
    Write-InfoToHost "���i��" $wifiInterface "" 2 3
    Write-InfoToHost "IPv4�A�h���X" $ipv4Wifi "" 2 3
    Write-InfoToHost "IPv6�A�h���X" $ipv6Wifi "" 2 3
    Write-InfoToHost "MAC�A�h���X" $wifiMacAddress "" 2 3
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

    Write-Host "�� CPU���"
    Write-InfoToHost "���i��" $cpu.Name "" 1 4
    Write-InfoToHost "�R�A��" $cpu.NumberOfCores "" 1 4
    Write-InfoToHost "�_���v���Z�b�T��" $cpu.NumberOfLogicalProcessors "" 1 2
    Write-InfoToHost "�ő�N���b�N��" $clockSpeed "GHz" 1 3
    Write-Host
}

function Show-GpuInfo() {
    $gpu = (Get-CimInstance Win32_VideoController).Name
    $seq = 1

    Write-Host "�� GPU���"

    if ($gpu -is [array]) {
        foreach ($g in $gpu) {
            Write-Host "�yGPU $($seq)�z"
            Write-InfoToHost "���i��" $g "" 2 3
            Write-InfoToHost "VRAM" "�擾�s��" "" 2 4
            Write-Host

            $seq++
        }
    }
    else {
        Write-Host "�yGPU $($seq)�z"
        Write-InfoToHost "���i��" $gpu "" 2 3
        Write-InfoToHost "VRAM" "�擾�s��" "" 2 4
        Write-Host
    }
}

function Show-RamInfo() {
    $ram = "{0:f2}" -f ([double](((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory) / 1GB))

    Write-Host "�� ���������"
    Write-InfoToHost "RAM" $ram "GB" 1 4
    Write-Host
}

function Show-EachDiskInfo($d, [int]$seq) {
    $size = "{0:f2}" -f ([double]($d.AllocatedSize / 1GB))

    Write-Host "�y�f�B�X�N $($seq)�z"
    Write-InfoToHost "���i��" $d.FriendlyName "" 2 3
    Write-InfoToHost "�f�B�X�N���" $d.MediaType "" 2 3
    Write-InfoToHost "���蓖�ăT�C�Y" $size "GB" 2 2
    Write-Host
}

function Show-DiskInfo() {
    $disk = Get-PhysicalDisk | Select-Object FriendlyName, MediaType, AllocatedSize
    $seq = 1

    Write-Host "�� �f�B�X�N���"
 
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
