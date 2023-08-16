Write-Host "######################################## Wake on LAN ########################################"
Write-Host "#                                                                                           #"
Write-Host "#     �w�肵�� MAC �A�h���X�̒[���ɑ΂��ă}�W�b�N�p�P�b�g�𑗐M���ċN�����܂��B             #"
Write-Host "#                                                                                           #"
Write-Host "#     ����v��                                                                              #"
Write-Host "#       �E" -NoNewline
Write-Host "���M��[���� Wake on LAN ���L��������Ă��邱��                " -NoNewline -ForegroundColor Red
Write-Host "                   #"
Write-Host "#       �E" -NoNewline
Write-Host "���M��[��������l�b�g���[�N��ɑ��݂��邱��                 " -NoNewline -ForegroundColor Red
Write-Host "                     #"
Write-Host "#                                                                                           #"
Write-Host "#############################################################################################"
Write-Host

# ToDo : �g�p����O�� WoL �ΏۂƂ��� MAC �A�h���X�� @macAddresses �ɔz��Œ�`���邱��
$macAddresses = @()
$header = [byte[]](@(0xFF) * 6)

$client = New-Object System.Net.Sockets.UdpClient

try {
    foreach ($macAddress in $macAddresses) {
        try {
            Write-Host "���M��[�� MAC �A�h���X:" $macAddress

            $addr = [byte[]]($macAddress.split(":") | ForEach-Object { [Convert]::ToInt32($_, 16) })
            $magicpacket = $header + $addr * 16
            $broadcast = [System.Net.IPAddress]::Broadcast

            $client.Connect($broadcast, 2304)
            $client.Send($magicpacket, $magicpacket.Length) | Out-Null

            Write-Host " > �}�W�b�N�p�P�b�g���M�� ... " -NoNewline
            Write-Host "done" -ForegroundColor Blue
        }
        catch {
            Write-Host "error" -ForegroundColor Red
        }

        Write-Host 
    }
}
finally {
    if ($null -eq $client) {
        $client.Close()
    }
}
