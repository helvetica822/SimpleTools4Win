Write-Host "######################################## Wake on LAN ########################################"
Write-Host "#                                                                                           #"
Write-Host "#     指定した MAC アドレスの端末に対してマジックパケットを送信して起動します。             #"
Write-Host "#                                                                                           #"
Write-Host "#     動作要件                                                                              #"
Write-Host "#       ・" -NoNewline
Write-Host "送信先端末の Wake on LAN が有効化されていること                " -NoNewline -ForegroundColor Red
Write-Host "                   #"
Write-Host "#       ・" -NoNewline
Write-Host "送信先端末が同一ネットワーク上に存在すること                 " -NoNewline -ForegroundColor Red
Write-Host "                     #"
Write-Host "#                                                                                           #"
Write-Host "#############################################################################################"
Write-Host

# ToDo : 使用する前に WoL 対象とする MAC アドレスを @macAddresses に配列で定義すること
$macAddresses = @()
$header = [byte[]](@(0xFF) * 6)

$client = New-Object System.Net.Sockets.UdpClient

try {
    foreach ($macAddress in $macAddresses) {
        try {
            Write-Host "送信先端末 MAC アドレス:" $macAddress

            $addr = [byte[]]($macAddress.split(":") | ForEach-Object { [Convert]::ToInt32($_, 16) })
            $magicpacket = $header + $addr * 16
            $broadcast = [System.Net.IPAddress]::Broadcast

            $client.Connect($broadcast, 2304)
            $client.Send($magicpacket, $magicpacket.Length) | Out-Null

            Write-Host " > マジックパケット送信中 ... " -NoNewline
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
