# EDID Override + ドライバリスタート スクリプト (管理者権限で実行)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$paramPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\DISPLAY\HJW2130\4&7d3369b&0&UID206371\Device Parameters"

# 1. EDID Override 書き込み
$edid = [byte[]](Get-ItemProperty -Path $paramPath -Name EDID).EDID
Write-Host "Original EDID length: $($edid.Length)"

# 2048x2732 @ 30Hz DTD
$width = 2048; $height = 2732
$hBlank = 160; $vBlank = 68
$hFront = 48; $hSync = 32; $vFront = 3; $vSync = 10
$hTotal = [double]$width + $hBlank
$vTotal = [double]$height + $vBlank
$pixClk = [int][math]::Round($hTotal * $vTotal * 30.0 / 10000.0)

$dtd = [byte[]]::new(18)
$dtd[0]  = [byte]($pixClk -band 0xFF)
$dtd[1]  = [byte](($pixClk -shr 8) -band 0xFF)
$dtd[2]  = [byte]($width -band 0xFF)
$dtd[3]  = [byte]($hBlank -band 0xFF)
$dtd[4]  = [byte]((($width -shr 8) -band 0x0F) -shl 4 -bor (($hBlank -shr 8) -band 0x0F))
$dtd[5]  = [byte]($height -band 0xFF)
$dtd[6]  = [byte]($vBlank -band 0xFF)
$dtd[7]  = [byte]((($height -shr 8) -band 0x0F) -shl 4 -bor (($vBlank -shr 8) -band 0x0F))
$dtd[8]  = [byte]($hFront -band 0xFF)
$dtd[9]  = [byte]($hSync -band 0xFF)
$dtd[10] = [byte]((($vFront -band 0x0F) -shl 4) -bor ($vSync -band 0x0F))
$dtd[11] = [byte]0x00
$dtd[12] = [byte]0x00; $dtd[13] = [byte]0x00; $dtd[14] = [byte]0x00
$dtd[15] = [byte]0x00; $dtd[16] = [byte]0x00
$dtd[17] = [byte]0x18

Write-Host "DTD: ${width}x${height} @ 30Hz, PixClk=$($pixClk/100.0)MHz"

for ($i = 0; $i -lt 18; $i++) { $edid[90 + $i] = $dtd[$i] }
$sum = 0; for ($i = 0; $i -lt 127; $i++) { $sum += $edid[$i] }
$edid[127] = [byte]((256 - ($sum % 256)) % 256)

Set-ItemProperty -Path $paramPath -Name "EDID_OVERRIDE" -Value $edid -Type Binary
Write-Host "EDID_OVERRIDE written."

# 2. モニターデバイスを無効化→再有効化してドライバにEDIDを再読み込みさせる
Write-Host ""
Write-Host "Restarting monitor device to reload EDID..."
$monitorId = "DISPLAY\HJW2130\4&7D3369B&0&UID206371"
Disable-PnpDevice -InstanceId $monitorId -Confirm:$false
Start-Sleep -Seconds 2
Enable-PnpDevice -InstanceId $monitorId -Confirm:$false
Start-Sleep -Seconds 2
Write-Host "Monitor device restarted."

# 3. GPUドライバも再起動（これでディスプレイモード一覧が更新される）
Write-Host "Restarting GPU driver..."
$gpuId = "PCI\VEN_8086&DEV_3EA5&SUBSYS_22128086&REV_01\3&11583659&0&10"
Disable-PnpDevice -InstanceId $gpuId -Confirm:$false
Start-Sleep -Seconds 3
Enable-PnpDevice -InstanceId $gpuId -Confirm:$false
Start-Sleep -Seconds 3
Write-Host "GPU driver restarted."

Write-Host ""
Write-Host "Done! Check if 2048x2732 appears in supported resolutions."
Write-Host ""
Read-Host "Press Enter to close"
