# 管理者権限チェック＆自動昇格
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class DisplayAPI {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DEVMODE {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public ushort dmSpecVersion;
        public ushort dmDriverVersion;
        public ushort dmSize;
        public ushort dmDriverExtra;
        public uint dmFields;
        public int dmPositionX;
        public int dmPositionY;
        public uint dmDisplayOrientation;
        public uint dmDisplayFixedOutput;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;
        public ushort dmLogPixels;
        public uint dmBitsPerPel;
        public uint dmPelsWidth;
        public uint dmPelsHeight;
        public uint dmDisplayFlags;
        public uint dmDisplayFrequency;
        public uint dmICMMethod;
        public uint dmICMIntent;
        public uint dmMediaType;
        public uint dmDitherType;
        public uint dmReserved1;
        public uint dmReserved2;
        public uint dmPanningWidth;
        public uint dmPanningHeight;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool EnumDisplaySettingsW(string lpszDeviceName, int iModeNum, ref DEVMODE lpDevMode);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int ChangeDisplaySettingsExW(string lpszDeviceName, ref DEVMODE lpDevMode, IntPtr hwnd, uint dwflags, IntPtr lParam);

    public const uint DM_PELSWIDTH = 0x80000;
    public const uint DM_PELSHEIGHT = 0x100000;
    public const uint DM_DISPLAYFREQUENCY = 0x400000;
    public const uint CDS_UPDATEREGISTRY = 0x01;
    public const uint CDS_TEST = 0x02;
    public const int ENUM_CURRENT_SETTINGS = -1;
}
"@

function Get-Displays {
    $displays = @()
    $screens = [System.Windows.Forms.Screen]::AllScreens
    foreach ($scr in $screens) {
        $displays += [PSCustomObject]@{
            Name        = $scr.DeviceName
            Description = $scr.DeviceName
            Primary     = $scr.Primary
            Bounds      = $scr.Bounds
        }
    }
    return $displays
}

function Get-CurrentResolution($deviceName) {
    $dm = New-Object DisplayAPI+DEVMODE
    $dm.dmSize = [uint16][System.Runtime.InteropServices.Marshal]::SizeOf([type][DisplayAPI+DEVMODE])
    if ([DisplayAPI]::EnumDisplaySettingsW($deviceName, [DisplayAPI]::ENUM_CURRENT_SETTINGS, [ref]$dm)) {
        return @{ Width = $dm.dmPelsWidth; Height = $dm.dmPelsHeight; Freq = $dm.dmDisplayFrequency }
    }
    return $null
}

function Get-SupportedModes($deviceName) {
    $modes = @{}
    $i = 0
    while ($true) {
        $dm = New-Object DisplayAPI+DEVMODE
        $dm.dmSize = [uint16][System.Runtime.InteropServices.Marshal]::SizeOf([type][DisplayAPI+DEVMODE])
        if (-not [DisplayAPI]::EnumDisplaySettingsW($deviceName, $i, [ref]$dm)) { break }
        $key = "$($dm.dmPelsWidth)x$($dm.dmPelsHeight)@$($dm.dmDisplayFrequency)/$($dm.dmBitsPerPel)"
        if (-not $modes.ContainsKey($key)) {
            $modes[$key] = [PSCustomObject]@{
                Width  = $dm.dmPelsWidth
                Height = $dm.dmPelsHeight
                Freq   = $dm.dmDisplayFrequency
                Bpp    = $dm.dmBitsPerPel
            }
        }
        $i++
    }
    return $modes.Values | Sort-Object { $_.Width * $_.Height } -Descending
}

function Get-GCD($a, $b) {
    while ($b -ne 0) { $t = $b; $b = $a % $b; $a = $t }
    return $a
}

function Get-RatioStr($w, $h) {
    $g = Get-GCD $w $h
    return "$([int]($w / $g)):$([int]($h / $g))"
}

function Build-DTD($width, $height, $refreshRate) {
    # EDID Detailed Timing Descriptor (18 bytes) を生成
    # CVT-like blanking intervals
    $hBlank = 160; $hFront = 48; $hSync = 32
    $vBlank = 68;  $vFront = 3;  $vSync = 10
    $hTotal = [double]$width + $hBlank
    $vTotal = [double]$height + $vBlank
    # ピクセルクロック (10kHz単位): hTotal * vTotal * refreshRate / 10000
    $pixClk = [int][math]::Round($hTotal * $vTotal * [double]$refreshRate / 10000.0)

    $dtd = [byte[]]::new(18)
    $dtd[0] = $pixClk -band 0xFF
    $dtd[1] = ($pixClk -shr 8) -band 0xFF
    $dtd[2] = $width -band 0xFF
    $dtd[3] = $hBlank -band 0xFF
    $dtd[4] = (($width -shr 8) -band 0x0F) -shl 4 -bor (($hBlank -shr 8) -band 0x0F)
    $dtd[5] = $height -band 0xFF
    $dtd[6] = $vBlank -band 0xFF
    $dtd[7] = (($height -shr 8) -band 0x0F) -shl 4 -bor (($vBlank -shr 8) -band 0x0F)
    $dtd[8] = $hFront -band 0xFF
    $dtd[9] = $hSync -band 0xFF
    $dtd[10] = (($vFront -band 0x0F) -shl 4) -bor ($vSync -band 0x0F)
    $dtd[11] = 0x00
    $dtd[12] = 0x00; $dtd[13] = 0x00; $dtd[14] = 0x00
    $dtd[15] = 0x00; $dtd[16] = 0x00
    $dtd[17] = 0x18  # non-interlaced, digital separate sync
    return $dtd
}

function Register-CustomResolution($width, $height) {
    # UVCデバイス (HDMI TO USB) のEDIDにカスタム解像度をOverrideとして登録
    $monitorPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\DISPLAY\HJW2130"
    if (-not (Test-Path $monitorPath)) {
        return @{ Success = $false; Message = "UVC device (HJW2130) not found in registry" }
    }
    # アクティブなインスタンスを探す
    $instances = Get-ChildItem $monitorPath
    $targetPath = $null
    foreach ($inst in $instances) {
        $paramPath = Join-Path $inst.PSPath "Device Parameters"
        $edid = (Get-ItemProperty -Path $paramPath -Name EDID -EA SilentlyContinue).EDID
        if ($edid -and $edid.Length -ge 128) {
            $targetPath = $paramPath
            break
        }
    }
    if (-not $targetPath) {
        return @{ Success = $false; Message = "No EDID found for UVC device" }
    }

    $edid = [byte[]](Get-ItemProperty -Path $targetPath -Name EDID).EDID
    # バックアップ (EDID_BACKUP)
    $backup = (Get-ItemProperty -Path $targetPath -Name EDID_BACKUP -EA SilentlyContinue).EDID_BACKUP
    if (-not $backup) {
        Set-ItemProperty -Path $targetPath -Name "EDID_BACKUP" -Value $edid -Type Binary
    }

    # DTD#3 (offset 90-107) にカスタム解像度を挿入
    $dtd = Build-DTD $width $height 30  # UVCは30Hzが一般的
    for ($i = 0; $i -lt 18; $i++) {
        $edid[90 + $i] = $dtd[$i]
    }

    # チェックサム再計算 (byte 127)
    $sum = 0
    for ($i = 0; $i -lt 127; $i++) { $sum += $edid[$i] }
    $edid[127] = [byte]((256 - ($sum % 256)) % 256)

    # EDID_OVERRIDE として書き込み
    Set-ItemProperty -Path $targetPath -Name "EDID_OVERRIDE" -Value $edid -Type Binary

    return @{ Success = $true; Message = "EDID override written: ${width}x${height}. Reboot or re-plug UVC device to apply." }
}

function Remove-CustomResolution {
    $monitorPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\DISPLAY\HJW2130"
    if (-not (Test-Path $monitorPath)) {
        return @{ Success = $false; Message = "UVC device not found" }
    }
    $instances = Get-ChildItem $monitorPath
    foreach ($inst in $instances) {
        $paramPath = Join-Path $inst.PSPath "Device Parameters"
        $override = (Get-ItemProperty -Path $paramPath -Name EDID_OVERRIDE -EA SilentlyContinue).EDID_OVERRIDE
        if ($override) {
            Remove-ItemProperty -Path $paramPath -Name "EDID_OVERRIDE" -EA SilentlyContinue
            return @{ Success = $true; Message = "EDID override removed. Reboot or re-plug to restore original." }
        }
    }
    return @{ Success = $false; Message = "No EDID override found" }
}

function Set-DisplayResolution($deviceName, $width, $height, $freq) {
    $dm = New-Object DisplayAPI+DEVMODE
    $dm.dmSize = [uint16][System.Runtime.InteropServices.Marshal]::SizeOf([type][DisplayAPI+DEVMODE])
    $dm.dmPelsWidth = $width
    $dm.dmPelsHeight = $height
    $dm.dmFields = [DisplayAPI]::DM_PELSWIDTH -bor [DisplayAPI]::DM_PELSHEIGHT
    if ($freq -gt 0) {
        $dm.dmDisplayFrequency = $freq
        $dm.dmFields = $dm.dmFields -bor [DisplayAPI]::DM_DISPLAYFREQUENCY
    }
    $result = [DisplayAPI]::ChangeDisplaySettingsExW($deviceName, [ref]$dm, [IntPtr]::Zero, [DisplayAPI]::CDS_TEST, [IntPtr]::Zero)
    if ($result -ne 0) { return @{ Success = $false; Message = "Not supported (code: $result)" } }
    $result = [DisplayAPI]::ChangeDisplaySettingsExW($deviceName, [ref]$dm, [IntPtr]::Zero, [DisplayAPI]::CDS_UPDATEREGISTRY, [IntPtr]::Zero)
    if ($result -eq 0) { return @{ Success = $true; Message = "Resolution changed!" } }
    if ($result -eq 1) { return @{ Success = $true; Message = "Resolution changed. Restart required." } }
    return @{ Success = $false; Message = "Change failed (code: $result)" }
}

# --- GUI ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "iPad UVC Resolution Tool"
$form.Size = New-Object System.Drawing.Size(800, 800)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Meiryo UI", 9)

$y = 10

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "iPad UVC Resolution Tool"
$lblTitle.Font = New-Object System.Drawing.Font("Meiryo UI", 14, [System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(15, $y)
$lblTitle.AutoSize = $true
$form.Controls.Add($lblTitle)
$y += 40

$lblDisp = New-Object System.Windows.Forms.Label
$lblDisp.Text = "Display:"
$lblDisp.Location = New-Object System.Drawing.Point(15, $y)
$lblDisp.AutoSize = $true
$form.Controls.Add($lblDisp)
$y += 22

$cmbDisplay = New-Object System.Windows.Forms.ComboBox
$cmbDisplay.Location = New-Object System.Drawing.Point(15, $y)
$cmbDisplay.Size = New-Object System.Drawing.Size(620, 25)
$cmbDisplay.DropDownStyle = "DropDownList"
$form.Controls.Add($cmbDisplay)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh"
$btnRefresh.Location = New-Object System.Drawing.Point(645, $y)
$btnRefresh.Size = New-Object System.Drawing.Size(80, 25)
$form.Controls.Add($btnRefresh)
$y += 30

$lblCurrent = New-Object System.Windows.Forms.Label
$lblCurrent.Text = "Select a display"
$lblCurrent.Location = New-Object System.Drawing.Point(15, $y)
$lblCurrent.AutoSize = $true
$form.Controls.Add($lblCurrent)
$y += 30

# Presets
$grpPreset = New-Object System.Windows.Forms.GroupBox
$grpPreset.Text = "iPad Presets (click to apply)"
$grpPreset.Location = New-Object System.Drawing.Point(15, $y)
$grpPreset.Size = New-Object System.Drawing.Size(750, 180)
$form.Controls.Add($grpPreset)

$presets = @(
    @{ Name = "iPad Pro 13 (M4) Native"; W = 2752; H = 2064; Portrait = $false },
    @{ Name = "iPad Pro 13 (M4) Half"; W = 1376; H = 1032; Portrait = $false },
    @{ Name = "iPad Pro 12.9 / Air 13 Native"; W = 2732; H = 2048; Portrait = $false },
    @{ Name = "iPad Pro 12.9 / Air 13 Half"; W = 1366; H = 1024; Portrait = $false },
    @{ Name = "iPad Air 13 Portrait Native"; W = 2048; H = 2732; Portrait = $true },
    @{ Name = "iPad Air 13 Portrait Half"; W = 1024; H = 1366; Portrait = $true },
    @{ Name = "iPad 4:3 (2048x1536)"; W = 2048; H = 1536; Portrait = $false },
    @{ Name = "iPad 4:3 Half (1024x768)"; W = 1024; H = 768; Portrait = $false },
    @{ Name = "XGA 4:3 (1280x960)"; W = 1280; H = 960; Portrait = $false },
    @{ Name = "SXGA 4:3 (1400x1050)"; W = 1400; H = 1050; Portrait = $false }
)

$bx = 10; $by = 22
for ($pi = 0; $pi -lt $presets.Count; $pi++) {
    $p = $presets[$pi]
    $btn = New-Object System.Windows.Forms.Button
    $btnLabel = $p.Name + " (" + $p.W.ToString() + "x" + $p.H.ToString() + ")"
    $btn.Text = $btnLabel
    $btn.Location = New-Object System.Drawing.Point($bx, $by)
    $btn.Size = New-Object System.Drawing.Size(360, 28)
    $btn.Tag = $p
    $btn.Add_Click({
        $pr = $this.Tag
        $idx = $cmbDisplay.SelectedIndex
        if ($idx -lt 0) { [System.Windows.Forms.MessageBox]::Show("Select a display first", "Warning"); return }

        if ($pr.Portrait) {
            # 縦型: EDID Override で登録が必要
            $msg = $pr.Name + " " + $pr.W.ToString() + "x" + $pr.H.ToString() + "`n`nThis portrait resolution requires EDID override (registry modification).`nAfter registration, you need to re-plug the UVC device or reboot.`n`nRegister and apply?"
            $confirm = [System.Windows.Forms.MessageBox]::Show($msg, "Portrait Resolution - EDID Override", "YesNo", "Question")
            if ($confirm -ne "Yes") { return }
            $regRes = Register-CustomResolution $pr.W $pr.H
            if ($regRes.Success) {
                $lblStatus.Text = "EDID: " + $regRes.Message
                [System.Windows.Forms.MessageBox]::Show($regRes.Message + "`n`nPlease re-plug the UVC device, then click the preset button again to apply the resolution.", "EDID Override Registered", "OK", "Information")
            } else {
                $lblStatus.Text = "NG: " + $regRes.Message
                [System.Windows.Forms.MessageBox]::Show($regRes.Message + "`n`nNote: Administrator privileges may be required.", "Error")
            }
            # EDID登録後、解像度適用も試みる
            $devName = $script:displayList[$idx].Name
            $res = Set-DisplayResolution $devName $pr.W $pr.H 0
            if ($res.Success) {
                $lblStatus.Text = "OK: " + $res.Message
                [System.Windows.Forms.MessageBox]::Show($res.Message, "Success")
                & $refreshAction
            }
        } else {
            # 横型: 通常の解像度変更
            $msg = $pr.Name + " " + $pr.W.ToString() + "x" + $pr.H.ToString() + " - Apply?"
            $confirm = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm", "YesNo")
            if ($confirm -ne "Yes") { return }
            $devName = $script:displayList[$idx].Name
            $res = Set-DisplayResolution $devName $pr.W $pr.H 0
            if ($res.Success) {
                $lblStatus.Text = "OK: " + $res.Message
                [System.Windows.Forms.MessageBox]::Show($res.Message, "Success")
                & $refreshAction
            } else {
                $lblStatus.Text = "NG: " + $res.Message
                [System.Windows.Forms.MessageBox]::Show($res.Message, "Error")
            }
        }
    })
    $grpPreset.Controls.Add($btn)
    $bx += 370
    if ($bx -gt 400) { $bx = 10; $by += 30 }
}
$y += 190

# EDID Override controls
$grpEdid = New-Object System.Windows.Forms.GroupBox
$grpEdid.Text = "EDID Override (Portrait Resolution Registration)"
$grpEdid.Location = New-Object System.Drawing.Point(15, $y)
$grpEdid.Size = New-Object System.Drawing.Size(750, 55)
$form.Controls.Add($grpEdid)

$btnRemoveEdid = New-Object System.Windows.Forms.Button
$btnRemoveEdid.Text = "Remove EDID Override (Restore Original)"
$btnRemoveEdid.Location = New-Object System.Drawing.Point(10, 20)
$btnRemoveEdid.Size = New-Object System.Drawing.Size(300, 26)
$btnRemoveEdid.Add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show("Remove EDID override and restore original settings?`nReboot or re-plug required after removal.", "Confirm", "YesNo")
    if ($confirm -ne "Yes") { return }
    $res = Remove-CustomResolution
    if ($res.Success) {
        $lblStatus.Text = "OK: " + $res.Message
        [System.Windows.Forms.MessageBox]::Show($res.Message, "Success")
    } else {
        $lblStatus.Text = "NG: " + $res.Message
        [System.Windows.Forms.MessageBox]::Show($res.Message, "Error")
    }
})
$grpEdid.Controls.Add($btnRemoveEdid)

$lblEdidNote = New-Object System.Windows.Forms.Label
$lblEdidNote.Text = "* Portrait presets automatically register EDID. Admin rights required."
$lblEdidNote.Location = New-Object System.Drawing.Point(320, 24)
$lblEdidNote.AutoSize = $true
$grpEdid.Controls.Add($lblEdidNote)
$y += 65

# Custom resolution
$grpCustom = New-Object System.Windows.Forms.GroupBox
$grpCustom.Text = "Custom Resolution"
$grpCustom.Location = New-Object System.Drawing.Point(15, $y)
$grpCustom.Size = New-Object System.Drawing.Size(750, 55)
$form.Controls.Add($grpCustom)

$lblW = New-Object System.Windows.Forms.Label
$lblW.Text = "W:"
$lblW.Location = New-Object System.Drawing.Point(10, 22)
$lblW.AutoSize = $true
$grpCustom.Controls.Add($lblW)

$txtW = New-Object System.Windows.Forms.TextBox
$txtW.Text = "2732"
$txtW.Location = New-Object System.Drawing.Point(30, 20)
$txtW.Size = New-Object System.Drawing.Size(70, 22)
$grpCustom.Controls.Add($txtW)

$lblH = New-Object System.Windows.Forms.Label
$lblH.Text = "H:"
$lblH.Location = New-Object System.Drawing.Point(110, 22)
$lblH.AutoSize = $true
$grpCustom.Controls.Add($lblH)

$txtH = New-Object System.Windows.Forms.TextBox
$txtH.Text = "2048"
$txtH.Location = New-Object System.Drawing.Point(130, 20)
$txtH.Size = New-Object System.Drawing.Size(70, 22)
$grpCustom.Controls.Add($txtH)

$btnApply = New-Object System.Windows.Forms.Button
$btnApply.Text = "Apply"
$btnApply.Location = New-Object System.Drawing.Point(220, 18)
$btnApply.Size = New-Object System.Drawing.Size(70, 26)
$grpCustom.Controls.Add($btnApply)

$btnFind = New-Object System.Windows.Forms.Button
$btnFind.Text = "Find Closest"
$btnFind.Location = New-Object System.Drawing.Point(300, 18)
$btnFind.Size = New-Object System.Drawing.Size(120, 26)
$grpCustom.Controls.Add($btnFind)
$y += 65

# Modes list
$lblModes = New-Object System.Windows.Forms.Label
$lblModes.Text = "Supported Resolutions (sorted by 4:3 proximity / double-click to apply)"
$lblModes.Location = New-Object System.Drawing.Point(15, $y)
$lblModes.AutoSize = $true
$form.Controls.Add($lblModes)
$y += 20

$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(15, $y)
$listView.Size = New-Object System.Drawing.Size(750, 240)
$listView.View = "Details"
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Columns.Add("Resolution", 160) | Out-Null
$listView.Columns.Add("Freq (Hz)", 100) | Out-Null
$listView.Columns.Add("Depth", 80) | Out-Null
$listView.Columns.Add("Aspect Ratio", 120) | Out-Null
$listView.Anchor = "Top,Left,Right,Bottom"
$form.Controls.Add($listView)
$y += 250

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready"
$lblStatus.Location = New-Object System.Drawing.Point(15, $y)
$lblStatus.Size = New-Object System.Drawing.Size(750, 20)
$lblStatus.BorderStyle = "Fixed3D"
$lblStatus.Anchor = "Bottom,Left,Right"
$form.Controls.Add($lblStatus)

# --- Logic ---
$script:displayList = @()

$refreshAction = {
    $script:displayList = @(Get-Displays)
    $cmbDisplay.Items.Clear()
    $selectIdx = 0
    for ($di = 0; $di -lt $script:displayList.Count; $di++) {
        $d = $script:displayList[$di]
        $cur = Get-CurrentResolution $d.Name
        $primaryTag = ""
        if ($d.Primary) { $primaryTag = " [Primary]" }
        $resStr = ""
        if ($cur) { $resStr = " (" + $cur.Width.ToString() + "x" + $cur.Height.ToString() + "@" + $cur.Freq.ToString() + "Hz)" }
        $label = $d.Name + $primaryTag + $resStr
        $cmbDisplay.Items.Add($label) | Out-Null
        if (-not $d.Primary) { $selectIdx = $di }
    }
    if ($cmbDisplay.Items.Count -gt 0) {
        $cmbDisplay.SelectedIndex = $selectIdx
    }
    $lblStatus.Text = $script:displayList.Count.ToString() + " display(s) found"
}

$refreshModes = {
    $listView.Items.Clear()
    $idx = $cmbDisplay.SelectedIndex
    if ($idx -lt 0) { return }
    $devName = $script:displayList[$idx].Name
    $cur = Get-CurrentResolution $devName
    if ($cur) {
        $ratio = Get-RatioStr $cur.Width $cur.Height
        $lblCurrent.Text = "Current: " + $cur.Width.ToString() + " x " + $cur.Height.ToString() + " @ " + $cur.Freq.ToString() + "Hz (Aspect: " + $ratio + ")"
    }
    $modes = @(Get-SupportedModes $devName)
    $targetRatio = 4.0 / 3.0
    $sorted = $modes | Sort-Object { [Math]::Abs($_.Width / $_.Height - $targetRatio) }, { -($_.Width * $_.Height) }
    foreach ($m in $sorted) {
        $ratio = Get-RatioStr $m.Width $m.Height
        $resLabel = $m.Width.ToString() + " x " + $m.Height.ToString()
        $item = New-Object System.Windows.Forms.ListViewItem($resLabel)
        $item.SubItems.Add($m.Freq.ToString()) | Out-Null
        $item.SubItems.Add($m.Bpp.ToString() + "bit") | Out-Null
        $item.SubItems.Add($ratio) | Out-Null
        $item.Tag = $m
        $ratioVal = $m.Width / $m.Height
        if ([Math]::Abs($ratioVal - $targetRatio) -lt 0.01) {
            $item.BackColor = [System.Drawing.Color]::FromArgb(212, 237, 218)
        }
        $listView.Items.Add($item) | Out-Null
    }
}

$btnRefresh.Add_Click({ & $refreshAction })
$cmbDisplay.Add_SelectedIndexChanged({ & $refreshModes })

$btnApply.Add_Click({
    $idx = $cmbDisplay.SelectedIndex
    if ($idx -lt 0) { [System.Windows.Forms.MessageBox]::Show("Select a display first", "Warning"); return }
    $w = 0; $h = 0
    if (-not [int]::TryParse($txtW.Text, [ref]$w) -or -not [int]::TryParse($txtH.Text, [ref]$h)) {
        [System.Windows.Forms.MessageBox]::Show("Enter valid numbers", "Error"); return
    }
    $msg = $w.ToString() + "x" + $h.ToString() + " - Apply?"
    $confirm = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm", "YesNo")
    if ($confirm -ne "Yes") { return }
    $devName = $script:displayList[$idx].Name
    $res = Set-DisplayResolution $devName $w $h 0
    if ($res.Success) {
        $lblStatus.Text = "OK: " + $res.Message
        [System.Windows.Forms.MessageBox]::Show($res.Message, "Success")
        & $refreshAction
    } else {
        $lblStatus.Text = "NG: " + $res.Message
        [System.Windows.Forms.MessageBox]::Show($res.Message, "Error")
    }
})

$btnFind.Add_Click({
    $idx = $cmbDisplay.SelectedIndex
    if ($idx -lt 0) { [System.Windows.Forms.MessageBox]::Show("Select a display first", "Warning"); return }
    $tw = 0; $th = 0
    if (-not [int]::TryParse($txtW.Text, [ref]$tw) -or -not [int]::TryParse($txtH.Text, [ref]$th)) {
        [System.Windows.Forms.MessageBox]::Show("Enter valid numbers", "Error"); return
    }
    $devName = $script:displayList[$idx].Name
    $modes = @(Get-SupportedModes $devName)
    $targetRatio = $tw / $th
    $best = $null; $bestScore = [double]::MaxValue
    foreach ($m in $modes) {
        $ratio = $m.Width / $m.Height
        $score = [Math]::Abs($ratio - $targetRatio) * 10000 + [Math]::Abs($m.Width - $tw) + [Math]::Abs($m.Height - $th)
        if ($score -lt $bestScore) { $bestScore = $score; $best = $m }
    }
    if ($best) {
        $ratio = Get-RatioStr $best.Width $best.Height
        $msg = "Closest: " + $best.Width.ToString() + " x " + $best.Height.ToString() + " @ " + $best.Freq.ToString() + "Hz  Aspect: " + $ratio + "  Apply?"
        $confirm = [System.Windows.Forms.MessageBox]::Show($msg, "Result", "YesNo")
        if ($confirm -eq "Yes") {
            $res = Set-DisplayResolution $devName $best.Width $best.Height $best.Freq
            if ($res.Success) { $lblStatus.Text = "OK: " + $res.Message; & $refreshAction }
            else { [System.Windows.Forms.MessageBox]::Show($res.Message, "Error") }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No supported resolution found", "Result")
    }
})

$listView.Add_DoubleClick({
    if ($listView.SelectedItems.Count -eq 0) { return }
    $selItem = $listView.SelectedItems[0]
    $m = $selItem.Tag
    $idx = $cmbDisplay.SelectedIndex
    if ($idx -lt 0) { return }
    $msg = $m.Width.ToString() + "x" + $m.Height.ToString() + " @ " + $m.Freq.ToString() + "Hz - Apply?"
    $confirm = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm", "YesNo")
    if ($confirm -ne "Yes") { return }
    $devName = $script:displayList[$idx].Name
    $res = Set-DisplayResolution $devName $m.Width $m.Height $m.Freq
    if ($res.Success) {
        $lblStatus.Text = "OK: " + $res.Message
        [System.Windows.Forms.MessageBox]::Show($res.Message, "Success")
        & $refreshAction
    } else {
        $lblStatus.Text = "NG: " + $res.Message
        [System.Windows.Forms.MessageBox]::Show($res.Message, "Error")
    }
})

& $refreshAction

$form.ShowDialog() | Out-Null
