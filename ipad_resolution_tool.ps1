Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not ([System.Management.Automation.PSTypeName]'DisplayAPI').Type) {
    try {
        Add-Type -Language CSharp -TypeDefinition @"
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
"@ -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to compile DisplayAPI type.`n`nError: $($_.Exception.Message)`n`nPossible causes:`n- PowerShell language mode is restricted`n- .NET compiler (csc.exe) not available`n`nTry running: powershell.exe -ExecutionPolicy Bypass -File `"$PSCommandPath`"",
            "Initialization Error", "OK", "Error")
        exit 1
    }
}

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

function Set-DisplayResolution($deviceName, $width, $height, $freq) {
    # 現在の設定をベースに取得（空のDEVMODEだとBADMODEになる）
    $dm = New-Object DisplayAPI+DEVMODE
    $dm.dmSize = [uint16][System.Runtime.InteropServices.Marshal]::SizeOf([type][DisplayAPI+DEVMODE])
    $gotCurrent = [DisplayAPI]::EnumDisplaySettingsW($deviceName, [DisplayAPI]::ENUM_CURRENT_SETTINGS, [ref]$dm)


    # 解像度を上書き
    $dm.dmPelsWidth = $width
    $dm.dmPelsHeight = $height
    $dm.dmFields = [DisplayAPI]::DM_PELSWIDTH -bor [DisplayAPI]::DM_PELSHEIGHT
    if ($freq -gt 0) {
        $dm.dmDisplayFrequency = $freq
        $dm.dmFields = $dm.dmFields -bor [DisplayAPI]::DM_DISPLAYFREQUENCY
    }

    # まずテスト
    $result = [DisplayAPI]::ChangeDisplaySettingsExW($deviceName, [ref]$dm, [IntPtr]::Zero, [DisplayAPI]::CDS_TEST, [IntPtr]::Zero)
    if ($result -ne 0) {
        # テスト失敗時: サポートモード一覧から完全一致のDEVMODEを取得して再試行
        $dm2 = New-Object DisplayAPI+DEVMODE
        $dm2.dmSize = [uint16][System.Runtime.InteropServices.Marshal]::SizeOf([type][DisplayAPI+DEVMODE])
        $modeIdx = 0
        while ([DisplayAPI]::EnumDisplaySettingsW($deviceName, $modeIdx, [ref]$dm2)) {
            if ($dm2.dmPelsWidth -eq $width -and $dm2.dmPelsHeight -eq $height) {
                $result = [DisplayAPI]::ChangeDisplaySettingsExW($deviceName, [ref]$dm2, [IntPtr]::Zero, [DisplayAPI]::CDS_TEST, [IntPtr]::Zero)
                if ($result -eq 0) {
                    $result = [DisplayAPI]::ChangeDisplaySettingsExW($deviceName, [ref]$dm2, [IntPtr]::Zero, [DisplayAPI]::CDS_UPDATEREGISTRY, [IntPtr]::Zero)
                    if ($result -eq 0) { return @{ Success = $true; Message = "Resolution changed! ($($dm2.dmPelsWidth)x$($dm2.dmPelsHeight))" } }
                    if ($result -eq 1) { return @{ Success = $true; Message = "Resolution changed. Restart required." } }
                    return @{ Success = $false; Message = "Change failed (code: $result)" }
                }
                break
            }
            $modeIdx++
        }
        return @{ Success = $false; Message = "Not supported (code: $result)" }
    }

    # テスト成功、実際に変更
    $result = [DisplayAPI]::ChangeDisplaySettingsExW($deviceName, [ref]$dm, [IntPtr]::Zero, [DisplayAPI]::CDS_UPDATEREGISTRY, [IntPtr]::Zero)
    if ($result -eq 0) { return @{ Success = $true; Message = "Resolution changed!" } }
    if ($result -eq 1) { return @{ Success = $true; Message = "Resolution changed. Restart required." } }
    return @{ Success = $false; Message = "Change failed (code: $result)" }
}

# --- GUI ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "iPad UVC Resolution Tool"
$form.Size = New-Object System.Drawing.Size(800, 720)
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
$grpPreset.Size = New-Object System.Drawing.Size(750, 150)
$form.Controls.Add($grpPreset)

$presets = @(
    @{ Name = "iPad Pro 13 (M4) Native"; W = 2752; H = 2064 },
    @{ Name = "iPad Pro 13 (M4) Half"; W = 1376; H = 1032 },
    @{ Name = "iPad Pro 12.9 / Air 13 Native"; W = 2732; H = 2048 },
    @{ Name = "iPad Pro 12.9 / Air 13 Half"; W = 1366; H = 1024 },
    @{ Name = "iPad 4:3 (2048x1536)"; W = 2048; H = 1536 },
    @{ Name = "iPad 4:3 Half (1024x768)"; W = 1024; H = 768 },
    @{ Name = "XGA 4:3 (1280x960)"; W = 1280; H = 960 },
    @{ Name = "SXGA 4:3 (1400x1050)"; W = 1400; H = 1050 }
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
            # 失敗時: 最も近いサポート解像度を提案
            $modes = @(Get-SupportedModes $devName)
            $targetRatio = $pr.W / $pr.H
            $best = $null; $bestScore = [double]::MaxValue
            foreach ($m in $modes) {
                $ratio = $m.Width / $m.Height
                $score = [Math]::Abs($ratio - $targetRatio) * 10000 + [Math]::Abs($m.Width - $pr.W) + [Math]::Abs($m.Height - $pr.H)
                if ($score -lt $bestScore) { $bestScore = $score; $best = $m }
            }
            if ($best) {
                $ratio = Get-RatioStr $best.Width $best.Height
                $suggestMsg = $pr.W.ToString() + "x" + $pr.H.ToString() + " is not supported by this capture card.`n`n" +
                    "Closest available resolution:`n" +
                    $best.Width.ToString() + " x " + $best.Height.ToString() + " @ " + $best.Freq.ToString() + "Hz (Aspect: " + $ratio + ")`n`n" +
                    "Apply this resolution instead?"
                $confirm2 = [System.Windows.Forms.MessageBox]::Show($suggestMsg, "Suggest Closest Resolution", "YesNo", "Question")
                if ($confirm2 -eq "Yes") {
                    $res2 = Set-DisplayResolution $devName $best.Width $best.Height $best.Freq
                    if ($res2.Success) {
                        $lblStatus.Text = "OK: " + $res2.Message
                        [System.Windows.Forms.MessageBox]::Show($res2.Message, "Success")
                        & $refreshAction
                    } else {
                        $lblStatus.Text = "NG: " + $res2.Message
                        [System.Windows.Forms.MessageBox]::Show($res2.Message, "Error")
                    }
                }
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
$y += 160

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
$lblModes.Text = "Supported Resolutions (double-click to apply)"
$lblModes.Location = New-Object System.Drawing.Point(15, $y)
$lblModes.AutoSize = $true
$form.Controls.Add($lblModes)

$lblSort = New-Object System.Windows.Forms.Label
$lblSort.Text = "Sort:"
$lblSort.Location = New-Object System.Drawing.Point(450, $y)
$lblSort.AutoSize = $true
$form.Controls.Add($lblSort)

$cmbSort = New-Object System.Windows.Forms.ComboBox
$cmbSort.Location = New-Object System.Drawing.Point(490, $y - 3)
$cmbSort.Size = New-Object System.Drawing.Size(220, 25)
$cmbSort.DropDownStyle = "DropDownList"
$cmbSort.Items.Add("Resolution (High to Low)") | Out-Null
$cmbSort.Items.Add("4:3 Proximity") | Out-Null
$cmbSort.Items.Add("Frequency (High to Low)") | Out-Null
$cmbSort.SelectedIndex = 0
$form.Controls.Add($cmbSort)
$y += 24

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
    $sortMode = $cmbSort.SelectedIndex
    switch ($sortMode) {
        0 { $sorted = $modes | Sort-Object { $_.Width * $_.Height } -Descending }
        1 { $sorted = $modes | Sort-Object { [Math]::Abs($_.Width / $_.Height - $targetRatio) }, { -($_.Width * $_.Height) } }
        2 { $sorted = $modes | Sort-Object { $_.Freq } -Descending }
        default { $sorted = $modes | Sort-Object { $_.Width * $_.Height } -Descending }
    }
    # 同一解像度で最高周波数のみ表示（ソートモード2以外）
    $seen = @{}
    foreach ($m in $sorted) {
        $ratio = Get-RatioStr $m.Width $m.Height
        $resKey = $m.Width.ToString() + "x" + $m.Height.ToString()
        if ($sortMode -ne 2 -and $seen.ContainsKey($resKey)) { continue }
        $seen[$resKey] = $true
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
$cmbSort.Add_SelectedIndexChanged({ & $refreshModes })

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
