# Suport natiu: DPI Per-Monitor V2 (IntPtr) + DestroyIcon per evitar fuites GDI
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class NativeHelper {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool DestroyIcon(IntPtr handle);
}
"@
try { [NativeHelper]::SetProcessDpiAwarenessContext([IntPtr](-4)) | Out-Null } catch {}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Control d'instància única per evitar icones i avisos duplicats
$createdNew = $false
$mutex = New-Object System.Threading.Mutex($true, "Local\AlertaOutlookTray_SingleInstance", [ref]$createdNew)
if (-not $createdNew) {
    [System.Windows.Forms.MessageBox]::Show(
        "AlertaOutlookTray ja s'està executant a la safata del sistema.",
        "AlertaOutlookTray",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
    exit
}

$configDir  = Join-Path $env:APPDATA "AlertaOutlook"
$configPath = Join-Path $configDir "config.json"

if (-not (Test-Path $configDir)) {
    New-Item -Path $configDir -ItemType Directory -Force | Out-Null
}

function Get-FreshDefaultConfig {
    return [ordered]@{
        Theme           = "System"    # "System", "Dark" o "Light"
        AccentHex       = "#D93025"   # Vermell per defecte
        IconStyle       = "Envelope"  # "Circle", "Envelope" o "Custom"
        CustomIdleIcon  = ""
        CustomAlertIcon = ""
        IntervalSec     = 10
        PlaySound       = $true
        Position        = "Center"    # "Center" o "BottomRight"
        AllMonitors     = $false      # Desactivat per defecte
    }
}

function Load-Config {
    $cfg = Get-FreshDefaultConfig
    if (Test-Path $configPath) {
        try {
            $loaded = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($key in @($cfg.Keys)) {
                if ($null -ne $loaded.$key) { $cfg[$key] = $loaded.$key }
            }
        } catch {}
    }
    return $cfg
}

function Save-Config {
    param($cfg)
    if (-not (Test-Path $configDir)) {
        New-Item -Path $configDir -ItemType Directory -Force | Out-Null
    }
    $cfg | ConvertTo-Json | Set-Content -Path $configPath -Encoding UTF8
}

function Reset-ConfigInPlace {
    $fresh = Get-FreshDefaultConfig
    foreach ($key in @($fresh.Keys)) {
        $config[$key] = $fresh[$key]
    }
}

function Get-EffectiveTheme {
    if ($config.Theme -eq "Light" -or $config.Theme -eq "Dark") {
        return $config.Theme
    }
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $val = Get-ItemPropertyValue -Path $regPath -Name "AppsUseLightTheme" -ErrorAction Stop
        if ($val -eq 1) { return "Light" }
    } catch {}
    return "Dark"
}

$config = Load-Config

# Generador d'icones vectorials amb alliberament net de handles HICON
function New-TrayIcon {
    param (
        [System.Drawing.Color]$BgColor,
        [string]$Style,
        [bool]$IsAlert,
        [string]$CustomPath
    )

    if ($Style -eq "Custom" -and $CustomPath -and (Test-Path $CustomPath)) {
        try {
            return New-Object System.Drawing.Icon($CustomPath, 32, 32)
        } catch {}
    }

    $bmp = New-Object System.Drawing.Bitmap(32, 32)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

    $brushBg   = New-Object System.Drawing.SolidBrush($BgColor)
    $penWhite  = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 2)
    $brushText = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)

    try {
        if ($Style -eq "Envelope") {
            $g.FillRectangle($brushBg, 2, 5, 28, 22)
            $g.DrawRectangle($penWhite, 2, 5, 28, 22)
            $g.DrawLine($penWhite, 3, 6, 16, 17)
            $g.DrawLine($penWhite, 29, 6, 16, 17)
            if ($IsAlert) {
                $badgeBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
                $dotBrush   = New-Object System.Drawing.SolidBrush($BgColor)
                try {
                    $g.FillEllipse($badgeBrush, 19, 1, 12, 12)
                    $g.FillEllipse($dotBrush, 22, 4, 6, 6)
                } finally {
                    $badgeBrush.Dispose()
                    $dotBrush.Dispose()
                }
            }
        } else {
            $g.FillEllipse($brushBg, 2, 2, 28, 28)
            $g.DrawEllipse($penWhite, 2, 2, 28, 28)
            $txt  = if ($IsAlert) { "!" } else { "M" }
            $font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
            $sf   = New-Object System.Drawing.StringFormat
            try {
                $sf.Alignment     = [System.Drawing.StringAlignment]::Center
                $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
                $g.DrawString($txt, $font, $brushText, (New-Object System.Drawing.RectangleF(0, 1, 32, 32)), $sf)
            } finally {
                $font.Dispose()
                $sf.Dispose()
            }
        }

        $hIcon = $bmp.GetHicon()
        try {
            $tmpIcon = [System.Drawing.Icon]::FromHandle($hIcon)
            return $tmpIcon.Clone()
        } finally {
            [NativeHelper]::DestroyIcon($hIcon) | Out-Null
        }
    } finally {
        $brushBg.Dispose()
        $penWhite.Dispose()
        $brushText.Dispose()
        $g.Dispose()
        $bmp.Dispose()
    }
}

# Estat compartit
$state = @{
    NotifiedIds     = New-Object System.Collections.Generic.HashSet[string]
    Outlook         = $null
    Inbox           = $null
    IsPaused        = $false
    IsAlertOpen     = $false
    IsClosingAlerts = $false
    AlertCount      = 0
    AlertWindows    = New-Object System.Collections.ArrayList
    IconIdle        = $null
    IconAlert       = $null
    IconPaused      = $null
}

$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Visible = $true

function Close-AllAlerts {
    if ($state.IsClosingAlerts) { return }
    $state.IsClosingAlerts = $true
    try {
        foreach ($w in @($state.AlertWindows)) {
            if ($w.Form -and -not $w.Form.IsDisposed) {
                $w.Form.Close()
            }
        }
        $state.AlertWindows.Clear()
        $state.IsAlertOpen = $false
        $state.AlertCount  = 0
        Rebuild-IconsAndVisuals
    } finally {
        $state.IsClosingAlerts = $false
    }
}

function Apply-FormPosition {
    param($w)
    if (-not $w.Form -or $w.Form.IsDisposed) { return }
    $wa = $w.Screen.WorkingArea
    $w.Form.StartPosition = "Manual"
    if ($config.Position -eq "BottomRight") {
        $w.Form.Location = New-Object System.Drawing.Point(($wa.Right - $w.Form.Width - 16), ($wa.Bottom - $w.Form.Height - 16))
    } else {
        $w.Form.Location = New-Object System.Drawing.Point(
            ($wa.Left + [int](($wa.Width - $w.Form.Width) / 2)),
            ($wa.Top  + [int](($wa.Height - $w.Form.Height) / 2))
        )
    }
}

function Apply-FormTheme {
    param($w)
    if (-not $w.Form -or $w.Form.IsDisposed) { return }
    $accent         = [System.Drawing.ColorTranslator]::FromHtml($config.AccentHex)
    $effectiveTheme = Get-EffectiveTheme

    if ($effectiveTheme -eq "Light") {
        $w.Form.BackColor       = [System.Drawing.Color]::FromArgb(248, 249, 250)
        $w.LblSender.ForeColor  = [System.Drawing.Color]::FromArgb(25, 25, 28)
        $w.LblSubject.ForeColor = [System.Drawing.Color]::FromArgb(65, 65, 70)
        $w.LblHeader.ForeColor  = $accent
    } else {
        $w.Form.BackColor       = [System.Drawing.Color]::FromArgb(28, 28, 30)
        $w.LblSender.ForeColor  = [System.Drawing.Color]::White
        $w.LblSubject.ForeColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
        $w.LblHeader.ForeColor  = [System.Drawing.Color]::FromArgb(
            [Math]::Min(255, $accent.R + 45),
            [Math]::Min(255, $accent.G + 45),
            [Math]::Min(255, $accent.B + 45)
        )
    }
    $w.Banner.BackColor = $accent
    $w.BtnOk.BackColor  = $accent
}

function Rebuild-IconsAndVisuals {
    $accentColor = [System.Drawing.ColorTranslator]::FromHtml($config.AccentHex)
    $idleColor   = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $pausedColor = [System.Drawing.Color]::FromArgb(110, 110, 115)

    $oldIdle   = $state.IconIdle
    $oldAlert  = $state.IconAlert
    $oldPaused = $state.IconPaused

    $state.IconIdle   = New-TrayIcon -BgColor $idleColor   -Style $config.IconStyle -IsAlert $false -CustomPath $config.CustomIdleIcon
    $state.IconAlert  = New-TrayIcon -BgColor $accentColor -Style $config.IconStyle -IsAlert $true  -CustomPath $config.CustomAlertIcon
    $state.IconPaused = New-TrayIcon -BgColor $pausedColor -Style $config.IconStyle -IsAlert $false -CustomPath ""

    if ($state.IsPaused) {
        $notifyIcon.Icon = $state.IconPaused
        $notifyIcon.Text = "Monitor Outlook (PAUSAT)"
    } elseif ($state.IsAlertOpen) {
        $notifyIcon.Icon = $state.IconAlert
    } else {
        $notifyIcon.Icon = $state.IconIdle
        $notifyIcon.Text = "Monitor Outlook actiu"
    }

    foreach ($w in @($state.AlertWindows)) {
        if ($w.Form -and -not $w.Form.IsDisposed) {
            $w.Form.Icon = $state.IconAlert
            Apply-FormTheme -w $w
            Apply-FormPosition -w $w
        }
    }

    if ($oldIdle)   { $oldIdle.Dispose() }
    if ($oldAlert)  { $oldAlert.Dispose() }
    if ($oldPaused) { $oldPaused.Dispose() }
}

# --- MENÚ CONTEXTUAL ---
$menu = New-Object System.Windows.Forms.ContextMenuStrip

$miTest  = New-Object System.Windows.Forms.ToolStripMenuItem("Provar avís ara")
$miPause = New-Object System.Windows.Forms.ToolStripMenuItem("Pausar avisos")
$miPause.CheckOnClick = $true

$miTheme       = New-Object System.Windows.Forms.ToolStripMenuItem("Tema dels avisos")
$miThemeSystem = New-Object System.Windows.Forms.ToolStripMenuItem("Automàtic (Sistema)")
$miThemeDark   = New-Object System.Windows.Forms.ToolStripMenuItem("Fosc")
$miThemeLight  = New-Object System.Windows.Forms.ToolStripMenuItem("Clar")
$miTheme.DropDownItems.AddRange(@($miThemeSystem, (New-Object System.Windows.Forms.ToolStripSeparator), $miThemeDark, $miThemeLight))

$miColor = New-Object System.Windows.Forms.ToolStripMenuItem("Color de destacat")
$presets = @(
    @{ Name = "Vermell alerta"; Hex = "#D93025" },
    @{ Name = "Taronja ambre";  Hex = "#E67C00" },
    @{ Name = "Blau Outlook";   Hex = "#0078D4" },
    @{ Name = "Verd maragda";   Hex = "#107C41" },
    @{ Name = "Lila intens";    Hex = "#7719AA" }
)
foreach ($p in $presets) {
    $item = New-Object System.Windows.Forms.ToolStripMenuItem($p.Name)
    $item.Tag = $p.Hex
    $item.Add_Click({
        param($s, $e)
        $config.AccentHex = $s.Tag
        Save-Config $config
        Rebuild-IconsAndVisuals
        Sync-MenuChecks
    })
    $miColor.DropDownItems.Add($item) | Out-Null
}
$miColor.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
$miColorCustom = New-Object System.Windows.Forms.ToolStripMenuItem("Tria color personalitzat...")
$miColor.DropDownItems.Add($miColorCustom) | Out-Null

$miIcon         = New-Object System.Windows.Forms.ToolStripMenuItem("Estil de la icona")
$miIconEnvelope = New-Object System.Windows.Forms.ToolStripMenuItem("Sobre de correu")
$miIconCircle   = New-Object System.Windows.Forms.ToolStripMenuItem("Cercle (M / !)")
$miIconCustom   = New-Object System.Windows.Forms.ToolStripMenuItem("Carregar fitxers .ico...")
$miIcon.DropDownItems.AddRange(@($miIconEnvelope, $miIconCircle, (New-Object System.Windows.Forms.ToolStripSeparator), $miIconCustom))

$miOptions     = New-Object System.Windows.Forms.ToolStripMenuItem("Opcions d'avís")
$miSound       = New-Object System.Windows.Forms.ToolStripMenuItem("Reproduir so en arribar")
$miSound.CheckOnClick = $true
$miAllMonitors = New-Object System.Windows.Forms.ToolStripMenuItem("Mostrar a tots els monitors")
$miAllMonitors.CheckOnClick = $true
$miPosCenter   = New-Object System.Windows.Forms.ToolStripMenuItem("Posició: Centre pantalla")
$miPosCorner   = New-Object System.Windows.Forms.ToolStripMenuItem("Posició: Cantonada inferior dreta")
$miOptions.DropDownItems.AddRange(@(
    $miSound,
    $miAllMonitors,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $miPosCenter,
    $miPosCorner
))

$miConfigGroup = New-Object System.Windows.Forms.ToolStripMenuItem("Configuració")
$miOpenConfig  = New-Object System.Windows.Forms.ToolStripMenuItem("Obrir config.json")
$miResetConfig = New-Object System.Windows.Forms.ToolStripMenuItem("Restablir valors per defecte")
$miConfigGroup.DropDownItems.AddRange(@($miOpenConfig, $miResetConfig))

$miExit = New-Object System.Windows.Forms.ToolStripMenuItem("Sortir")

$menu.Items.AddRange(@(
    $miTest,
    $miPause,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $miTheme,
    $miColor,
    $miIcon,
    $miOptions,
    $miConfigGroup,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $miExit
))
$notifyIcon.ContextMenuStrip = $menu

function Sync-MenuChecks {
    $miThemeSystem.Checked   = ($config.Theme -eq "System")
    $miThemeDark.Checked     = ($config.Theme -eq "Dark")
    $miThemeLight.Checked    = ($config.Theme -eq "Light")
    $miIconEnvelope.Checked  = ($config.IconStyle -eq "Envelope")
    $miIconCircle.Checked    = ($config.IconStyle -eq "Circle")
    $miIconCustom.Checked    = ($config.IconStyle -eq "Custom")
    $miSound.Checked         = [bool]$config.PlaySound
    $miAllMonitors.Checked   = [bool]$config.AllMonitors
    $miPosCenter.Checked     = ($config.Position -eq "Center")
    $miPosCorner.Checked     = ($config.Position -eq "BottomRight")

    foreach ($sub in $miColor.DropDownItems) {
        if ($sub -is [System.Windows.Forms.ToolStripMenuItem] -and $sub.Tag) {
            $sub.Checked = ($sub.Tag -eq $config.AccentHex)
        }
    }
}

Rebuild-IconsAndVisuals
Sync-MenuChecks

$miPause.Add_Click({
    $state.IsPaused = $miPause.Checked
    Rebuild-IconsAndVisuals
})

$miThemeSystem.Add_Click({ $config.Theme = "System"; Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks })
$miThemeDark.Add_Click({   $config.Theme = "Dark";   Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks })
$miThemeLight.Add_Click({  $config.Theme = "Light";  Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks })

$miColorCustom.Add_Click({
    $dlg = New-Object System.Windows.Forms.ColorDialog
    try {
        $dlg.Color    = [System.Drawing.ColorTranslator]::FromHtml($config.AccentHex)
        $dlg.FullOpen = $true
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $config.AccentHex = "#{0:X2}{1:X2}{2:X2}" -f $dlg.Color.R, $dlg.Color.G, $dlg.Color.B
            Save-Config $config
            Rebuild-IconsAndVisuals
            Sync-MenuChecks
        }
    } finally {
        $dlg.Dispose()
    }
})

$miIconEnvelope.Add_Click({ $config.IconStyle = "Envelope"; Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks })
$miIconCircle.Add_Click({   $config.IconStyle = "Circle";   Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks })
$miIconCustom.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    try {
        $ofd.Filter = "Fitxers d'icona (*.ico)|*.ico"
        $ofd.Title  = "1/2 - Selecciona la icona en REPÒS (normal)"
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $config.CustomIdleIcon = $ofd.FileName
            $ofd.Title = "2/2 - Selecciona la icona d'ALERTA (correu nou)"
            if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $config.CustomAlertIcon = $ofd.FileName
                $config.IconStyle = "Custom"
                Save-Config $config
                Rebuild-IconsAndVisuals
                Sync-MenuChecks
            }
        }
    } finally {
        $ofd.Dispose()
    }
})

$miSound.Add_Click({       $config.PlaySound   = $miSound.Checked;       Save-Config $config })
$miAllMonitors.Add_Click({ $config.AllMonitors = $miAllMonitors.Checked; Save-Config $config })
$miPosCenter.Add_Click({   $config.Position    = "Center";      Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks })
$miPosCorner.Add_Click({   $config.Position    = "BottomRight"; Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks })

$miOpenConfig.Add_Click({
    if (-not (Test-Path $configPath)) { Save-Config $config }
    Start-Process -FilePath "notepad.exe" -ArgumentList "`"$configPath`""
})

$miResetConfig.Add_Click({
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "Vols restablir tota la configuració als valors per defecte?`n`nFitxer: $configPath",
        "Restablir configuració",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
        Reset-ConfigInPlace
        Save-Config $config
        $timer.Interval = [int]$config.IntervalSec * 1000
        Rebuild-IconsAndVisuals
        Sync-MenuChecks
    }
})

# --- LÒGICA D'OUTLOOK I FINESTRES ---
function Initialize-OutlookFolder {
    try {
        if (-not $state.Outlook) {
            $state.Outlook = New-Object -ComObject Outlook.Application
        }
        $ns = $state.Outlook.GetNamespace("MAPI")
        $state.Inbox = $ns.GetDefaultFolder(6)
        return $true
    } catch {
        $state.Outlook = $null
        $state.Inbox   = $null
        return $false
    }
}

if (Initialize-OutlookFolder) {
    try {
        foreach ($item in $state.Inbox.Items.Restrict("[Unread] = True")) {
            if ($item.EntryID) { [void]$state.NotifiedIds.Add($item.EntryID) }
        }
    } catch {}
}

function Update-AllAlertWindowsText {
    param ([string]$Sender, [string]$Subject, [string]$Time)

    $headerTxt = if ($state.AlertCount -le 1) {
        "NOU CORREU REBUT ($Time)"
    } else {
        "$($state.AlertCount) CORREUS NOUS PENDENTS - ÚLTIM ($Time)"
    }

    foreach ($w in @($state.AlertWindows)) {
        if ($w.Form -and -not $w.Form.IsDisposed) {
            $w.LblHeader.Text  = $headerTxt
            $w.LblSender.Text  = "De: $Sender"
            $w.LblSubject.Text = "Assumpte: $Subject"
            $w.Form.TopMost    = $true
        }
    }
    $notifyIcon.Text = "Outlook: $($state.AlertCount) correu(s) nou(s) sense acceptar"
}

function New-AlertWindowForScreen {
    param ([System.Windows.Forms.Screen]$Screen)

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = "Avís de Nou Correu"
    $form.Icon            = $state.IconAlert
    $form.Size            = New-Object System.Drawing.Size(480, 245)
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false
    $form.TopMost         = $true

    $banner = New-Object System.Windows.Forms.Panel
    $banner.Dock   = "Top"
    $banner.Height = 8
    $form.Controls.Add($banner)

    $lblHeader = New-Object System.Windows.Forms.Label
    $lblHeader.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblHeader.BackColor = [System.Drawing.Color]::Transparent
    $lblHeader.Location  = New-Object System.Drawing.Point(22, 22)
    $lblHeader.Size      = New-Object System.Drawing.Size(420, 22)
    $form.Controls.Add($lblHeader)

    $lblSender = New-Object System.Windows.Forms.Label
    $lblSender.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblSender.BackColor = [System.Drawing.Color]::Transparent
    $lblSender.Location  = New-Object System.Drawing.Point(22, 52)
    $lblSender.Size      = New-Object System.Drawing.Size(420, 26)
    $form.Controls.Add($lblSender)

    $lblSubject = New-Object System.Windows.Forms.Label
    $lblSubject.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblSubject.BackColor = [System.Drawing.Color]::Transparent
    $lblSubject.Location  = New-Object System.Drawing.Point(22, 86)
    $lblSubject.Size      = New-Object System.Drawing.Size(420, 56)
    $form.Controls.Add($lblSubject)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text      = "Acceptar"
    $btnOk.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnOk.Size      = New-Object System.Drawing.Size(135, 38)
    $btnOk.Location  = New-Object System.Drawing.Point(307, 152)
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.FlatStyle = "Flat"
    $btnOk.FlatAppearance.BorderSize = 0
    $btnOk.Add_Click({ Close-AllAlerts })
    $form.Controls.Add($btnOk)

    $form.AcceptButton = $btnOk
    $form.Add_FormClosed({ Close-AllAlerts })

    $w = @{
        Form       = $form
        Banner     = $banner
        LblHeader  = $lblHeader
        LblSender  = $lblSender
        LblSubject = $lblSubject
        BtnOk      = $btnOk
        Screen     = $Screen
    }

    Apply-FormTheme -w $w
    Apply-FormPosition -w $w
    return $w
}

function Show-Or-UpdateAlert {
    param ([string]$Sender, [string]$Subject, [string]$Time)

    $state.AlertCount++
    $notifyIcon.Icon = $state.IconAlert
    if ($config.PlaySound) { [System.Media.SystemSounds]::Exclamation.Play() }

    if ($state.IsAlertOpen -and $state.AlertWindows.Count -gt 0) {
        foreach ($w in @($state.AlertWindows)) {
            Apply-FormTheme -w $w
        }
        Update-AllAlertWindowsText -Sender $Sender -Subject $Subject -Time $Time
        return
    }

    $screens = if ($config.AllMonitors) {
        [System.Windows.Forms.Screen]::AllScreens
    } else {
        @([System.Windows.Forms.Screen]::PrimaryScreen)
    }

    $state.AlertWindows.Clear()
    foreach ($scr in $screens) {
        $w = New-AlertWindowForScreen -Screen $scr
        [void]$state.AlertWindows.Add($w)
    }

    $state.IsAlertOpen = $true
    Update-AllAlertWindowsText -Sender $Sender -Subject $Subject -Time $Time

    foreach ($w in @($state.AlertWindows)) {
        $w.Form.Show()
    }
}

$miTest.Add_Click({
    Show-Or-UpdateAlert -Sender "Prova de Sistema" -Subject "Aquesta és una vista prèvia de l'avís amb el tema i color actuals." -Time (Get-Date -Format "HH:mm")
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [int]$config.IntervalSec * 1000
$timer.Add_Tick({
    if ($state.IsPaused) { return }
    if (-not $state.Inbox -and -not (Initialize-OutlookFolder)) { return }

    try {
        $unreadItems = $state.Inbox.Items.Restrict("[Unread] = True")
        foreach ($mail in $unreadItems) {
            if ($mail.EntryID -and -not $state.NotifiedIds.Contains($mail.EntryID)) {
                [void]$state.NotifiedIds.Add($mail.EntryID)
                $hora = $mail.ReceivedTime.ToString("HH:mm")
                Show-Or-UpdateAlert -Sender $mail.SenderName -Subject $mail.Subject -Time $hora
            }
        }
    } catch {
        $state.Outlook = $null
        $state.Inbox   = $null
    }
})

$miExit.Add_Click({
    $timer.Stop()
    Close-AllAlerts
    $notifyIcon.Visible = $false
    $notifyIcon.Dispose()
    [System.Windows.Forms.Application]::Exit()
})

try {
    $timer.Start()
    [System.Windows.Forms.Application]::Run()
} finally {
    if ($state.IconIdle)   { $state.IconIdle.Dispose() }
    if ($state.IconAlert)  { $state.IconAlert.Dispose() }
    if ($state.IconPaused) { $state.IconPaused.Dispose() }
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
