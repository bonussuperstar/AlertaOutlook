# Consciència DPI Per-Monitor V2 per a pantalles d'alta densitat
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DpiHelper {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDpiAwarenessContext(int dpiFlag);
}
"@
[DpiHelper]::SetProcessDpiAwarenessContext(-4) | Out-Null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$configDir  = Join-Path $env:APPDATA "AlertaOutlook"
$configPath = Join-Path $configDir "config.json"

if (-not (Test-Path $configDir)) {
    New-Item -Path $configDir -ItemType Directory -Force | Out-Null
}

# Configuració per defecte
$defaultConfig = [ordered]@{
    Theme           = "System"    # "System", "Dark" o "Light"
    AccentHex       = "#D93025"   # Vermell per defecte
    IconStyle       = "Envelope"  # "Circle", "Envelope" o "Custom"
    CustomIdleIcon  = ""
    CustomAlertIcon = ""
    IntervalSec     = 10
    PlaySound       = $true
    Position        = "Center"    # "Center" o "BottomRight"
}

function Get-FreshDefaultConfig {
    return [ordered]@{
        Theme           = $defaultConfig.Theme
        AccentHex       = $defaultConfig.AccentHex
        IconStyle       = $defaultConfig.IconStyle
        CustomIdleIcon  = $defaultConfig.CustomIdleIcon
        CustomAlertIcon = $defaultConfig.CustomAlertIcon
        IntervalSec     = $defaultConfig.IntervalSec
        PlaySound       = $defaultConfig.PlaySound
        Position        = $defaultConfig.Position
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

# Resol el tema efectiu ("Dark" o "Light") consultant el registre de Windows si està en "System"
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

# Generador d'icones vectorials en memòria
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
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

    $brushBg   = New-Object System.Drawing.SolidBrush($BgColor)
    $penWhite  = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 2)
    $brushText = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)

    if ($Style -eq "Envelope") {
        $g.FillRectangle($brushBg, 2, 5, 28, 22)
        $g.DrawRectangle($penWhite, 2, 5, 28, 22)
        $g.DrawLine($penWhite, 3, 6, 16, 17)
        $g.DrawLine($penWhite, 29, 6, 16, 17)
        if ($IsAlert) {
            $badgeBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
            $g.FillEllipse($badgeBrush, 19, 1, 12, 12)
            $dotBrush = New-Object System.Drawing.SolidBrush($BgColor)
            $g.FillEllipse($dotBrush, 22, 4, 6, 6)
        }
    } else {
        $g.FillEllipse($brushBg, 2, 2, 28, 28)
        $g.DrawEllipse($penWhite, 2, 2, 28, 28)
        $txt  = if ($IsAlert) { "!" } else { "M" }
        $font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        $sf   = New-Object System.Drawing.StringFormat
        $sf.Alignment     = [System.Drawing.StringAlignment]::Center
        $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        $g.DrawString($txt, $font, $brushText, (New-Object System.Drawing.RectangleF(0, 1, 32, 32)), $sf)
    }

    $hIcon = $bmp.GetHicon()
    $icon  = [System.Drawing.Icon]::FromHandle($hIcon)
    $g.Dispose()
    $bmp.Dispose()
    return $icon
}

# Estat compartit
$state = @{
    NotifiedIds = New-Object System.Collections.Generic.HashSet[string]
    Outlook     = $null
    Inbox       = $null
    IsPaused    = $false
    IsAlertOpen = $false
    AlertCount  = 0
    Form        = $null
    Banner      = $null
    LblHeader   = $null
    LblSender   = $null
    LblSubject  = $null
    BtnOk       = $null
    IconIdle    = $null
    IconAlert   = $null
    IconPaused  = $null
}

$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Visible = $true

function Rebuild-IconsAndVisuals {
    $accentColor = [System.Drawing.ColorTranslator]::FromHtml($config.AccentHex)
    $idleColor   = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $pausedColor = [System.Drawing.Color]::FromArgb(110, 110, 115)

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

    if ($state.IsAlertOpen -and $state.Form -and -not $state.Form.IsDisposed) {
        Apply-FormTheme
        Apply-FormPosition
    }
}

function Apply-FormPosition {
    if (-not $state.Form -or $state.Form.IsDisposed) { return }
    if ($config.Position -eq "BottomRight") {
        $state.Form.StartPosition = "Manual"
        $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $state.Form.Location = New-Object System.Drawing.Point(($wa.Right - $state.Form.Width - 16), ($wa.Bottom - $state.Form.Height - 16))
    } else {
        $state.Form.StartPosition = "CenterScreen"
        $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $state.Form.Location = New-Object System.Drawing.Point(
            ($wa.Left + [int](($wa.Width - $state.Form.Width) / 2)),
            ($wa.Top  + [int](($wa.Height - $state.Form.Height) / 2))
        )
    }
}

function Apply-FormTheme {
    if (-not $state.Form -or $state.Form.IsDisposed) { return }
    $accent         = [System.Drawing.ColorTranslator]::FromHtml($config.AccentHex)
    $effectiveTheme = Get-EffectiveTheme

    if ($effectiveTheme -eq "Light") {
        $state.Form.BackColor     = [System.Drawing.Color]::FromArgb(248, 249, 250)
        $state.LblSender.ForeColor  = [System.Drawing.Color]::FromArgb(25, 25, 28)
        $state.LblSubject.ForeColor = [System.Drawing.Color]::FromArgb(65, 65, 70)
        $state.LblHeader.ForeColor  = $accent
    } else {
        $state.Form.BackColor     = [System.Drawing.Color]::FromArgb(28, 28, 30)
        $state.LblSender.ForeColor  = [System.Drawing.Color]::White
        $state.LblSubject.ForeColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
        $state.LblHeader.ForeColor  = [System.Drawing.Color]::FromArgb(
            [Math]::Min(255, $accent.R + 45),
            [Math]::Min(255, $accent.G + 45),
            [Math]::Min(255, $accent.B + 45)
        )
    }
    $state.Banner.BackColor = $accent
    $state.BtnOk.BackColor  = $accent
}

# --- CONSTRUCCIÓ DEL MENÚ CONTEXTUAL ---
$menu = New-Object System.Windows.Forms.ContextMenuStrip

$miTest  = New-Object System.Windows.Forms.ToolStripMenuItem("Provar avís ara")
$miPause = New-Object System.Windows.Forms.ToolStripMenuItem("Pausar avisos")
$miPause.CheckOnClick = $true

# Submenú Tema
$miTheme       = New-Object System.Windows.Forms.ToolStripMenuItem("Tema dels avisos")
$miThemeSystem = New-Object System.Windows.Forms.ToolStripMenuItem("Automàtic (Sistema)")
$miThemeDark   = New-Object System.Windows.Forms.ToolStripMenuItem("Fosc")
$miThemeLight  = New-Object System.Windows.Forms.ToolStripMenuItem("Clar")
$miTheme.DropDownItems.AddRange(@($miThemeSystem, (New-Object System.Windows.Forms.ToolStripSeparator), $miThemeDark, $miThemeLight))

# Submenú Color de destacat
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

# Submenú Icona
$miIcon         = New-Object System.Windows.Forms.ToolStripMenuItem("Estil de la icona")
$miIconEnvelope = New-Object System.Windows.Forms.ToolStripMenuItem("Sobre de correu")
$miIconCircle   = New-Object System.Windows.Forms.ToolStripMenuItem("Cercle (M / !)")
$miIconCustom   = New-Object System.Windows.Forms.ToolStripMenuItem("Carregar fitxers .ico...")
$miIcon.DropDownItems.AddRange(@($miIconEnvelope, $miIconCircle, (New-Object System.Windows.Forms.ToolStripSeparator), $miIconCustom))

# Submenú Posició i So
$miOptions     = New-Object System.Windows.Forms.ToolStripMenuItem("Opcions d'avís")
$miSound       = New-Object System.Windows.Forms.ToolStripMenuItem("Reproduir so en arribar")
$miSound.CheckOnClick = $true
$miPosCenter   = New-Object System.Windows.Forms.ToolStripMenuItem("Posició: Centre pantalla")
$miPosCorner   = New-Object System.Windows.Forms.ToolStripMenuItem("Posició: Cantonada inferior dreta")
$miOptions.DropDownItems.AddRange(@($miSound, (New-Object System.Windows.Forms.ToolStripSeparator), $miPosCenter, $miPosCorner))

# Submenú Gestió de configuració
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

# Esdeveniments del menú
$miPause.Add_Click({
    $state.IsPaused = $miPause.Checked
    Rebuild-IconsAndVisuals
})

$miThemeSystem.Add_Click({
    $config.Theme = "System"; Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks
})
$miThemeDark.Add_Click({
    $config.Theme = "Dark"; Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks
})
$miThemeLight.Add_Click({
    $config.Theme = "Light"; Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks
})

$miColorCustom.Add_Click({
    $dlg = New-Object System.Windows.Forms.ColorDialog
    $dlg.Color = [System.Drawing.ColorTranslator]::FromHtml($config.AccentHex)
    $dlg.FullOpen = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $config.AccentHex = "#{0:X2}{1:X2}{2:X2}" -f $dlg.Color.R, $dlg.Color.G, $dlg.Color.B
        Save-Config $config
        Rebuild-IconsAndVisuals
        Sync-MenuChecks
    }
})

$miIconEnvelope.Add_Click({
    $config.IconStyle = "Envelope"; Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks
})
$miIconCircle.Add_Click({
    $config.IconStyle = "Circle"; Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks
})
$miIconCustom.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
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
})

$miSound.Add_Click({
    $config.PlaySound = $miSound.Checked; Save-Config $config
})
$miPosCenter.Add_Click({
    $config.Position = "Center"; Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks
})
$miPosCorner.Add_Click({
    $config.Position = "BottomRight"; Save-Config $config; Rebuild-IconsAndVisuals; Sync-MenuChecks
})

$miOpenConfig.Add_Click({
    if (-not (Test-Path $configPath)) {
        Save-Config $config
    }
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
        $script:config = Get-FreshDefaultConfig
        Save-Config $config
        $timer.Interval = [int]$config.IntervalSec * 1000
        Rebuild-IconsAndVisuals
        Sync-MenuChecks
    }
})

# --- LÒGICA D'OUTLOOK I FINESTRA ---
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

function Update-AlertWindowText {
    param ([string]$Sender, [string]$Subject, [string]$Time)
    if (-not $state.Form -or $state.Form.IsDisposed) { return }

    if ($state.AlertCount -le 1) {
        $state.LblHeader.Text = "NOU CORREU REBUT ($Time)"
    } else {
        $state.LblHeader.Text = "$($state.AlertCount) CORREUS NOUS PENDENTS - ÚLTIM ($Time)"
    }
    $state.LblSender.Text  = "De: $Sender"
    $state.LblSubject.Text = "Assumpte: $Subject"
    $state.Form.TopMost    = $true
    $notifyIcon.Text       = "Outlook: $($state.AlertCount) correu(s) nou(s) sense acceptar"
}

function Show-Or-UpdateAlert {
    param ([string]$Sender, [string]$Subject, [string]$Time)

    $state.AlertCount++
    $notifyIcon.Icon = $state.IconAlert
    if ($config.PlaySound) { [System.Media.SystemSounds]::Exclamation.Play() }

    if ($state.IsAlertOpen -and $state.Form -and -not $state.Form.IsDisposed) {
        Apply-FormTheme
        Update-AlertWindowText -Sender $Sender -Subject $Subject -Time $Time
        return
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Avís de Nou Correu"
    $form.Size = New-Object System.Drawing.Size(480, 245)
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $banner = New-Object System.Windows.Forms.Panel
    $banner.Dock = "Top"
    $banner.Height = 8
    $form.Controls.Add($banner)

    $lblHeader = New-Object System.Windows.Forms.Label
    $lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblHeader.BackColor = [System.Drawing.Color]::Transparent
    $lblHeader.Location = New-Object System.Drawing.Point(22, 22)
    $lblHeader.Size = New-Object System.Drawing.Size(420, 22)
    $form.Controls.Add($lblHeader)

    $lblSender = New-Object System.Windows.Forms.Label
    $lblSender.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblSender.BackColor = [System.Drawing.Color]::Transparent
    $lblSender.Location = New-Object System.Drawing.Point(22, 52)
    $lblSender.Size = New-Object System.Drawing.Size(420, 26)
    $form.Controls.Add($lblSender)

    $lblSubject = New-Object System.Windows.Forms.Label
    $lblSubject.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblSubject.BackColor = [System.Drawing.Color]::Transparent
    $lblSubject.Location = New-Object System.Drawing.Point(22, 86)
    $lblSubject.Size = New-Object System.Drawing.Size(420, 56)
    $form.Controls.Add($lblSubject)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Acceptar"
    $btnOk.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnOk.Size = New-Object System.Drawing.Size(135, 38)
    $btnOk.Location = New-Object System.Drawing.Point(307, 152)
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.FlatStyle = "Flat"
    $btnOk.FlatAppearance.BorderSize = 0
    $btnOk.Add_Click({ $form.Close() })
    $form.Controls.Add($btnOk)

    $form.AcceptButton = $btnOk
    $form.Add_FormClosed({
        $state.IsAlertOpen = $false
        $state.AlertCount  = 0
        $state.Form        = $null
        Rebuild-IconsAndVisuals
    })

    $state.Form        = $form
    $state.Banner      = $banner
    $state.LblHeader   = $lblHeader
    $state.LblSender   = $lblSender
    $state.LblSubject  = $lblSubject
    $state.BtnOk       = $btnOk
    $state.IsAlertOpen = $true

    Apply-FormTheme
    Apply-FormPosition
    Update-AlertWindowText -Sender $Sender -Subject $Subject -Time $Time
    $form.Show()
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
    if ($state.Form -and -not $state.Form.IsDisposed) { $state.Form.Close() }
    $notifyIcon.Visible = $false
    $notifyIcon.Dispose()
    [System.Windows.Forms.Application]::Exit()
})

$timer.Start()
[System.Windows.Forms.Application]::Run()
