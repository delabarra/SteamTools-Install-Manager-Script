<#
.SYNOPSIS
    Install, update, uninstall, or export offline packs for SkyTools, SteamTools, and OpenSteamTools.

.DESCRIPTION
    Windows PowerShell manager for Steam integration DLLs and related tooling.

    - SkyTools / SteamTools: installs DLLs plus Millennium and Luatools when applicable.
    - OpenSteamTools: latest GitHub release; may prompt for a temporary Defender exclusion
      (UAC elevation only for that exclusion step).
    - Uninstall: removes integration files, STFixer / CloudRedirect artifacts, Millennium, Luatools, and registry keys.
    - Export pack: downloads components into a folder (optional zip) for offline use.
    - SteamTools install can optionally apply STFixer / CloudRedirect (prompted; not used by SkyTools or OpenSteamTools).
      CloudRedirect is removed when switching to SkyTools or OpenSteamTools.

    Run without -Action for the interactive menu. Logs go to Steam-Install.log beside this script
    (overwritten each run).

.PARAMETER Action
    InstallSkyTools, InstallSteamTools, InstallOpenSteamTools,
    ReinstallSkyTools, ReinstallSteamTools, ReinstallOpenSteamTools,
    Install, Reinstall, Uninstall, ExportPack,
    Download (alias for ExportPack), or LaunchSteam.
    Legacy aliases: InstallOpenSteamTool, ReinstallOpenSteamTool, -Integration OpenSteamTool.

.PARAMETER Integration
    SkyTools, SteamTools, or OpenSteamTools — required for -Action Install or Reinstall; used by ExportPack.

.PARAMETER Force
    Skip confirmation prompts (install/uninstall when Steam is running, and uninstall file-removal prompt).

.PARAMETER OutputDir
    Base output directory for ExportPack (creates skytools-pack, steamtools-pack, or
    opensteamtool-pack underneath).

.PARAMETER IncludeStPlugin
    Include unlocked game Lua scripts from the local Steam install in an export pack.

.NOTES
    Download sources (see also $SkyToolsBaseUrl, Get-SteamtoolsDllMirrors):
      SkyTools        — skyflare30.vercel.app/plugin/{dwmapi,xinput1_4,OpenSteamTool}.dll
      SteamTools      — files.catbox.moe (pinned DLLs), update.steamcdn.com (8s fallback)
      OpenSteamTools  — GitHub OpenSteam001/OpenSteamTool latest *-Release.zip asset
      Millennium      — GitHub SteamClientHomebrew/Millennium Windows zip (install + export)
      Luatools        — GitHub piqseu/ltsteamplugin latest zip
      STFixer         — GitHub Selectively11/CloudRedirect latest CloudRedirect.exe (/stfixer)

    Defender (OpenSteamTools install/export only):
      User consent → Add-MpPreference on temp + Steam/pack paths → work → Remove-MpPreference.
      If not admin, UAC elevation via hidden powershell.exe -Verb RunAs (see Invoke-ElevatedDefenderExclusionAction).
      SkyTools and SteamTools do not use Defender exclusions.

.EXAMPLE
    .\run.ps1

.EXAMPLE
    .\run.ps1 -Action InstallSkyTools

.EXAMPLE
    .\run.ps1 -Action Uninstall -Force

.EXAMPLE
    .\run.ps1 -Action Reinstall -Integration SteamTools

.EXAMPLE
    .\run.ps1 -Action ExportPack -Integration SteamTools -IncludeStPlugin
#>

param(
    [ValidateScript({
        if ([string]::IsNullOrWhiteSpace($_)) { return $true }
        $_ -in @(
            'InstallSkyTools', 'InstallSteamTools', 'InstallOpenSteamTools', 'InstallOpenSteamTool',
            'ReinstallSkyTools', 'ReinstallSteamTools', 'ReinstallOpenSteamTools', 'ReinstallOpenSteamTool',
            'Install', 'Reinstall', 'Uninstall', 'ExportPack', 'Download', 'LaunchSteam'
        )
    })]
    [string]$Action,
    [ValidateScript({
        if ([string]::IsNullOrWhiteSpace($_)) { return $true }
        $_ -in @('SkyTools', 'SteamTools', 'OpenSteamTools', 'OpenSteamTool')
    })]
    [string]$Integration,
    [switch]$Force,
    [string]$OutputDir,
    [switch]$IncludeStPlugin
)

if ([string]::IsNullOrWhiteSpace($Action)) { $Action = $null }
if ([string]::IsNullOrWhiteSpace($Integration)) { $Integration = $null }

if ($Integration -eq 'OpenSteamTool') { $Integration = 'OpenSteamTools' }
if ($Action -eq 'InstallOpenSteamTool') { $Action = 'InstallOpenSteamTools' }
if ($Action -eq 'ReinstallOpenSteamTool') { $Action = 'ReinstallOpenSteamTools' }

# "irm | iex" has no script file path; interactive menus do not work there.
# Download to a temp file and re-launch with -File in the same PowerShell host.
if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $bootstrapUrl = 'https://raw.githubusercontent.com/delabarra/SteamTools-Install-Manager-Script/main/run.ps1'
    $bootstrapFile = Join-Path $env:TEMP 'SteamTools-Install-Manager-Script-run.ps1'

    try {
        Invoke-WebRequest -Uri $bootstrapUrl -OutFile $bootstrapFile -UseBasicParsing
        if (-not (Test-Path -LiteralPath $bootstrapFile) -or (Get-Item -LiteralPath $bootstrapFile).Length -lt 1024) {
            throw 'Downloaded file is empty or incomplete.'
        }
    } catch {
        Write-Host "[ERR] Could not download run.ps1: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    $shell = (Get-Process -Id $PID).Path
    $bootstrapArgs = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-NoExit',
        '-File', $bootstrapFile
    )
    foreach ($name in @('Action', 'Integration', 'OutputDir')) {
        $value = Get-Variable -Name $name -ValueOnly
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $bootstrapArgs += "-$name"
            $bootstrapArgs += $value
        }
    }
    if ($PSBoundParameters.ContainsKey('Force') -and $Force) { $bootstrapArgs += '-Force' }
    if ($PSBoundParameters.ContainsKey('IncludeStPlugin') -and $IncludeStPlugin) { $bootstrapArgs += '-IncludeStPlugin' }

    & $shell @bootstrapArgs
    return
}

# Runtime defaults
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ProgressPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$ErrorActionPreference = 'Stop'

# Thrown when the user presses ESC to leave a submenu back to the main menu.
class MainMenuReturnException : System.Exception {
    MainMenuReturnException() : base('Returned to main menu.') {}
}

# -----------------------------------------------------------------------------
# Download source URLs (used by Save-RemoteFile / Invoke-SteamtoolsDllDownloads)
#
#   Component          | Primary source                                      | Fallback / notes
#   -------------------|-----------------------------------------------------|---------------------------
#   SkyTools           | $SkyToolsBaseUrl + each DLL filename                | none
#   SteamTools         | files.catbox.moe (pinned heom44/32p6f9)             | update.steamcdn.com, 8s timeout
#   OpenSteamTools     | GitHub releases API → *-Release.zip                 | Defender may block; see below
#   Millennium         | GitHub SteamClientHomebrew/Millennium zip           | none
#   Luatools           | GitHub piqseu/ltsteamplugin latest zip              | none
#   STFixer            | GitHub Selectively11/CloudRedirect .exe             | none
# -----------------------------------------------------------------------------

$MillenniumGithubRepo = 'SteamClientHomebrew/Millennium'
$LuatoolsGithubRepo   = 'piqseu/ltsteamplugin'

$PluginName = 'luatools'
$PluginUrl  = "https://github.com/$LuatoolsGithubRepo/releases/latest/download/ltsteamplugin.zip"

# STFixer / CloudRedirect — GitHub releases only
$CloudRedirectDownloadUrl = 'https://github.com/Selectively11/CloudRedirect/releases/latest/download/CloudRedirect.exe'

# SkyTools / OpenSteamTools share DLL names in the Steam folder; .dolintools-skytools marks a SkyTools install.
$SkyToolsBaseUrl = 'https://skyflare30.vercel.app/plugin/'
$SkyToolsMarkerFile = '.dolintools-skytools'
$SkyToolsConfigFile = 'opensteamtool.toml'
$SkyToolsFiles = @('dwmapi.dll', 'xinput1_4.dll', 'OpenSteamTool.dll')
$OpenSteamToolFiles = $SkyToolsFiles
$OpenSteamToolLatestReleaseUrl = 'https://api.github.com/repos/OpenSteam001/OpenSteamTool/releases/latest'

$script:LogFile = $null
$script:MenuIntegration = $null   # Set by export submenu; passed to Start-ExportPackFlow from the menu loop
$script:ManagerTitle = 'SteamTools Install Manager Script'

# =============================================================================
# Shared helpers
# =============================================================================

# --- Logging and console output ---

function Set-ManagerWindowTitle {
    try {
        if ($Host.UI -and $Host.UI.RawUI) {
            $Host.UI.RawUI.WindowTitle = $script:ManagerTitle
        }
    } catch {}

    try {
        [Console]::Title = $script:ManagerTitle
    } catch {}
}

function Get-ManagerDataRoot {
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        return $PSScriptRoot
    }
    if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        $dir = Split-Path -Parent $PSCommandPath
        if (-not [string]::IsNullOrWhiteSpace($dir)) {
            return $dir
        }
    }
    return (Get-Location).Path
}

function Initialize-SessionLog {
    $script:LogFile = Join-Path (Get-ManagerDataRoot) 'Steam-Install.log'

    $header = @(
        "=== $script:ManagerTitle ==="
        "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "User: $([Environment]::UserName) @ $([Environment]::MachineName)"
        "PowerShell: $($PSVersionTable.PSVersion)"
        "Script: $PSCommandPath"
        ''
    )

    Set-Content -Path $script:LogFile -Value $header -Encoding UTF8
}

function Write-Log {
    param([string]$Message = '')

    if (-not $script:LogFile) { return }

    $line = if ($Message) {
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message"
    } else {
        ''
    }

    [System.IO.File]::AppendAllText(
        $script:LogFile,
        "$line`r`n",
        [System.Text.UTF8Encoding]::new($false)
    )
}

function Write-Out {
    param(
        [string]$Message,
        [ConsoleColor]$ForegroundColor,
        [switch]$NoLog
    )

    if ($PSBoundParameters.ContainsKey('ForegroundColor')) {
        Write-Host $Message -ForegroundColor $ForegroundColor
    } else {
        Write-Host $Message
    }

    if ($Message -and -not $NoLog) { Write-Log $Message }
}

function Write-ActionBoundary {
    param(
        [string]$Action,
        [ValidateSet('start', 'completed', 'failed', 'cancelled')]
        [string]$Phase
    )

    Write-Log "--- $Action ($Phase) ---"
}

function Write-Step {
    param(
        [ValidateSet('OK', 'INFO', 'WARN', 'ERR')]
        [string]$Level,
        [string]$Message,
        [switch]$NoLog,
        [switch]$LogOnly
    )
    $color = switch ($Level) {
        'OK'   { 'Green' }
        'INFO' { 'Cyan' }
        'WARN' { 'Yellow' }
        'ERR'  { 'Red' }
    }
    $line = "[$Level] $Message"
    if (-not $LogOnly) {
        Write-Host $line -ForegroundColor $color
    }
    if (-not $NoLog) {
        Write-Log $line
    }
}

# --- Interactive input (single-key menus, yes/no, free text; ESC returns to main menu) ---

function Write-QuestionPrompt {
    param([string]$Prompt)

    # ANSI color 219: light pink on Windows 10+; Magenta is the non-ANSI fallback.
    if ([Environment]::OSVersion.Platform -eq 'Win32NT') {
        Write-Host ([char]27 + "[38;5;219m$Prompt " + [char]27 + '[0m') -NoNewline
    } else {
        Write-Host "$Prompt " -ForegroundColor Magenta -NoNewline
    }
}

function Get-ReadKeyCharacter {
    param([System.Management.Automation.Host.KeyInfo]$KeyInfo)

    if ($KeyInfo.KeyChar -and -not [char]::IsControl($KeyInfo.KeyChar)) {
        return [string]$KeyInfo.KeyChar
    }

    $shift = [bool]($KeyInfo.Modifiers -band [ConsoleModifiers]::Shift)
    $vk = $KeyInfo.VirtualKeyCode
    if ($vk -ge 48 -and $vk -le 57) {
        if ($shift) {
            return ')!@#$%^&*('[$vk - 48]
        }
        return [char]$vk
    }
    if ($vk -ge 65 -and $vk -le 90) {
        if ($shift) { return [char]$vk }
        return [string][char]::ToLower([char]$vk)
    }
    if ($vk -ge 96 -and $vk -le 105) { return [string]($vk - 96) }
    if ($vk -eq 32) { return ' ' }

    $oem = @{
        190 = @{ Normal = '.'; Shift = '>' }
        189 = @{ Normal = '-'; Shift = '_' }
        188 = @{ Normal = ','; Shift = '<' }
        220 = @{ Normal = '\'; Shift = '|' }
        191 = @{ Normal = '/'; Shift = '?' }
        186 = @{ Normal = ';'; Shift = ':' }
        222 = @{ Normal = ''''; Shift = '"' }
        219 = @{ Normal = '['; Shift = '{' }
        221 = @{ Normal = ']'; Shift = '}' }
        192 = @{ Normal = '`'; Shift = '~' }
        187 = @{ Normal = '='; Shift = '+' }
    }
    if ($oem.ContainsKey($vk)) {
        $entry = $oem[$vk]
        return if ($shift) { $entry.Shift } else { $entry.Normal }
    }

    return $null
}

function Read-ManagerKey {
    return $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

function Read-ManagerMenuKey {
    param(
        [string]$Prompt,
        [string[]]$ValidChoices,
        [string]$InvalidChoiceMessage,
        [ConsoleColor]$ForegroundColor = 'DarkGray',
        [switch]$NoLog
    )

    if (-not $NoLog) { Write-Log "> $Prompt" }
    Write-Host "$Prompt " -ForegroundColor $ForegroundColor -NoNewline

    if (-not ($Host.UI -and $Host.UI.RawUI)) {
        $value = (Read-Host).Trim()
        if (-not $NoLog) { Write-Log "< $(if ($value) { $value } else { '(default)' })" }
        return $value
    }

    while ($true) {
        $key = Read-ManagerKey
        if ($key.VirtualKeyCode -eq 27) {
            Write-Host ''
            if (-not $NoLog) { Write-Log '< ESC (main menu)' }
            throw [MainMenuReturnException]::new()
        }
        if ($key.VirtualKeyCode -eq 13) { continue }

        $ch = Get-ReadKeyCharacter -KeyInfo $key
        if (-not $ch -or $ch.Length -ne 1) { continue }
        if ($ValidChoices -notcontains $ch) {
            if ($InvalidChoiceMessage) {
                Write-Host ''
                Write-Out $InvalidChoiceMessage -ForegroundColor Yellow -NoLog
                Write-Host "$Prompt " -ForegroundColor $ForegroundColor -NoNewline
            }
            continue
        }

        Write-Host $ch
        if (-not $NoLog) { Write-Log "< $ch" }
        return $ch
    }
}

function Read-ManagerYesNoKey {
    param(
        [string]$Prompt,
        [switch]$NoLog
    )

    Write-QuestionPrompt -Prompt $Prompt

    if (-not ($Host.UI -and $Host.UI.RawUI)) {
        $value = Read-Host
        if (-not $NoLog) {
            $logged = if ($value) { $value } else { '(default)' }
            Write-Log "< $logged"
        }
        return $value
    }

    while ($true) {
        $key = Read-ManagerKey
        if ($key.VirtualKeyCode -eq 27) {
            Write-Host ''
            if (-not $NoLog) { Write-Log '< ESC (main menu)' }
            throw [MainMenuReturnException]::new()
        }
        if ($key.VirtualKeyCode -eq 13) {
            Write-Host ''
            if (-not $NoLog) { Write-Log '< (default)' }
            return ''
        }

        $ch = Get-ReadKeyCharacter -KeyInfo $key
        if ($ch -match '^[yY]$') {
            Write-Host $ch
            if (-not $NoLog) { Write-Log "< $ch" }
            return $ch.ToLowerInvariant()
        }
        if ($ch -match '^[nN]$') {
            Write-Host $ch
            if (-not $NoLog) { Write-Log "< $ch" }
            return $ch.ToLowerInvariant()
        }
    }
}

function Read-ManagerInput {
    param(
        [string]$Prompt,
        [ConsoleColor]$ForegroundColor,
        [switch]$Question,
        [switch]$NoLog
    )

    if (-not $NoLog) { Write-Log "> $Prompt" }

    if ($Question -and $Prompt -match '\[(Y/n|y/N)\]') {
        return Read-ManagerYesNoKey -Prompt $Prompt -NoLog:$NoLog
    }

    if ($Question) {
        Write-QuestionPrompt -Prompt $Prompt
    } elseif ($PSBoundParameters.ContainsKey('ForegroundColor')) {
        Write-Host "$Prompt " -ForegroundColor $ForegroundColor -NoNewline
    } else {
        Write-Host "$Prompt " -NoNewline
    }

    $value = if ($Host.UI -and $Host.UI.RawUI) {
        $buffer = [System.Text.StringBuilder]::new()
        while ($true) {
            $key = Read-ManagerKey
            if ($key.VirtualKeyCode -eq 27) {
                Write-Host ''
                if (-not $NoLog) { Write-Log '< ESC (main menu)' }
                throw [MainMenuReturnException]::new()
            }
            if ($key.VirtualKeyCode -eq 13) {
                Write-Host ''
                break
            }
            if ($key.VirtualKeyCode -eq 8) {
                if ($buffer.Length -gt 0) {
                    $null = $buffer.Remove($buffer.Length - 1, 1)
                    Write-Host "`b `b" -NoNewline
                }
                continue
            }

            $ch = Get-ReadKeyCharacter -KeyInfo $key
            if (-not $ch) { continue }

            [void]$buffer.Append($ch)
            Write-Host $ch -NoNewline
        }
        $buffer.ToString()
    } else {
        Read-Host
    }

    if (-not $NoLog) {
        $logged = if ($value) { $value } else { '(default)' }
        Write-Log "< $logged"
    }
    return $value
}

function Format-VersionLabel {
    param([string]$Version)
    if ($Version) { return $Version }
    return 'unknown'
}

# --- Steam client discovery and process control ---

function Get-SteamPath {
    foreach ($reg in @(
        'HKLM:\SOFTWARE\WOW6432Node\Valve\Steam',
        'HKLM:\SOFTWARE\Valve\Steam',
        'HKCU:\SOFTWARE\Valve\Steam'
    )) {
        if (-not (Test-Path $reg)) { continue }
        $path = (Get-ItemProperty -Path $reg -Name 'InstallPath' -ErrorAction SilentlyContinue).InstallPath
        if ($path -and (Test-Path (Join-Path $path 'steam.exe'))) { return $path }
    }
    throw 'Steam not found. Is it installed?'
}

function Get-SteamProcessNames {
    return @(
        'steam',
        'steamwebhelper',
        'gameoverlayui',
        'steamerrorreporter',
        'streaming_client'
    )
}

function Get-RunningSteamProcesses {
    $running = [System.Collections.Generic.List[System.Diagnostics.Process]]::new()
    foreach ($name in Get-SteamProcessNames) {
        foreach ($proc in @(Get-Process -Name $name -ErrorAction SilentlyContinue)) {
            $running.Add($proc)
        }
    }
    return $running
}

function Test-SteamRunning {
    $procs = Get-RunningSteamProcesses
    foreach ($proc in $procs) { $proc.Dispose() }
    return $procs.Count -gt 0
}

function Stop-Steam {
    param(
        [int]$GracefulTimeoutSec = 45,
        [switch]$Force
    )

    if (-not $Force -and (Test-SteamRunning)) {
        $answer = Read-ManagerInput 'Steam will be closed. Continue? [Y/n]' -Question
        if ($answer -and ($answer.Trim() -match '^n(o)?$')) {
            throw 'Cancelled.'
        }
    }

    if (-not (Test-SteamRunning)) { return }

    Write-Step INFO 'Closing Steam gracefully...'

    $steamExe = $null
    try {
        $steamExe = Join-Path (Get-SteamPath) 'steam.exe'
    } catch {
        Write-Log "Could not resolve Steam path for -shutdown: $($_.Exception.Message)"
    }

    if ($steamExe -and (Test-Path -LiteralPath $steamExe)) {
        try {
            $shutdown = Start-Process -FilePath $steamExe -ArgumentList '-shutdown' -PassThru -ErrorAction Stop
            if ($shutdown) { $shutdown.Dispose() }
        } catch {
            Write-Log "steam.exe -shutdown failed: $($_.Exception.Message)"
        }
    }

    Write-Step INFO 'Waiting for Steam to exit...'
    $deadline = [datetime]::UtcNow.AddSeconds($GracefulTimeoutSec)
    while ([datetime]::UtcNow -lt $deadline) {
        if (-not (Test-SteamRunning)) {
            Write-Step OK 'Steam closed.'
            return
        }
        Start-Sleep -Milliseconds 500
    }

    $remaining = Get-RunningSteamProcesses
    if (-not $remaining) {
        Write-Step OK 'Steam closed.'
        return
    }

    Write-Step WARN 'Steam did not close in time; stopping remaining processes...'
    try {
        foreach ($proc in $remaining) {
            try { $proc | Stop-Process -Force -ErrorAction SilentlyContinue } catch {}
        }

        Start-Sleep -Seconds 2

        if (-not (Test-SteamRunning)) {
            Write-Step OK 'Steam stopped.'
        } else {
            Write-Step WARN 'Some Steam processes are still running. Exit Steam from the system tray, then retry.'
        }
    } finally {
        foreach ($proc in $remaining) { $proc.Dispose() }
    }
}

function Start-SteamClient {
    param([string]$SteamPath)

    $steamExe = Join-Path $SteamPath 'steam.exe'
    if (-not (Test-Path -LiteralPath $steamExe)) {
        throw "steam.exe not found at $steamExe"
    }

    Start-Process -FilePath $steamExe
    Write-Step OK 'Steam launched.'
}

# --- Remote downloads (mirror fallback, parallel SteamTools DLL jobs) ---
#
# Save-RemoteFile tries each URL in order; per-URL timeouts come from -TimeoutSeconds or -TimeoutSec.
# SteamTools uses Get-SteamtoolsDllMirrors: catbox first (90s), steamcdn second (8s quick-fail).

function Save-RemoteFile {
    param(
        [string[]]$Urls,
        [string]$Destination,
        [int]$TimeoutSec = 120,
        [int[]]$TimeoutSeconds,
        [switch]$Quiet
    )

    for ($i = 0; $i -lt $Urls.Count; $i++) {
        $url = $Urls[$i]
        $timeout = $TimeoutSec
        if ($TimeoutSeconds -and $i -lt $TimeoutSeconds.Count) {
            $timeout = $TimeoutSeconds[$i]
        }

        try {
            if (-not $Quiet) {
                Write-Step INFO "Downloading $(Split-Path $Destination -Leaf) from $url"
            }
            Invoke-WebRequest -Uri $url -OutFile $Destination -UseBasicParsing -TimeoutSec $timeout
            if ((Test-Path $Destination) -and ((Get-Item $Destination).Length -gt 0)) {
                return $url
            }
        } catch {
            $message = "Download failed: $($_.Exception.Message)"
            if ($i -lt ($Urls.Count - 1)) {
                Write-Log "$message (trying next mirror)"
            } elseif (-not $Quiet) {
                Write-Step WARN $message
            } else {
                Write-Log $message
            }
        }
    }

    throw "Could not download $(Split-Path $Destination -Leaf)."
}

function Get-SteamtoolsDllMirrors {
    param([ValidateSet('xinput', 'dwmapi')][string]$Dll)

    # Catbox hosts pinned builds; steamcdn is tried with a short timeout so a slow updater does not block install.
    if ($Dll -eq 'xinput') {
        return @(
            @{ Url = 'https://files.catbox.moe/heom44.dll'; TimeoutSec = 90 },
            @{ Url = 'http://update.steamcdn.com/update'; TimeoutSec = 8 }
        )
    }

    return @(
        @{ Url = 'https://files.catbox.moe/32p6f9.dll'; TimeoutSec = 90 },
        @{ Url = 'http://update.steamcdn.com/dwmapi'; TimeoutSec = 8 }
    )
}

function Get-SteamtoolsDllMirrorSpec {
    param([ValidateSet('xinput', 'dwmapi')][string]$Dll)

    $mirrors = Get-SteamtoolsDllMirrors -Dll $Dll
    return @{
        Urls     = ($mirrors | ForEach-Object { $_.Url }) -join '|'
        Timeouts = ($mirrors | ForEach-Object { [string]$_.TimeoutSec }) -join ','
    }
}

function Save-SteamtoolsDll {
    param(
        [ValidateSet('xinput', 'dwmapi')][string]$Dll,
        [string]$Destination,
        [switch]$Quiet
    )

    $mirrors = Get-SteamtoolsDllMirrors -Dll $Dll
    $null = Save-RemoteFile `
        -Urls ($mirrors | ForEach-Object { $_.Url }) `
        -TimeoutSeconds ($mirrors | ForEach-Object { $_.TimeoutSec }) `
        -Destination $Destination `
        -Quiet:$Quiet
}

function Invoke-SteamtoolsDllDownloads {
    param(
        [string]$XinputDestination,
        [string]$DwmapiDestination,
        [switch]$Quiet
    )

    # Download both DLLs in parallel; mirror URL lists are passed as strings because Start-Job runs out-of-process.
    $jobScript = {
        param(
            [string]$Destination,
            [string]$UrlList,
            [string]$TimeoutList,
            [string]$Label
        )

        $urls = $UrlList -split '\|'
        $timeouts = $TimeoutList -split ','
        $warnings = [System.Collections.Generic.List[string]]::new()

        for ($i = 0; $i -lt $urls.Count; $i++) {
            $url = $urls[$i]
            $timeout = if ($i -lt $timeouts.Count) { [int]$timeouts[$i] } else { 90 }
            try {
                Invoke-WebRequest -Uri $url -OutFile $Destination -UseBasicParsing -TimeoutSec $timeout
                if ((Test-Path $Destination) -and ((Get-Item $Destination).Length -gt 0)) {
                    return ,@($warnings)   # unary comma: always return an array (even when empty)
                }
            } catch {
                $warnings.Add("${url}: $($_.Exception.Message)")
            }
        }

        if ($warnings.Count) {
            throw ($warnings -join '; ')
        }
        throw "Could not download $Label."
    }

    $downloads = @(
        @{
            Label    = 'xinput1_4.dll'
            Dest     = $XinputDestination
            Spec     = Get-SteamtoolsDllMirrorSpec -Dll 'xinput'
        },
        @{
            Label    = 'dwmapi.dll'
            Dest     = $DwmapiDestination
            Spec     = Get-SteamtoolsDllMirrorSpec -Dll 'dwmapi'
        }
    )

    $jobs = foreach ($download in $downloads) {
        Start-Job -ScriptBlock $jobScript -ArgumentList `
            $download.Dest, $download.Spec.Urls, $download.Spec.Timeouts, $download.Label
    }

    try {
        Wait-Job $jobs | Out-Null

        foreach ($i in 0..($downloads.Count - 1)) {
            try {
                $warnings = Receive-Job $jobs[$i] -ErrorAction Stop
                foreach ($warning in $warnings) {
                    $message = "Download failed: $warning"
                    if (-not $Quiet) {
                        Write-Step WARN $message
                    } else {
                        Write-Log $message
                    }
                }
            } catch {
                throw "Could not download $($downloads[$i].Label). $($_.Exception.Message)"
            }
        }
    } finally {
        Remove-Job $jobs -Force -ErrorAction SilentlyContinue
    }
}

# --- Version labels, GitHub releases, and plugin.json parsing ---

function Get-FileSha256Hex {
    param([string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpper()
}

function Get-DllBuildLabel {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $null }

    $versionInfo = (Get-Item -LiteralPath $Path).VersionInfo
    if ($versionInfo.FileVersion -and $versionInfo.FileVersion -notmatch '^0\.0\.0') {
        return $versionInfo.FileVersion
    }

    # Unsigned or placeholder version metadata — show a short hash instead.
    $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    return "build $($hash.Substring(0, 8))"
}

function Get-GithubLatestTag {
    param([string]$Repo)

    $release = Invoke-RestMethod `
        -Uri "https://api.github.com/repos/$Repo/releases/latest" `
        -Headers @{ 'User-Agent' = 'SteamTools-Install-Manager-Script' }

    return $release.tag_name
}

function Get-GithubWindowsZipAsset {
    param([string]$Repo)

    $release = Invoke-RestMethod `
        -Uri "https://api.github.com/repos/$Repo/releases/latest" `
        -Headers @{ 'User-Agent' = 'SteamTools-Install-Manager-Script' }

    foreach ($asset in $release.assets) {
        if ($asset.name -match 'windows' -and $asset.name -match 'zip') {
            return @{
                Url     = $asset.browser_download_url
                Version = $release.tag_name
                Name    = $asset.name
            }
        }
    }

    throw "No Windows zip asset found for $Repo"
}

function Get-PluginVersionFromJson {
    param([string]$JsonText)

    try {
        $meta = $JsonText | ConvertFrom-Json
        if ($meta.version) { return [string]$meta.version }
    } catch {}

    return $null
}

function Get-PluginVersionFromDir {
    param([string]$PluginDir)

    $jsonPath = Join-Path $PluginDir 'plugin.json'
    if (-not (Test-Path $jsonPath)) { return $null }
    return Get-PluginVersionFromJson -JsonText (Get-Content $jsonPath -Raw -Encoding UTF8)
}

function Get-PluginVersionFromZip {
    param([string]$ZipPath)

    if (-not (Test-Path $ZipPath)) { return $null }

    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        $entry = $zip.Entries | Where-Object { $_.Name -eq 'plugin.json' } | Select-Object -First 1
        if (-not $entry) { return $null }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        try {
            return Get-PluginVersionFromJson -JsonText $reader.ReadToEnd()
        } finally {
            $reader.Close()
        }
    } finally {
        if ($zip) { $zip.Dispose() }
    }
}

function Test-VersionUpdateNeeded {
    param(
        [string]$Installed,
        [string]$Latest
    )

    if (-not $Latest) { return $false }
    if (-not $Installed) { return $true }

    # Tag/string compare only — not full semantic versioning.
    $installedNorm = $Installed.Trim().TrimStart('v')
    $latestNorm    = $Latest.Trim().TrimStart('v')
    return $installedNorm -ne $latestNorm
}

function Find-PluginDir {
    param(
        [string]$SteamPath,
        [string]$MillDir,
        [string]$Name
    )

    $scanDirs = @(
        (Join-Path $MillDir 'plugins'),
        (Join-Path $SteamPath 'plugins')
    ) | Where-Object { Test-Path $_ } | Select-Object -Unique

    foreach ($dir in $scanDirs) {
        foreach ($folder in Get-ChildItem $dir -Directory -ErrorAction SilentlyContinue) {
            $jsonPath = Join-Path $folder.FullName 'plugin.json'
            if (-not (Test-Path $jsonPath)) { continue }
            try {
                $meta = Get-Content $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($meta.name -eq $Name) { return $folder.FullName }
            } catch {}
        }
    }

    return $null
}

# --- Archives, paths, and Lua script folder layout ---

function Expand-ZipEntry {
    param(
        [string]$ZipPath,
        [string]$DestinationRoot,
        [switch]$UnblockExtracted
    )

    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) { continue }

            $dest   = Join-Path $DestinationRoot $entry.FullName
            $parent = Split-Path $dest -Parent
            if ($parent) { [void][System.IO.Directory]::CreateDirectory($parent) }
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dest, $true)
            if ($UnblockExtracted) {
                Unblock-File -LiteralPath $dest -ErrorAction SilentlyContinue
            }
        }
    } finally {
        if ($zip) { $zip.Dispose() }
    }
}

function Remove-PathIfExists {
    param(
        [string]$Path,
        [switch]$Recurse,
        [int]$MaxAttempts = 1
    )

    if (-not (Test-Path -LiteralPath $Path)) { return $false }

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            if ($Recurse) {
                Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            } else {
                Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
            }
            return $true
        } catch {
            if ($attempt -lt $MaxAttempts) {
                Start-Sleep -Milliseconds 600
                continue
            }
            throw
        }
    }

    return $false
}

function Move-LuaFiles {
    param(
        [string]$Source,
        [string]$Destination
    )

    if (-not (Test-Path -LiteralPath $Source -PathType Container)) { return }

    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    Get-ChildItem -LiteralPath $Source -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '\.lua(\.disabled)?$' } |
        Move-Item -Destination $Destination -Force
}

function Get-STFixerArtifactPaths {
    param([string]$SteamPath)

    $paths = [System.Collections.Generic.List[string]]::new()

    foreach ($name in @(
        'cloud_redirect.dll',
        'stella_fallback.dll',
        'stella.cfg',
        'stella_debug.log',
        'dwmapi.dll.bak',
        'dwmapi.dll.orig',
        'xinput1_4.dll.bak',
        'xinput1_4.dll.orig'
    )) {
        $paths.Add((Join-Path $SteamPath $name))
    }

    $paths.Add((Join-Path $SteamPath 'cloud_redirect'))

    # STFixer patches the SteamTools payload cache under appcache\httpcache\3b.
    $cacheDir = Join-Path $SteamPath 'appcache\httpcache\3b'
    if (Test-Path -LiteralPath $cacheDir -PathType Container) {
        foreach ($file in Get-ChildItem -LiteralPath $cacheDir -File -Force -ErrorAction SilentlyContinue) {
            $name = $file.Name
            $isPayloadCache = $name.Length -eq 16 -and $name -match '^[0-9A-Fa-f]{16}$' `
                -and $file.Length -gt 500000 -and $file.Length -lt 5000000
            $isPayloadBackup = $name -match '^[0-9A-Fa-f]{16}\.(bak|orig)$' `
                -and $file.Length -gt 500000 -and $file.Length -lt 5000000
            if ($isPayloadCache -or $isPayloadBackup) {
                $paths.Add($file.FullName)
            }
        }
    }

    return $paths | Select-Object -Unique
}

function Get-CloudRedirectAppDataPath {
    return Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'CloudRedirect'
}

function Remove-CloudRedirect {
    param([string]$SteamPath)

    # STFixer / CloudRedirect only work with SteamTools; strip when switching integrations.
    foreach ($path in (Get-STFixerArtifactPaths -SteamPath $SteamPath)) {
        $isDir = Test-Path -LiteralPath $path -PathType Container
        Remove-Item -LiteralPath $path -Recurse:$isDir -Force -ErrorAction SilentlyContinue
    }

    Remove-Item -LiteralPath (Get-CloudRedirectAppDataPath) -Recurse -Force -ErrorAction SilentlyContinue
}

function Test-STFixerSupport {
    param([string]$SteamPath)

    return Test-Path -LiteralPath (Join-Path $SteamPath 'cloud_redirect.dll')
}

function Test-UserAllowsSTFixer {
    $answer = Read-ManagerInput 'Install STFixer / CloudRedirect? (Fixes cloud-save issues). [Y/n]' -Question
    if ($answer -and ($answer.Trim() -match '^n(o)?$')) {
        return $false
    }
    return $true
}

function Install-STFixerSupport {
    param([string]$SteamPath)

    Write-Step INFO 'Installing STFixer / CloudRedirect...'

    $tempDir = Join-Path ([IO.Path]::GetTempPath()) ("SteamInstall.STFixer." + [Guid]::NewGuid().ToString('N'))
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

    try {
        $exePath = Join-Path $tempDir 'CloudRedirect.exe'
        $null = Save-RemoteFile -Urls @($CloudRedirectDownloadUrl) -Destination $exePath -Quiet -TimeoutSec 120

        Write-Log 'Running CloudRedirect /stfixer...'
        $proc = Start-Process -FilePath $exePath -ArgumentList '/stfixer' -WindowStyle Hidden -Wait -PassThru

        if ($proc.ExitCode -eq 0 -and (Test-STFixerSupport $SteamPath) -and (Test-Steamtools $SteamPath)) {
            Write-Step OK 'STFixer / CloudRedirect support installed.'
            Write-Step INFO 'For Lua game cloud saves, run CloudRedirect.exe and sign in to a cloud provider.'
            return $true
        }

        if ($proc.ExitCode) {
            Write-Log "CloudRedirect /stfixer exit code: $($proc.ExitCode)"
        }

        Write-Step INFO 'STFixer / CloudRedirect skipped.'
        return $false
    } catch {
        Write-Log "STFixer / CloudRedirect failed: $($_.Exception.Message)"
        Write-Step INFO 'STFixer / CloudRedirect skipped.'
        return $false
    } finally {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Sync-STFixerSupport {
    param(
        [string]$SteamPath,
        [switch]$Reapply,
        [switch]$SkipUserPrompt
    )

    $alreadyPresent = (Test-STFixerSupport $SteamPath) -and (Test-Steamtools $SteamPath)

    if ($alreadyPresent -and -not $Reapply) {
        Write-Step OK 'STFixer / CloudRedirect support already present.'
        return
    }

    if ($Reapply -and $alreadyPresent) {
        Write-Step INFO 'Reapplying STFixer patches after SteamTools update...'
        [void](Install-STFixerSupport -SteamPath $SteamPath)
        return
    }

    if ($SkipUserPrompt) {
        Write-Step INFO 'STFixer / CloudRedirect skipped.'
        return
    }

    if (-not (Test-UserAllowsSTFixer)) {
        Write-Step INFO 'STFixer / CloudRedirect skipped.'
        return
    }

    [void](Install-STFixerSupport -SteamPath $SteamPath)
}

function Add-SteamIntegrationRemovalTargets {
    param(
        [string]$SteamPath,
        [System.Collections.Generic.List[string]]$Targets
    )

    foreach ($file in @(
        'dwmapi.dll',
        'xinput1_4.dll',
        'OpenSteamTool.dll',
        $SkyToolsConfigFile,
        $SkyToolsMarkerFile,
        'steamtools-sync-temp',
        'skytools-sync-temp',
        'opensteamtool-sync-temp',
        "$PluginName.zip"
    )) {
        $Targets.Add((Join-Path $SteamPath $file))
    }

    foreach ($path in (Get-STFixerArtifactPaths -SteamPath $SteamPath)) {
        $Targets.Add($path)
    }

    $appDataConfig = Get-CloudRedirectAppDataPath
    if (Test-Path -LiteralPath $appDataConfig) {
        $Targets.Add($appDataConfig)
    }

    # Catch extra OpenSteamTool artifacts (e.g. versioned folders) not listed above.
    Get-ChildItem -LiteralPath $SteamPath -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like 'OpenSteamTool*' } |
        ForEach-Object { $Targets.Add($_.FullName) }
}

function Move-LuaScriptsToStPlugin {
    param([string]$SteamPath)

    # SkyTools / SteamTools expect unlocked-game scripts under config\stplug-in.
    $steamToolsScripts = Join-Path $SteamPath 'config\stplug-in'
    $openSteamScripts  = Join-Path $SteamPath 'config\lua'
    New-Item -ItemType Directory -Path $steamToolsScripts -Force | Out-Null
    Move-LuaFiles $openSteamScripts $steamToolsScripts

    if ((Test-Path -LiteralPath $openSteamScripts -PathType Container) -and
        -not (Get-ChildItem -LiteralPath $openSteamScripts -Force -ErrorAction SilentlyContinue)) {
        Remove-Item -LiteralPath $openSteamScripts -Force -ErrorAction SilentlyContinue
    }
}

function Move-LuaScriptsToLua {
    param([string]$SteamPath)

    # OpenSteamTools expects the same scripts under config\lua.
    $steamToolsScripts = Join-Path $SteamPath 'config\stplug-in'
    $openSteamScripts  = Join-Path $SteamPath 'config\lua'
    New-Item -ItemType Directory -Path $openSteamScripts -Force | Out-Null
    Move-LuaFiles $steamToolsScripts $openSteamScripts
}

# --- Windows Defender temporary exclusions (OpenSteamTools only) ---
#
# Why: OpenSteamTools DLLs/zips are often flagged as potentially unwanted. SkyTools/SteamTools
# use different payloads and do not go through this path.
#
# Flow:
#   1. Test-UserAllowsDefenderExclusion — interactive [Y/n] (skipped if consent already given)
#   2. Enable-OpenSteamToolDefenderExclusions — excludes temp download dir + Steam folder (or pack dir)
#   3. Download / extract / copy OpenSteamTool (Test-IsDefenderBlock on failure)
#   4. Disable-OpenSteamToolDefenderExclusions — always in finally; exclusions are never left behind
#
# Elevation (Invoke-ElevatedDefenderExclusionAction):
#   - Already admin → Add-MpPreference / Remove-MpPreference in this process
#   - Standard user → write paths to %TEMP%\SteamInstall.Defender.{guid}.json, spawn hidden
#     powershell.exe -Verb RunAs with -EncodedCommand that reads the JSON and calls Defender cmdlets
#   - User declining UAC or missing Defender module → install/export aborts with a clear message

function Test-IsDefenderBlock {
    param([string]$Message)

    return $Message -match 'virus|potentially unwanted|malware|0x800700E1|amenaza|software potencialmente no deseado'
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-DefenderExclusionCmdletsAvailable {
    return [bool](Get-Command Add-MpPreference -ErrorAction SilentlyContinue) `
        -and [bool](Get-Command Remove-MpPreference -ErrorAction SilentlyContinue)
}

function Test-UserAllowsDefenderExclusion {
    $answer = Read-ManagerInput 'Windows Defender may block OpenSteamTools. Add a temporary exclusion for download and install? [Y/n]' -Question -NoLog
    if ($answer -and ($answer.Trim() -match '^n(o)?$')) {
        return $false
    }
    return $true
}

function Add-TemporaryDefenderExclusions {
    param([string[]]$Paths)

    foreach ($path in ($Paths | Select-Object -Unique)) {
        $fullPath = [IO.Path]::GetFullPath($path)
        if (-not (Test-Path -LiteralPath $fullPath)) {
            New-Item -Path $fullPath -ItemType Directory -Force | Out-Null
        }
        Add-MpPreference -ExclusionPath $fullPath -ErrorAction Stop
        Write-Log "Defender exclusion added: $fullPath"
    }
}

function Remove-TemporaryDefenderExclusions {
    param([string[]]$Paths)

    foreach ($path in ($Paths | Select-Object -Unique)) {
        $fullPath = [IO.Path]::GetFullPath($path)
        Remove-MpPreference -ExclusionPath $fullPath -ErrorAction SilentlyContinue
        Write-Log "Defender exclusion removed: $fullPath"
    }
}

function Remove-TemporaryDefenderProcessExclusions {
    param([string[]]$Processes)

    foreach ($processPath in ($Processes | Select-Object -Unique)) {
        if (-not $processPath) { continue }
        Remove-MpPreference -ExclusionProcess $processPath -ErrorAction SilentlyContinue
        Write-Log "Defender process exclusion removed: $processPath"
    }
}

function Add-TemporaryDefenderProcessExclusions {
    param([string[]]$Processes)

    foreach ($processPath in ($Processes | Select-Object -Unique)) {
        if (-not $processPath -or -not (Test-Path -LiteralPath $processPath)) { continue }
        Add-MpPreference -ExclusionProcess $processPath -ErrorAction Stop
        Write-Log "Defender process exclusion added: $processPath"
    }
}

function Invoke-ElevatedDefenderExclusionAction {
    param(
        [ValidateSet('Add', 'Remove')]
        [string]$Action,
        [string[]]$Paths,
        [string[]]$Processes
    )

    $uniquePaths = @($Paths | Where-Object { $_ } | Select-Object -Unique | ForEach-Object { [IO.Path]::GetFullPath($_) })
    $uniqueProcesses = @($Processes | Where-Object { $_ } | Select-Object -Unique)
    if (-not $uniquePaths -and -not $uniqueProcesses) {
        return
    }

    if (Test-IsAdministrator) {
        if ($Action -eq 'Add') {
            if ($uniquePaths) { Add-TemporaryDefenderExclusions -Paths $uniquePaths }
            if ($uniqueProcesses) { Add-TemporaryDefenderProcessExclusions -Processes $uniqueProcesses }
        } else {
            if ($uniquePaths) { Remove-TemporaryDefenderExclusions -Paths $uniquePaths }
            if ($uniqueProcesses) { Remove-TemporaryDefenderProcessExclusions -Processes $uniqueProcesses }
        }
        return
    }

    # Non-admin: spawn a hidden elevated PowerShell to call Add-MpPreference / Remove-MpPreference.
    $payloadPath = Join-Path ([IO.Path]::GetTempPath()) ("SteamInstall.Defender.$([Guid]::NewGuid().ToString('N')).json")
    try {
        (@{
            Action    = $Action
            Paths     = $uniquePaths
            Processes = $uniqueProcesses
        }) | ConvertTo-Json | Set-Content -LiteralPath $payloadPath -Encoding UTF8

        $payloadEscaped = $payloadPath.Replace("'", "''")
        $scriptBody = @"
`$ErrorActionPreference = 'Stop'
if (-not (Get-Command Add-MpPreference -ErrorAction SilentlyContinue)) { exit 1 }
`$data = Get-Content -LiteralPath '$payloadEscaped' -Raw | ConvertFrom-Json
foreach (`$path in @(`$data.Paths)) {
    if (`$data.Action -eq 'Add') {
        if (-not (Test-Path -LiteralPath `$path)) { New-Item -Path `$path -ItemType Directory -Force | Out-Null }
        Add-MpPreference -ExclusionPath `$path -ErrorAction Stop
    } else {
        Remove-MpPreference -ExclusionPath `$path -ErrorAction SilentlyContinue
    }
}
foreach (`$processPath in @(`$data.Processes)) {
    if (-not `$processPath) { continue }
    if (`$data.Action -eq 'Add') {
        if (Test-Path -LiteralPath `$processPath) { Add-MpPreference -ExclusionProcess `$processPath -ErrorAction Stop }
    } else {
        Remove-MpPreference -ExclusionProcess `$processPath -ErrorAction SilentlyContinue
    }
}
exit 0
"@

        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($scriptBody))
        try {
            $proc = Start-Process -FilePath 'powershell.exe' `
                -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-EncodedCommand', $encoded `
                -Verb RunAs -Wait -PassThru -WindowStyle Hidden
        } catch {
            $message = if ($Action -eq 'Add') {
                'Could not add the temporary Windows Defender exclusion. Administrator permission is required.'
            } else {
                'Could not remove the temporary Windows Defender exclusion.'
            }
            throw $message
        }

        if (-not $proc -or $proc.ExitCode -ne 0) {
            if ($Action -eq 'Add') {
                throw 'Could not add the temporary Windows Defender exclusion. Administrator permission is required.'
            }
            Write-Log "Elevated Defender exclusion remove may have failed (exit $($proc.ExitCode))."
            return
        }

        foreach ($path in $uniquePaths) {
            if ($Action -eq 'Add') {
                Write-Log "Defender exclusion added: $path"
            } else {
                Write-Log "Defender exclusion removed: $path"
            }
        }
        foreach ($processPath in $uniqueProcesses) {
            if ($Action -eq 'Add') {
                Write-Log "Defender process exclusion added: $processPath"
            } else {
                Write-Log "Defender process exclusion removed: $processPath"
            }
        }
    } finally {
        Remove-Item -LiteralPath $payloadPath -Force -ErrorAction SilentlyContinue
    }
}

function Enable-OpenSteamToolDefenderExclusions {
    param(
        [string]$TempDir,
        [string]$DestinationPath,
        [string[]]$ExtraPaths,
        [string]$CancelledMessage = 'OpenSteamTools install cancelled. A temporary Windows Defender exclusion is required.',
        [switch]$SkipConsentPrompt
    )

    if (-not $SkipConsentPrompt -and -not (Test-UserAllowsDefenderExclusion)) {
        throw $CancelledMessage
    }

    if (-not (Test-DefenderExclusionCmdletsAvailable)) {
        throw 'Windows Defender exclusion commands are not available on this system.'
    }

    $paths = @($TempDir, $DestinationPath) + @($ExtraPaths) |
        Where-Object { $_ } |
        ForEach-Object { [IO.Path]::GetFullPath($_) } |
        Select-Object -Unique

    $processes = @()
    try {
        $processPath = (Get-Process -Id $PID -ErrorAction Stop).Path
        if ($processPath) { $processes += $processPath }
    } catch {
        Write-Log "Could not resolve current PowerShell path for Defender process exclusion: $($_.Exception.Message)"
    }

    if (-not (Test-IsAdministrator)) {
        Write-Step INFO 'Administrator permission is required to add the temporary Windows Defender exclusion.'
    }

    Invoke-ElevatedDefenderExclusionAction -Action Add -Paths $paths -Processes $processes
    Start-Sleep -Milliseconds 1500
    Write-Step INFO 'Temporary Windows Defender exclusions added.'

    return [PSCustomObject]@{
        Paths     = $paths
        Processes = $processes
    }
}

function Disable-OpenSteamToolDefenderExclusions {
    param($Exclusions)

    if (-not $Exclusions -or -not (Test-DefenderExclusionCmdletsAvailable)) {
        return
    }

    $paths = $null
    $processes = $null
    if ($Exclusions -is [string[]]) {
        $paths = $Exclusions
    } else {
        $paths = $Exclusions.Paths
        $processes = $Exclusions.Processes
    }

    Invoke-ElevatedDefenderExclusionAction -Action Remove -Paths $paths -Processes $processes
    Write-Step INFO 'Temporary Windows Defender exclusions removed.'
}

# --- Integration detection, config templates, and sync into the Steam folder ---

function Remove-CompetingSteamIntegrations {
    param(
        [string]$SteamPath,
        [ValidateSet('SkyTools', 'SteamTools', 'OpenSteamTools')]
        [string]$TargetIntegration
    )

    $targets = [System.Collections.Generic.List[string]]::new()
    Add-SteamIntegrationRemovalTargets -SteamPath $SteamPath -Targets $targets

    $existing = $targets | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -Unique
    if (-not $existing) {
        Remove-SteamtoolsRegistry -Quiet
        return
    }

    Write-Step INFO 'Removing previous SteamTools / SkyTools integration...'

    # Move game scripts to the target layout before deleting DLLs and config.
    if ($TargetIntegration -in 'SkyTools', 'SteamTools') {
        Move-LuaScriptsToStPlugin -SteamPath $SteamPath
    }

    if ($TargetIntegration -ne 'SteamTools') {
        Remove-CloudRedirect $SteamPath
    }

    foreach ($path in $existing) {
        $isDir = Test-Path -LiteralPath $path -PathType Container
        if (Remove-PathIfExists -Path $path -Recurse:$isDir) {
            Write-Log "Removed $path"
        }
    }

    Remove-SteamtoolsRegistry -Quiet
}

function Get-SkyToolsConfig {
    # Static manifest provider; SkyTools does not read DolinTools settings.
    return "[log]`r`nlevel = `"info`"`r`n`r`n[manifest]`r`nurl = `"opensteamtool`"`r`ntimeout_resolve_ms = 5000`r`ntimeout_connect_ms = 5000`r`ntimeout_send_ms = 10000`r`ntimeout_recv_ms = 10000`r`n"
}

function Test-SkyTools {
    param([string]$SteamPath)

    foreach ($file in $SkyToolsFiles) {
        if (-not (Test-Path -LiteralPath (Join-Path $SteamPath $file))) { return $false }
    }

    return (Test-Path -LiteralPath (Join-Path $SteamPath $SkyToolsMarkerFile))
}

function Get-SkyToolsVersionLabel {
    param([string]$SteamPath)
    return Get-DllBuildLabel (Join-Path $SteamPath 'dwmapi.dll')
}

function Sync-SkyTools {
    param([string]$SteamPath)

    $tempDir = Join-Path $SteamPath 'skytools-sync-temp'
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

    try {
        $items = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($fileName in $SkyToolsFiles) {
            $tempPath = Join-Path $tempDir $fileName
            $destPath = Join-Path $SteamPath $fileName
            $null = Save-RemoteFile -Urls @("$SkyToolsBaseUrl$fileName") -Destination $tempPath -Quiet
            $items.Add(@{
                Temp       = $tempPath
                Dest       = $destPath
                Label      = $fileName
                RemoteHash = Get-FileSha256Hex $tempPath
                LocalHash  = if (Test-Path -LiteralPath $destPath) { Get-FileSha256Hex $destPath } else { $null }
            })
        }

        $configTemp = Join-Path $tempDir $SkyToolsConfigFile
        $configDest = Join-Path $SteamPath $SkyToolsConfigFile
        Set-Content -LiteralPath $configTemp -Value (Get-SkyToolsConfig) -Encoding UTF8
        $configHash = Get-FileSha256Hex $configTemp
        $localConfigHash = if (Test-Path -LiteralPath $configDest) {
            Get-FileSha256Hex $configDest
        } else {
            $null
        }

        $markerPath = Join-Path $SteamPath $SkyToolsMarkerFile
        # Skip reinstall when remote hashes match local files and the SkyTools marker is present.
        $needsUpdate = ($items | Where-Object { $_.LocalHash -ne $_.RemoteHash }).Count -gt 0 `
            -or $configHash -ne $localConfigHash `
            -or -not (Test-Path -LiteralPath $markerPath)

        if (-not $needsUpdate -and (Test-SkyTools $SteamPath)) {
            $version = Format-VersionLabel (Get-SkyToolsVersionLabel $SteamPath)
            Write-Step OK "SkyTools already up to date ($version)"
            return
        }

        $version = Format-VersionLabel (Get-DllBuildLabel $items[0].Temp)
        Write-Step INFO "Installing SkyTools ($version)..."

        Move-LuaScriptsToStPlugin -SteamPath $SteamPath

        foreach ($item in $items) {
            Copy-Item -LiteralPath $item.Temp -Destination $item.Dest -Force
        }

        Copy-Item -LiteralPath $configTemp -Destination $configDest -Force
        Set-Content -LiteralPath $markerPath -Value 'SkyTools installed by SteamTools Install Manager Script' -Encoding ASCII

        if (-not (Test-SkyTools $SteamPath)) {
            throw 'SkyTools install failed - expected files were not found after install.'
        }

        # Post-copy hash check catches partial writes or AV tampering.
        foreach ($item in $items) {
            $installedHash = Get-FileSha256Hex $item.Dest
            if ($installedHash -ne $item.RemoteHash) {
                throw "SkyTools install failed - $($item.Label) hash mismatch after install."
            }
        }

        Write-Step OK "SkyTools installed ($version)"
    } finally {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-OpenSteamToolConfig {
    $provider = 'opensteamtool'
    $resolve = 5000
    $connect = 5000
    $send = 10000
    $receive = 10000

    # Honor DolinTools desktop settings when present (manifest provider and timeouts).
    $dolinSettingsPath = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'DolinTools\settings.json'
    if (Test-Path -LiteralPath $dolinSettingsPath) {
        try {
            $settings = Get-Content -LiteralPath $dolinSettingsPath -Raw | ConvertFrom-Json
            if ($settings.OpenSteamManifestProvider -in @('opensteamtool', 'steamrun', 'wudrm')) {
                $provider = [string]$settings.OpenSteamManifestProvider
            }
            if ($settings.OpenSteamResolveTimeoutMs) { $resolve = [int]$settings.OpenSteamResolveTimeoutMs }
            if ($settings.OpenSteamConnectTimeoutMs) { $connect = [int]$settings.OpenSteamConnectTimeoutMs }
            if ($settings.OpenSteamSendTimeoutMs) { $send = [int]$settings.OpenSteamSendTimeoutMs }
            if ($settings.OpenSteamReceiveTimeoutMs) { $receive = [int]$settings.OpenSteamReceiveTimeoutMs }
        } catch {}
    }

    $resolve = [Math]::Min([Math]::Max($resolve, 1000), 60000)
    $connect = [Math]::Min([Math]::Max($connect, 1000), 60000)
    $send = [Math]::Min([Math]::Max($send, 1000), 60000)
    $receive = [Math]::Min([Math]::Max($receive, 1000), 60000)

    return "[log]`r`nlevel = `"info`"`r`n`r`n[manifest]`r`nurl = `"$provider`"`r`ntimeout_resolve_ms = $resolve`r`ntimeout_connect_ms = $connect`r`ntimeout_send_ms = $send`r`ntimeout_recv_ms = $receive`r`n"
}

function Test-OpenSteamTool {
    param([string]$SteamPath)

    foreach ($file in $OpenSteamToolFiles) {
        if (-not (Test-Path -LiteralPath (Join-Path $SteamPath $file))) { return $false }
    }

    # SkyTools leaves a marker file; OpenSteamTools installs must not have it.
    if (Test-Path -LiteralPath (Join-Path $SteamPath $SkyToolsMarkerFile)) { return $false }

    return $true
}

function Get-OpenSteamToolVersionLabel {
    param([string]$SteamPath)
    return Get-DllBuildLabel (Join-Path $SteamPath 'OpenSteamTool.dll')
}

function Get-InstalledIntegrationType {
    param([string]$SteamPath)

    if (Test-OpenSteamTool $SteamPath) { return 'OpenSteamTools' }
    if (Test-SkyTools $SteamPath) { return 'SkyTools' }
    if (Test-Steamtools $SteamPath) { return 'SteamTools' }
    return $null
}

function Get-LatestSkyToolsVersionLabel {
    $tempDir = Join-Path ([IO.Path]::GetTempPath()) ("SteamInstall.SkyVer." + [Guid]::NewGuid().ToString('N'))
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

    try {
        $dwmapiPath = Join-Path $tempDir 'dwmapi.dll'
        $null = Save-RemoteFile -Urls @("$SkyToolsBaseUrl" + 'dwmapi.dll') -Destination $dwmapiPath -Quiet
        return Get-DllBuildLabel $dwmapiPath
    } finally {
        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-LatestSteamtoolsVersionLabel {
    $tempDir = Join-Path ([IO.Path]::GetTempPath()) ("SteamInstall.STVer." + [Guid]::NewGuid().ToString('N'))
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

    try {
        $dwmapiPath = Join-Path $tempDir 'dwmapi.dll'
        Save-SteamtoolsDll -Dll 'dwmapi' -Destination $dwmapiPath -Quiet
        return Get-DllBuildLabel $dwmapiPath
    } finally {
        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-StartupVersionChecks {
    param([string]$SteamPath)

    $checks = [System.Collections.Generic.List[object]]::new()
    $integration = Get-InstalledIntegrationType -SteamPath $SteamPath
    $millDir = Join-Path $SteamPath 'millennium'

    if ($integration -eq 'SkyTools') {
        $checks.Add([PSCustomObject]@{
            Key     = 'SkyTools'
            Name    = 'SkyTools'
            Current = Get-SkyToolsVersionLabel -SteamPath $SteamPath
            Latest  = $null
        })
    } elseif ($integration -eq 'SteamTools') {
        $checks.Add([PSCustomObject]@{
            Key     = 'SteamTools'
            Name    = 'SteamTools'
            Current = Get-SteamtoolsVersionLabel -SteamPath $SteamPath
            Latest  = $null
        })
    } elseif ($integration -eq 'OpenSteamTools') {
        $checks.Add([PSCustomObject]@{
            Key     = 'OpenSteamTools'
            Name    = 'OpenSteamTools'
            Current = Get-OpenSteamToolVersionLabel -SteamPath $SteamPath
            Latest  = $null
        })
    }

    if ($integration -in @('SkyTools', 'SteamTools')) {
        $checks.Add([PSCustomObject]@{
            Key     = 'Millennium'
            Name    = 'Millennium'
            Current = Get-MillenniumVersionLabel -SteamPath $SteamPath -MillDir $millDir
            Latest  = $null
        })

        $pluginDir = Find-PluginDir -SteamPath $SteamPath -MillDir $millDir -Name $PluginName
        if ($pluginDir) {
            $checks.Add([PSCustomObject]@{
                Key     = 'Luatools'
                Name    = $PluginName
                Current = Get-PluginVersionFromDir -PluginDir $pluginDir
                Latest  = $null
            })
        }
    }

    foreach ($check in $checks) {
        try {
            switch ($check.Key) {
                'SkyTools' {
                    $check.Latest = Get-LatestSkyToolsVersionLabel
                }
                'SteamTools' {
                    $check.Latest = Get-LatestSteamtoolsVersionLabel
                }
                'OpenSteamTools' {
                    $check.Latest = (Get-LatestOpenSteamToolAsset).Version
                }
                'Millennium' {
                    $check.Latest = Get-GithubLatestTag -Repo $MillenniumGithubRepo
                }
                'Luatools' {
                    $check.Latest = Get-GithubLatestTag -Repo $LuatoolsGithubRepo
                }
            }
        } catch {
            Write-Log "Could not check latest $($check.Name) version: $($_.Exception.Message)"
        }
    }

    return $checks
}

function Invoke-StartupUpdateCheck {
    param([switch]$Interactive)

    $steamPath = $null
    try {
        $steamPath = Get-SteamPath
    } catch {
        Write-Step WARN 'Could not locate Steam for startup update check. Skipping.'
        return
    }

    $checks = Get-StartupVersionChecks -SteamPath $steamPath
    if (-not $checks.Count) {
        Write-Step INFO 'No supported integration detected for startup update check.'
        return
    }

    Write-Host ''
    Write-Out '=== Version check ===' -ForegroundColor Cyan

    $updateCandidates = [System.Collections.Generic.List[object]]::new()

    foreach ($check in $checks) {
        $current = Format-VersionLabel $check.Current
        $latest = Format-VersionLabel $check.Latest
        Write-Step INFO "$($check.Name): current $current | latest $latest"

        if ($check.Latest -and (Test-VersionUpdateNeeded -Installed $check.Current -Latest $check.Latest)) {
            $updateCandidates.Add($check)
        }
    }

    if (-not $updateCandidates.Count) {
        Write-Step OK 'Everything is already up to date.'
        return
    }

    if (-not $Interactive) {
        Write-Step INFO 'Updates are available. Run interactive mode to choose updates.'
        return
    }

    $steamStopped = $false
    foreach ($candidate in $updateCandidates) {
        $current = Format-VersionLabel $candidate.Current
        $latest = Format-VersionLabel $candidate.Latest
        $answer = Read-ManagerInput "Update $($candidate.Name) from $current to $latest? [Y/n]" -Question
        if ($answer -and ($answer.Trim() -match '^n(o)?$')) {
            continue
        }

        if (-not $steamStopped) {
            Stop-Steam
            $steamStopped = $true
        }

        switch ($candidate.Key) {
            'SkyTools' {
                Sync-SkyTools -SteamPath $steamPath
            }
            'SteamTools' {
                $steamtoolsUpdated = Sync-Steamtools -SteamPath $steamPath
                Sync-STFixerSupport -SteamPath $steamPath -Reapply:$steamtoolsUpdated -SkipUserPrompt
            }
            'OpenSteamTools' {
                Sync-OpenSteamTool -SteamPath $steamPath
            }
            'Millennium' {
                Install-Millennium -SteamPath $steamPath -TargetVersion $candidate.Latest
            }
            'Luatools' {
                $millDir = Join-Path $steamPath 'millennium'
                Install-Luatools -SteamPath $steamPath -MillDir $millDir -Name $PluginName -TargetVersion $candidate.Latest
            }
        }
    }
}

function Get-LatestOpenSteamToolAsset {
    # Resolves browser_download_url from the latest GitHub release (*-Release.zip asset only).
    $release = Invoke-RestMethod `
        -Uri $OpenSteamToolLatestReleaseUrl `
        -Headers @{ 'User-Agent' = 'SteamTools-Install-Manager-Script' }

    $asset = $release.assets | Where-Object { $_.name -like '*-Release.zip' } | Select-Object -First 1
    if (-not $asset) {
        throw 'The latest OpenSteamTools release does not contain a Release zip.'
    }

    return [PSCustomObject]@{
        Version = $release.tag_name
        Name    = $asset.name
        Url     = $asset.browser_download_url
    }
}

function Sync-OpenSteamTool {
    param(
        [string]$SteamPath,
        [switch]$SkipDefenderConsentPrompt
    )

    $tempDir = Join-Path ([IO.Path]::GetTempPath()) ("SteamInstall.OST." + [Guid]::NewGuid().ToString('N'))
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
    $defenderExclusionPaths = $null   # Always cleared in finally — exclusions are temporary only

    try {
        $defenderExclusionPaths = Enable-OpenSteamToolDefenderExclusions -TempDir $tempDir -DestinationPath $SteamPath -SkipConsentPrompt:$SkipDefenderConsentPrompt

        $release = Get-LatestOpenSteamToolAsset
        Write-Step INFO "Downloading OpenSteamTools $(Format-VersionLabel $release.Version)..."

        $archivePath = Join-Path $tempDir $release.Name
        try {
            $null = Save-RemoteFile -Urls @($release.Url) -Destination $archivePath -Quiet
        } catch {
            if (Test-IsDefenderBlock $_.Exception.Message) {
                throw 'Windows Defender blocked the OpenSteamTools download even with a temporary exclusion. Allow the file in Windows Security, then retry.'
            }
            throw
        }

        Unblock-File -LiteralPath $archivePath -ErrorAction SilentlyContinue

        $extractPath = Join-Path $tempDir 'OpenSteamTool'
        New-Item -Path $extractPath -ItemType Directory -Force | Out-Null
        try {
            Expand-ZipEntry -ZipPath $archivePath -DestinationRoot $extractPath -UnblockExtracted
        } catch {
            Write-Log "OpenSteamTools extract failed: $($_.Exception.Message)"
            if (Test-IsDefenderBlock $_.Exception.Message) {
                throw 'Windows Defender blocked extracting OpenSteamTools even with a temporary exclusion. Allow the file in Windows Security, then retry.'
            }
            throw
        }

        $items = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($fileName in $OpenSteamToolFiles) {
            $file = Get-ChildItem -LiteralPath $extractPath -Filter $fileName -Recurse -File |
                Select-Object -First 1
            if (-not $file) {
                throw "Release $($release.Version) does not contain $fileName."
            }

            $destPath = Join-Path $SteamPath $fileName
            $items.Add(@{
                Temp       = $file.FullName
                Dest       = $destPath
                Label      = $fileName
                RemoteHash = Get-FileSha256Hex $file.FullName
                LocalHash  = if (Test-Path -LiteralPath $destPath) { Get-FileSha256Hex $destPath } else { $null }
            })
        }

        $configTemp = Join-Path $tempDir $SkyToolsConfigFile
        $configDest = Join-Path $SteamPath $SkyToolsConfigFile
        Set-Content -LiteralPath $configTemp -Value (Get-OpenSteamToolConfig) -Encoding UTF8
        $configHash = Get-FileSha256Hex $configTemp
        $localConfigHash = if (Test-Path -LiteralPath $configDest) {
            Get-FileSha256Hex $configDest
        } else {
            $null
        }

        # Presence of the SkyTools marker means we are switching integrations — force refresh.
        $needsUpdate = ($items | Where-Object { $_.LocalHash -ne $_.RemoteHash }).Count -gt 0 `
            -or $configHash -ne $localConfigHash `
            -or (Test-Path -LiteralPath (Join-Path $SteamPath $SkyToolsMarkerFile))

        if (-not $needsUpdate -and (Test-OpenSteamTool $SteamPath)) {
            $version = Format-VersionLabel (Get-OpenSteamToolVersionLabel $SteamPath)
            Write-Step OK "OpenSteamTools already up to date ($version)"
            return
        }

        $version = Format-VersionLabel $release.Version
        Write-Step INFO "Installing OpenSteamTools ($version)..."

        Remove-Item -LiteralPath (Join-Path $SteamPath $SkyToolsMarkerFile) -Force -ErrorAction SilentlyContinue
        Move-LuaScriptsToLua -SteamPath $SteamPath

        foreach ($item in $items) {
            try {
                Copy-Item -LiteralPath $item.Temp -Destination $item.Dest -Force
            } catch {
                if (Test-IsDefenderBlock $_.Exception.Message) {
                    throw 'Windows Defender blocked copying OpenSteamTools into the Steam folder even with a temporary exclusion.'
                }
                throw
            }
        }

        try {
            Copy-Item -LiteralPath $configTemp -Destination $configDest -Force
        } catch {
            if (Test-IsDefenderBlock $_.Exception.Message) {
                throw 'Windows Defender blocked copying opensteamtool.toml into the Steam folder even with a temporary exclusion.'
            }
            throw
        }

        if (-not (Test-OpenSteamTool $SteamPath)) {
            throw 'OpenSteamTools install failed - expected files were not found after install.'
        }

        foreach ($item in $items) {
            $installedHash = Get-FileSha256Hex $item.Dest
            if ($installedHash -ne $item.RemoteHash) {
                throw "OpenSteamTools install failed - $($item.Label) hash mismatch after install."
            }
        }

        Write-Step OK "OpenSteamTools installed ($version)"
    } finally {
        Disable-OpenSteamToolDefenderExclusions -Exclusions $defenderExclusionPaths
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# =============================================================================
# Install
# =============================================================================

# --- Millennium and Luatools (bundled with SkyTools / SteamTools installs) ---

function Get-MillenniumVersionLabel {
    param(
        [string]$SteamPath,
        [string]$MillDir
    )

    foreach ($candidate in @(
        (Join-Path $MillDir 'version'),
        (Join-Path $MillDir 'VERSION'),
        (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Millennium\version')
    )) {
        if (-not (Test-Path -LiteralPath $candidate)) { continue }
        $text = (Get-Content -LiteralPath $candidate -Raw -ErrorAction SilentlyContinue).Trim()
        if ($text) { return $text }
    }

    foreach ($dll in @('wsock32.dll', 'millennium.dll')) {
        $dllPath = Join-Path $SteamPath $dll
        if (-not (Test-Path -LiteralPath $dllPath)) { continue }

        $versionInfo = (Get-Item -LiteralPath $dllPath).VersionInfo
        if ($versionInfo.ProductVersion -and $versionInfo.ProductVersion -notmatch '^0\.0\.0') {
            return $versionInfo.ProductVersion
        }
        if ($versionInfo.FileVersion -and $versionInfo.FileVersion -notmatch '^0\.0\.0') {
            return $versionInfo.FileVersion
        }
    }

    return $null
}

# --- SteamTools DLL sync and HKCU registry (InstallSteamTools path) ---

function Test-Steamtools {
    param([string]$SteamPath)
    # Classic SteamTools integration: only dwmapi + xinput hijacks — no OpenSteamTool.dll.
    foreach ($file in @('dwmapi.dll', 'xinput1_4.dll')) {
        if (-not (Test-Path (Join-Path $SteamPath $file))) { return $false }
    }
    return $true
}

function Get-SteamtoolsVersionLabel {
    param([string]$SteamPath)
    return Get-DllBuildLabel (Join-Path $SteamPath 'dwmapi.dll')
}

function Initialize-SteamtoolsRegistry {
    $regPath = 'HKCU:\Software\Valve\Steamtools'
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    # Drop legacy unlock-related values; set iscdkey=false expected by current SteamTools builds.
    foreach ($name in @('ActivateUnlockMode', 'AlwaysStayUnlocked', 'notUnlockDepot')) {
        Remove-ItemProperty -Path $regPath -Name $name -ErrorAction SilentlyContinue
    }

    Set-ItemProperty -Path $regPath -Name 'iscdkey' -Value 'false' -Type String
}

function Sync-Steamtools {
    param([string]$SteamPath)

    $xinputDest = Join-Path $SteamPath 'xinput1_4.dll'
    $dwmapiDest = Join-Path $SteamPath 'dwmapi.dll'
    $tempDir    = Join-Path $SteamPath 'steamtools-sync-temp'

    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

    try {
        $xinputTemp = Join-Path $tempDir 'xinput1_4.dll'
        $dwmapiTemp = Join-Path $tempDir 'dwmapi.dll'

        Write-Step INFO 'Downloading SteamTools...'
        Invoke-SteamtoolsDllDownloads -XinputDestination $xinputTemp -DwmapiDestination $dwmapiTemp -Quiet

        $items = @(
            @{ Temp = $xinputTemp; Dest = $xinputDest; Label = 'xinput1_4.dll' },
            @{ Temp = $dwmapiTemp; Dest = $dwmapiDest; Label = 'dwmapi.dll' }
        )

        foreach ($item in $items) {
            $item.RemoteHash = Get-FileSha256Hex $item.Temp
            $item.LocalHash  = if (Test-Path -LiteralPath $item.Dest) { Get-FileSha256Hex $item.Dest } else { $null }
            $item.NeedsCopy  = $item.LocalHash -ne $item.RemoteHash
        }

        if (-not ($items | Where-Object { $_.NeedsCopy })) {
            if (-not (Test-Steamtools $SteamPath)) {
                throw 'SteamTools check failed - local DLLs are missing.'
            }

            Initialize-SteamtoolsRegistry
            $version = Format-VersionLabel (Get-SteamtoolsVersionLabel $SteamPath)
            Write-Step OK "SteamTools already up to date ($version)"
            return $false
        }

        $version = Format-VersionLabel (Get-DllBuildLabel $dwmapiTemp)
        Write-Step INFO "Installing SteamTools ($version)..."

        foreach ($item in ($items | Where-Object { $_.NeedsCopy })) {
            Copy-Item -LiteralPath $item.Temp -Destination $item.Dest -Force
        }

        Initialize-SteamtoolsRegistry

        if (-not (Test-Steamtools $SteamPath)) {
            throw 'SteamTools install failed - dwmapi.dll or xinput1_4.dll not found after install.'
        }

        foreach ($item in $items) {
            $installedHash = Get-FileSha256Hex $item.Dest
            if ($installedHash -ne $item.RemoteHash) {
                throw "SteamTools install failed - $($item.Label) hash mismatch after install."
            }
        }

        Write-Step OK "SteamTools installed ($version)"
        return $true
    } finally {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Test-Millennium {
    param([string]$SteamPath)
    foreach ($file in @('wsock32.dll', 'millennium.dll', 'python311.dll')) {
        if (Test-Path -LiteralPath (Join-Path $SteamPath $file)) { return $true }
    }
    return $false
}

function Get-LuatoolsDownloadUrls {
    return @($PluginUrl)
}

function Install-MillenniumFromGithubZip {
    param([string]$SteamPath)

    $asset = Get-GithubWindowsZipAsset -Repo $MillenniumGithubRepo
    $tempDir = Join-Path ([IO.Path]::GetTempPath()) ("SteamInstall.Mill." + [Guid]::NewGuid().ToString('N'))
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
    $zipPath = Join-Path $tempDir $asset.Name

    try {
        $null = Save-RemoteFile -Urls @($asset.Url) -Destination $zipPath -Quiet
        Expand-ZipEntry -ZipPath $zipPath -DestinationRoot $SteamPath
        return $asset.Version
    } finally {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Install-Millennium {
    param(
        [string]$SteamPath,
        [string]$TargetVersion
    )

    Write-Step INFO "Installing Millennium $(Format-VersionLabel $TargetVersion)..."

    $githubVersion = Install-MillenniumFromGithubZip -SteamPath $SteamPath

    $installedVersion = Get-MillenniumVersionLabel -SteamPath $SteamPath -MillDir (Join-Path $SteamPath 'millennium')
    if (-not $installedVersion) { $installedVersion = if ($githubVersion) { $githubVersion } else { $TargetVersion } }

    if (Test-Millennium $SteamPath) {
        Write-Step OK "Millennium installed ($(Format-VersionLabel $installedVersion))"
    } else {
        Write-Step WARN "Millennium installer finished ($(Format-VersionLabel $installedVersion)), but detection files were not found - it may still have worked."
    }
}

function Test-Luatools {
    param(
        [string]$SteamPath,
        [string]$MillDir,
        [string]$Name
    )
    return [bool](Find-PluginDir -SteamPath $SteamPath -MillDir $MillDir -Name $Name)
}

function Enable-Plugin {
    param(
        [string]$MillDir,
        [string]$Name
    )

    $configPath = Join-Path $MillDir 'config\config.json'
    $configDir  = Split-Path $configPath -Parent

    if (-not (Test-Path $configDir)) {
        New-Item -Path $configDir -ItemType Directory -Force | Out-Null
    }

    if (-not (Test-Path $configPath)) {
        $config = @{ plugins = @{ enabledPlugins = @($Name) } }
    } else {
        $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $config.plugins) {
            $config | Add-Member -NotePropertyName plugins -NotePropertyValue @{ enabledPlugins = @() } -Force
        }
        if (-not $config.plugins.enabledPlugins) {
            $config.plugins.enabledPlugins = @()
        }
        $enabled = @($config.plugins.enabledPlugins)
        if ($enabled -notcontains $Name) {
            $enabled += $Name
            $config.plugins.enabledPlugins = $enabled
        }
    }

    $config | ConvertTo-Json -Depth 10 | Set-Content $configPath -Encoding UTF8
}

function Install-Luatools {
    param(
        [string]$SteamPath,
        [string]$MillDir,
        [string]$Name,
        [string]$TargetVersion
    )

    Write-Step INFO "Installing $Name $(Format-VersionLabel $TargetVersion)..."

    $pluginsDir = Join-Path $MillDir 'plugins'
    $targetDir  = Join-Path $pluginsDir $Name
    $zipPath    = Join-Path $SteamPath "$Name.zip"

    if (-not (Test-Path $pluginsDir)) {
        New-Item -Path $pluginsDir -ItemType Directory -Force | Out-Null
    }

    $downloadUrls = Get-LuatoolsDownloadUrls
    $null = Save-RemoteFile -Urls $downloadUrls -Destination $zipPath -Quiet
    if (-not (Test-Path $zipPath)) { throw "Failed to download $Name." }

    $zipVersion = Get-PluginVersionFromZip -ZipPath $zipPath
    if ($zipVersion) {
        Write-Step INFO "  package version: $zipVersion"
    }

    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) { continue }

            $dest   = Join-Path $targetDir $entry.FullName
            $parent = Split-Path $dest -Parent
            if ($parent) { [void][System.IO.Directory]::CreateDirectory($parent) }
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dest, $true)
        }
        $zip.Dispose()
    } catch {
        if ($zip) { $zip.Dispose() }
        Expand-Archive -Path $zipPath -DestinationPath $targetDir -Force
    }

    Remove-Item $zipPath -Force -ErrorAction SilentlyContinue

    if (-not (Test-Luatools $SteamPath $MillDir $Name)) {
        throw "$Name installation failed - plugin folder not found after extract."
    }

    $pluginDir = Find-PluginDir -SteamPath $SteamPath -MillDir $MillDir -Name $Name
    $installedVersion = Get-PluginVersionFromDir -PluginDir $pluginDir
    if (-not $installedVersion) { $installedVersion = $zipVersion }
    if (-not $installedVersion) { $installedVersion = $TargetVersion }

    Enable-Plugin -MillDir $MillDir -Name $Name
    Write-Step OK "$Name installed and enabled ($(Format-VersionLabel $installedVersion))"
}

function Install-MillenniumAndLuatools {
    param([string]$SteamPath)

    $millDir = Join-Path $steamPath 'millennium'
    $millenniumLatest = Get-GithubLatestTag -Repo $MillenniumGithubRepo

    if (-not (Test-Millennium $steamPath)) {
        Install-Millennium -SteamPath $steamPath -TargetVersion $millenniumLatest
    } else {
        $millenniumVersion = Get-MillenniumVersionLabel -SteamPath $steamPath -MillDir $millDir
        if (Test-VersionUpdateNeeded -Installed $millenniumVersion -Latest $millenniumLatest) {
            Write-Step INFO "Millennium update available: $(Format-VersionLabel $millenniumVersion) -> $millenniumLatest"
            Install-Millennium -SteamPath $steamPath -TargetVersion $millenniumLatest
        } else {
            Write-Step OK "Millennium already up to date ($(Format-VersionLabel $millenniumVersion))"
        }
    }

    if (-not (Test-Path $millDir)) {
        New-Item -Path $millDir -ItemType Directory -Force | Out-Null
    }

    $luatoolsLatest = $null
    try { $luatoolsLatest = Get-GithubLatestTag -Repo $LuatoolsGithubRepo } catch {}

    if (-not (Test-Luatools $steamPath $millDir $PluginName)) {
        Install-Luatools -SteamPath $steamPath -MillDir $millDir -Name $PluginName -TargetVersion $luatoolsLatest
    } else {
        $pluginDir = Find-PluginDir -SteamPath $steamPath -MillDir $millDir -Name $PluginName
        $luatoolsVersion = Get-PluginVersionFromDir -PluginDir $pluginDir
        if (Test-VersionUpdateNeeded -Installed $luatoolsVersion -Latest $luatoolsLatest) {
            Write-Step INFO "Luatools update available: $(Format-VersionLabel $luatoolsVersion) -> $luatoolsLatest"
            Install-Luatools -SteamPath $steamPath -MillDir $millDir -Name $PluginName -TargetVersion $luatoolsLatest
        } else {
            Write-Step OK "Luatools already up to date ($(Format-VersionLabel $luatoolsVersion))"
        }
    }
}

# --- End-to-end install flows (stop Steam, remove old integration, sync, optional extras) ---
# Order matters: stop Steam, align Lua script folder, wipe competing DLLs, sync, then Millennium/Luatools.

function Start-SkyToolsInstallFlow {
    param([switch]$Force)

    $steamPath = Get-SteamPath

    Write-Step INFO "Steam path: $steamPath"
    Stop-Steam -Force:$Force
    Move-LuaScriptsToStPlugin -SteamPath $steamPath
    Remove-CompetingSteamIntegrations -SteamPath $steamPath -TargetIntegration SkyTools
    Sync-SkyTools -SteamPath $steamPath
    Install-MillenniumAndLuatools -SteamPath $steamPath
}

function Start-OpenSteamToolInstallFlow {
    param([switch]$Force)

    $steamPath = Get-SteamPath

    Write-Step INFO "Steam path: $steamPath"
    if (-not (Test-UserAllowsDefenderExclusion)) {
        throw 'OpenSteamTools install cancelled. A temporary Windows Defender exclusion is required.'
    }
    Stop-Steam -Force:$Force
    Move-LuaScriptsToLua -SteamPath $steamPath
    Remove-CompetingSteamIntegrations -SteamPath $steamPath -TargetIntegration OpenSteamTools
    Sync-OpenSteamTool -SteamPath $steamPath -SkipDefenderConsentPrompt
}

function Start-SteamtoolsInstallFlow {
    param([switch]$Force)

    $steamPath = Get-SteamPath

    Write-Step INFO "Steam path: $steamPath"
    Stop-Steam -Force:$Force
    Move-LuaScriptsToStPlugin -SteamPath $steamPath
    Remove-CompetingSteamIntegrations -SteamPath $steamPath -TargetIntegration SteamTools
    $steamtoolsUpdated = Sync-Steamtools -SteamPath $steamPath
    Sync-STFixerSupport -SteamPath $steamPath -Reapply:$steamtoolsUpdated -SkipUserPrompt:$Force
    Install-MillenniumAndLuatools -SteamPath $steamPath
}

function Start-ReinstallFlow {
    param(
        [ValidateSet('SkyTools', 'SteamTools', 'OpenSteamTools')]
        [string]$Integration,
        [switch]$Force
    )

    Write-Step INFO "Reinstall requested for $Integration"
    Start-UninstallFlow -Force:$Force

    switch ($Integration) {
        'SkyTools' { Start-SkyToolsInstallFlow -Force:$Force }
        'SteamTools' { Start-SteamtoolsInstallFlow -Force:$Force }
        'OpenSteamTools' { Start-OpenSteamToolInstallFlow -Force:$Force }
    }
}

# =============================================================================
# Uninstall
# =============================================================================

function Remove-SteamtoolsRegistry {
    param([switch]$Quiet)

    $regPath = 'HKCU:\Software\Valve\Steamtools'
    if (Remove-PathIfExists -Path $regPath -Recurse) {
        if ($Quiet) {
            Write-Log 'Removed SteamTools registry key'
        } else {
            Write-Step OK 'Removed SteamTools registry key'
        }
    }
}

function Start-UninstallFlow {
    param([switch]$Force)

    $steamPath = Get-SteamPath
    $millDir   = Join-Path $steamPath 'millennium'

    Write-Step INFO "Steam path: $steamPath"

    $targets = [System.Collections.Generic.List[string]]::new()

    Add-SteamIntegrationRemovalTargets -SteamPath $steamPath -Targets $targets

    foreach ($file in @('wsock32.dll', 'millennium.dll', 'python311.dll')) {
        $targets.Add((Join-Path $steamPath $file))
    }

    if (Test-Path -LiteralPath $millDir) {
        $targets.Add($millDir)
    }

    $legacyPlugin = Find-PluginDir -SteamPath $steamPath -MillDir $millDir -Name $PluginName
    # Older Luatools installs sometimes lived outside steam\millennium\plugins.
    if ($legacyPlugin -and $legacyPlugin -notlike "$millDir*") {
        $targets.Add($legacyPlugin)
    }

    $existing = $targets | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -Unique

    if (-not $existing) {
        Write-Step OK 'Nothing to uninstall.'
        Remove-SteamtoolsRegistry
        return
    }

    Write-Step WARN 'The following will be removed:'
    foreach ($path in $existing) {
        Write-Out "  - $path"
    }

    if (-not $Force) {
        Write-Host ''
        $confirmPrompt = if (Test-SteamRunning) {
            'Steam will be closed and the listed files will be removed. Continue? [Y/n]'
        } else {
            'The listed files will be removed. Continue? [Y/n]'
        }
        $answer = Read-ManagerInput $confirmPrompt -Question
        if ($answer -and ($answer.Trim() -match '^n(o)?$')) {
            Write-Step INFO 'Cancelled.'
            return
        }
    }

    Stop-Steam -Force

    $failedRemovals = [System.Collections.Generic.List[string]]::new()
    foreach ($path in $existing) {
        $isDir = Test-Path -LiteralPath $path -PathType Container
        try {
            if (Remove-PathIfExists -Path $path -Recurse:$isDir -MaxAttempts 5) {
                Write-Step OK "Removed $path"
            }
        } catch {
            $failedRemovals.Add($path)
            Write-Step WARN "Could not remove $path - $($_.Exception.Message)"
        }
    }

    Remove-SteamtoolsRegistry

    if ($failedRemovals.Count -gt 0) {
        Write-Step WARN 'Uninstall finished with errors. Close Steam from the system tray, end any steamwebhelper tasks in Task Manager, then run uninstall again.'
        throw "Could not remove: $($failedRemovals -join ', ')"
    }

    Write-Step OK 'Uninstall complete.'
}

# =============================================================================
# Export pack (offline bundle: steam\ DLLs + optional Millennium/Luatools + game scripts)
# =============================================================================

function Get-UnlockedGameScriptFilesForExport {
    param([string]$SteamPath)

    $files = [ordered]@{}
    $stplugIn = Join-Path $SteamPath 'config\stplug-in'
    $luaDir   = Join-Path $SteamPath 'config\lua'

    # Merge both script folders; same filename in stplug-in wins over config\lua.
    if (Test-Path -LiteralPath $stplugIn -PathType Container) {
        foreach ($file in Get-ChildItem -LiteralPath $stplugIn -File -ErrorAction SilentlyContinue) {
            if (-not $files.Contains($file.Name)) {
                $files[$file.Name] = $file
            }
        }
    }

    if (Test-Path -LiteralPath $luaDir -PathType Container) {
        foreach ($file in Get-ChildItem -LiteralPath $luaDir -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '\.lua(\.disabled)?$' }) {
            if (-not $files.Contains($file.Name)) {
                $files[$file.Name] = $file
            }
        }
    }

    return @($files.Values)
}

function Add-UnlockedGameScriptsToPack {
    param(
        [string]$SteamPath,
        [string]$PackRoot,
        [ValidateSet('stplug-in', 'lua')]
        [string]$ScriptFolder
    )

    $stplugIn = Join-Path $SteamPath 'config\stplug-in'
    $luaDir   = Join-Path $SteamPath 'config\lua'
    $hasSources = (Test-Path -LiteralPath $stplugIn -PathType Container) -or
        (Test-Path -LiteralPath $luaDir -PathType Container)

    if (-not $hasSources) {
        Write-Step WARN 'No unlocked game scripts found in your Steam install - skipped.'
        return
    }

    $files = Get-UnlockedGameScriptFilesForExport -SteamPath $SteamPath
    if (-not $files.Count) {
        Write-Step INFO 'No unlocked game scripts to include.'
        return
    }

    $dest = Join-Path $PackRoot "config\$ScriptFolder"
    New-Item -Path $dest -ItemType Directory -Force | Out-Null
    Copy-Item -LiteralPath ($files | ForEach-Object { $_.FullName }) -Destination $dest -Force

    Write-Step OK "Included $($files.Count) unlocked game script(s) in config\$ScriptFolder"
    foreach ($file in $files) {
        Write-Step INFO "  - $($file.Name)"
    }
}

function Add-StPluginScriptsToPack {
    param(
        [string]$SteamPath,
        [string]$PackRoot
    )

    Add-UnlockedGameScriptsToPack -SteamPath $SteamPath -PackRoot $PackRoot -ScriptFolder 'stplug-in'
}

function Add-LuaScriptsToPack {
    param(
        [string]$SteamPath,
        [string]$PackRoot
    )

    Add-UnlockedGameScriptsToPack -SteamPath $SteamPath -PackRoot $PackRoot -ScriptFolder 'lua'
}

# --- Download integration assets into a pack folder; prompts for export options ---

function Download-SkyToolsPack {
    param([string]$TargetRoot)

    $tempDir = Join-Path $TargetRoot '_sync-temp'
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

    try {
        $dwmapiTemp = $null
        foreach ($fileName in $SkyToolsFiles) {
            $tempPath = Join-Path $tempDir $fileName
            $null = Save-RemoteFile -Urls @("$SkyToolsBaseUrl$fileName") -Destination $tempPath -Quiet
            if ($fileName -eq 'dwmapi.dll') { $dwmapiTemp = $tempPath }
        }

        $label = Format-VersionLabel (Get-DllBuildLabel $dwmapiTemp)
        Write-Step INFO "Downloading SkyTools ($label)..."

        foreach ($fileName in $SkyToolsFiles) {
            $tempPath = Join-Path $tempDir $fileName
            Copy-Item -LiteralPath $tempPath -Destination (Join-Path $TargetRoot $fileName) -Force
        }

        Set-Content -LiteralPath (Join-Path $TargetRoot $SkyToolsConfigFile) -Value (Get-SkyToolsConfig) -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $TargetRoot $SkyToolsMarkerFile) -Value 'SkyTools offline pack' -Encoding ASCII

        Write-Step OK "SkyTools downloaded ($label)"
    } finally {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Download-SteamtoolsPack {
    param([string]$TargetRoot)

    $tempDir = Join-Path $TargetRoot '_sync-temp'
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

    try {
        $xinputTemp = Join-Path $tempDir 'xinput1_4.dll'
        $dwmapiTemp = Join-Path $tempDir 'dwmapi.dll'

        Write-Step INFO 'Downloading SteamTools...'
        Invoke-SteamtoolsDllDownloads -XinputDestination $xinputTemp -DwmapiDestination $dwmapiTemp -Quiet

        foreach ($item in @(
            @{ Temp = $xinputTemp; Dest = (Join-Path $TargetRoot 'xinput1_4.dll') },
            @{ Temp = $dwmapiTemp; Dest = (Join-Path $TargetRoot 'dwmapi.dll') }
        )) {
            Copy-Item -LiteralPath $item.Temp -Destination $item.Dest -Force
        }

        $label = Format-VersionLabel (Get-DllBuildLabel $dwmapiTemp)
        Write-Step OK "SteamTools downloaded ($label)"
    } finally {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Download-OpenSteamToolPack {
    param(
        [string]$TargetRoot,
        [string]$CacheRoot,
        [switch]$SkipDefenderConsentPrompt
    )

    $tempDir = Join-Path $CacheRoot 'ost-extract'
    $packRoot = [IO.Path]::GetFullPath((Split-Path $CacheRoot -Parent))
    $defenderExclusionPaths = $null

    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

    try {
        $defenderExclusionPaths = Enable-OpenSteamToolDefenderExclusions `
            -TempDir $packRoot `
            -DestinationPath $packRoot `
            -ExtraPaths @($CacheRoot, $TargetRoot, $tempDir) `
            -CancelledMessage 'OpenSteamTools export cancelled. A temporary Windows Defender exclusion is required.' `
            -SkipConsentPrompt:$SkipDefenderConsentPrompt

        $release = Get-LatestOpenSteamToolAsset
        $version = Format-VersionLabel $release.Version
        Write-Step INFO "Downloading OpenSteamTools ($version)..."

        $archivePath = Join-Path $CacheRoot $release.Name
        try {
            $null = Save-RemoteFile -Urls @($release.Url) -Destination $archivePath -Quiet
        } catch {
            if (Test-IsDefenderBlock $_.Exception.Message) {
                throw 'Windows Defender blocked the OpenSteamTools download even with a temporary exclusion. Allow the file in Windows Security, then retry.'
            }
            throw
        }

        Unblock-File -LiteralPath $archivePath -ErrorAction SilentlyContinue

        $extractPath = Join-Path $tempDir 'OpenSteamTool'
        New-Item -Path $extractPath -ItemType Directory -Force | Out-Null
        try {
            Expand-ZipEntry -ZipPath $archivePath -DestinationRoot $extractPath -UnblockExtracted
        } catch {
            Write-Log "OpenSteamTools export extract failed: $($_.Exception.Message)"
            if (Test-IsDefenderBlock $_.Exception.Message) {
                throw 'Windows Defender blocked extracting OpenSteamTools even with a temporary exclusion. Allow the file in Windows Security, then retry.'
            }
            throw
        }

        foreach ($fileName in $OpenSteamToolFiles) {
            $file = Get-ChildItem -LiteralPath $extractPath -Filter $fileName -Recurse -File |
                Select-Object -First 1
            if (-not $file) {
                throw "Release $($release.Version) does not contain $fileName."
            }

            try {
                Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $TargetRoot $fileName) -Force
            } catch {
                if (Test-IsDefenderBlock $_.Exception.Message) {
                    throw 'Windows Defender blocked copying OpenSteamTools into the pack folder even with a temporary exclusion.'
                }
                throw
            }
        }

        try {
            Set-Content -LiteralPath (Join-Path $TargetRoot $SkyToolsConfigFile) -Value (Get-OpenSteamToolConfig) -Encoding UTF8
        } catch {
            if (Test-IsDefenderBlock $_.Exception.Message) {
                throw 'Windows Defender blocked copying opensteamtool.toml into the pack folder even with a temporary exclusion.'
            }
            throw
        }

        Write-Step OK "OpenSteamTools downloaded ($version)"
    } finally {
        Disable-OpenSteamToolDefenderExclusions -Exclusions $defenderExclusionPaths
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Download-MillenniumPack {
    param(
        [string]$TargetRoot,
        [string]$CacheRoot
    )

    # GitHub release asset name must match *windows* and *zip*.
    $asset = Get-GithubWindowsZipAsset -Repo $MillenniumGithubRepo
    $zipPath = Join-Path $CacheRoot $asset.Name
    $version = Format-VersionLabel $asset.Version

    Write-Step INFO "Downloading Millennium ($version)..."
    if (-not (Test-Path $CacheRoot)) {
        New-Item -Path $CacheRoot -ItemType Directory -Force | Out-Null
    }

    if (-not (Test-Path $zipPath) -or (Get-Item $zipPath).Length -eq 0) {
        $null = Save-RemoteFile -Urls @($asset.Url) -Destination $zipPath -Quiet
    }
    Expand-ZipEntry -ZipPath $zipPath -DestinationRoot $TargetRoot

    Write-Step OK "Millennium downloaded ($version)"
}

function Download-LuatoolsPack {
    param(
        [string]$TargetRoot,
        [string]$CacheRoot,
        [string]$Name
    )

    $zipPath   = Join-Path $CacheRoot "$Name.zip"
    $pluginDir = Join-Path $TargetRoot "millennium\plugins\$Name"

    Write-Step INFO "Downloading $Name..."
    if (-not (Test-Path $CacheRoot)) {
        New-Item -Path $CacheRoot -ItemType Directory -Force | Out-Null
    }
    if (-not (Test-Path $zipPath) -or (Get-Item $zipPath).Length -eq 0) {
        $null = Save-RemoteFile -Urls (Get-LuatoolsDownloadUrls) -Destination $zipPath -Quiet
    }

    $version = Format-VersionLabel (Get-PluginVersionFromZip -ZipPath $zipPath)

    if (-not (Test-Path $pluginDir)) {
        New-Item -Path $pluginDir -ItemType Directory -Force | Out-Null
    }

    Expand-ZipEntry -ZipPath $zipPath -DestinationRoot $pluginDir

    $configDir  = Join-Path $TargetRoot 'millennium\config'
    $configPath = Join-Path $configDir 'config.json'
    if (-not (Test-Path $configDir)) {
        New-Item -Path $configDir -ItemType Directory -Force | Out-Null
    }

    $config = @{ plugins = @{ enabledPlugins = @($Name) } }
    $config | ConvertTo-Json -Depth 10 | Set-Content $configPath -Encoding UTF8

    Write-Step OK "$Name downloaded ($version)"
}

function New-OfflinePackZip {
    param(
        [string]$SourceDir,
        [string]$ZipPath
    )

    if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
    Compress-Archive -Path $SourceDir -DestinationPath $ZipPath -Force
}

function Read-ExportOutputDir {
    param([string]$DefaultDir)

    Write-Host ''
    Write-Out "  Default: $DefaultDir" -ForegroundColor DarkGray -NoLog
    $answer = Read-ManagerInput 'Output folder (press Enter for default)' -Question -NoLog
    $trimmed = $answer.Trim()
    if (-not $trimmed) { return $DefaultDir }
    # User may answer the next yes/no prompt here by mistake — treat y/n as "use default".
    if ($trimmed -match '^(y|yes|n|no)$') { return $DefaultDir }
    return $trimmed
}

function Read-ExportIncludeStPlugin {
    Write-Host ''
    $answer = Read-ManagerInput 'Include unlocked games from this Steam install? [y/N]' -Question -NoLog
    return [bool]($answer -and ($answer.Trim() -match '^y(es)?$'))
}

function Read-ExportCreateZip {
    Write-Host ''
    $answer = Read-ManagerInput 'Create zip? [y/N]' -Question -NoLog
    return [bool]($answer -and ($answer.Trim() -match '^y(es)?$'))
}

function Resolve-ExportIntegration {
    param(
        [string]$Integration,
        [switch]$Prompt
    )

    if ($Integration -in 'SkyTools', 'SteamTools', 'OpenSteamTools') {
        return $Integration
    }

    if ($Prompt) {
        switch (Read-ManagerMenuKey -Prompt 'Select export option [1-3]' -ValidChoices @('0', '1', '2', '3') -InvalidChoiceMessage 'Invalid choice. Enter 1, 2, or 3.' -NoLog) {
            '0' { throw [MainMenuReturnException]::new() }
            '1' { return 'SkyTools' }
            '2' { return 'SteamTools' }
            '3' { return 'OpenSteamTools' }
            default { throw 'Invalid choice. Enter 1, 2, or 3.' }
        }
    }

    throw 'Export pack requires -Integration SkyTools, SteamTools, or OpenSteamTools.'
}

function Start-ExportPackFlow {
    param(
        [string]$OutputDir,
        [string]$Integration,
        [switch]$PromptForStPlugin,
        [switch]$PromptForZip,
        [switch]$PromptForIntegration,
        [switch]$PromptForOutputDir,
        [switch]$IncludeStPlugin
    )

    if (-not $OutputDir) {
        $OutputDir = Join-Path (Get-ManagerDataRoot) 'steam-pack'
    }

    if ($PromptForIntegration) {
        Write-ExportSubMenuHeader
    }

    $integrationType = Resolve-ExportIntegration -Integration $Integration -Prompt:$PromptForIntegration

    if ($PromptForOutputDir) {
        $OutputDir = Read-ExportOutputDir -DefaultDir $OutputDir
    }

    $shouldIncludeStPlugin = $false
    if ($PSBoundParameters.ContainsKey('IncludeStPlugin')) {
        $shouldIncludeStPlugin = [bool]$IncludeStPlugin
    } elseif ($PromptForStPlugin) {
        $shouldIncludeStPlugin = Read-ExportIncludeStPlugin
    }

    $shouldCreateZip = $false
    if ($PromptForZip) {
        $shouldCreateZip = Read-ExportCreateZip
    }

    if ($integrationType -eq 'SkyTools') {
        $OutputDir = Join-Path $OutputDir 'skytools-pack'
    } elseif ($integrationType -eq 'SteamTools') {
        $OutputDir = Join-Path $OutputDir 'steamtools-pack'
    } else {
        $OutputDir = Join-Path $OutputDir 'opensteamtool-pack'
    }

    $steamRoot = Join-Path $OutputDir 'steam'   # Pack contents mirror a Steam folder layout
    $cacheDir  = Join-Path $OutputDir '_downloads'

    if ($integrationType -eq 'OpenSteamTools') {
        if (-not (Test-UserAllowsDefenderExclusion)) {
            throw 'OpenSteamTools export cancelled. A temporary Windows Defender exclusion is required.'
        }
    }

    Write-Host ''
    Write-Step INFO "Pack output: $OutputDir"

    try {
        New-Item -Path $steamRoot -ItemType Directory -Force | Out-Null
        New-Item -Path $cacheDir -ItemType Directory -Force | Out-Null

        if ($integrationType -eq 'SkyTools') {
            Download-SkyToolsPack -TargetRoot $steamRoot
        } elseif ($integrationType -eq 'SteamTools') {
            Download-SteamtoolsPack -TargetRoot $steamRoot
        } else {
            Download-OpenSteamToolPack -TargetRoot $steamRoot -CacheRoot $cacheDir -SkipDefenderConsentPrompt
        }

        # OpenSteamTools packs are DLL-only; SkyTools/SteamTools packs also bundle Millennium + Luatools.
        if ($integrationType -ne 'OpenSteamTools') {
            $null = Download-MillenniumPack -TargetRoot $steamRoot -CacheRoot $cacheDir
            $null = Download-LuatoolsPack -TargetRoot $steamRoot -CacheRoot $cacheDir -Name $PluginName
        }

        if ($shouldIncludeStPlugin) {
            try {
                $installedSteamPath = Get-SteamPath
                if ($integrationType -eq 'OpenSteamTools') {
                    Add-LuaScriptsToPack -SteamPath $installedSteamPath -PackRoot $steamRoot
                } else {
                    Add-StPluginScriptsToPack -SteamPath $installedSteamPath -PackRoot $steamRoot
                }
            } catch {
                Write-Step WARN $_.Exception.Message
            }
        }

        Write-Step OK 'Export pack ready.'
        Write-Step INFO "Pack folder: $steamRoot"

        if ($shouldCreateZip) {
            $zipName = switch ($integrationType) {
                'SkyTools' { 'SkyTools-Pack.zip' }
                'SteamTools' { 'SteamTools-Pack.zip' }
                'OpenSteamTools' { 'OpenSteamTools-Pack.zip' }
            }
            $zipPath = Join-Path $OutputDir $zipName
            Write-Step INFO 'Creating zip...'
            New-OfflinePackZip -SourceDir $steamRoot -ZipPath $zipPath
            Write-Step OK "Zip created: $zipPath"
        }
    } finally {
        if (Test-Path $cacheDir) {
            Remove-Item $cacheDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# =============================================================================
# Menu (interactive mode when -Action is omitted)
# =============================================================================

function Format-MenuOptionLine {
    param(
        [string]$Number,
        [string]$Label,
        [string]$Description,
        [int]$LabelWidth
    )

    $paddedLabel = $Label.PadRight($LabelWidth)
    if ($Description) {
        return "  $Number. $paddedLabel   - $Description"
    }
    return "  $Number. $paddedLabel"
}

function Write-MenuHeader {
    Write-Host ''
    Write-Out "=== $script:ManagerTitle ===" -ForegroundColor Cyan -NoLog
    Write-Host ''
    Write-Out (Format-MenuOptionLine '1' 'Install/Update' 'SkyTools, SteamTools, or OpenSteamTools' 16) -NoLog
    Write-Out (Format-MenuOptionLine '2' 'Clean Reinstall' 'Uninstall then install selected integration' 16) -NoLog
    Write-Out (Format-MenuOptionLine '3' 'Uninstall' 'Remove installed components' 16) -NoLog
    Write-Out (Format-MenuOptionLine '4' 'Export pack' 'Build offline bundle for manual install' 16) -NoLog
    Write-Out (Format-MenuOptionLine '5' 'Launch Steam' 'Open the Steam client' 16) -NoLog
    Write-Out (Format-MenuOptionLine '0' 'Exit' 'Close this window' 16) -ForegroundColor DarkGray -NoLog
    Write-Host ''
}

function Write-InstallSubMenuHeader {
    Write-Host ''
    Write-Out '=== Install / Update ===' -ForegroundColor Cyan -NoLog
    Write-Host ''
    Write-Out (Format-MenuOptionLine '1' 'SkyTools' 'SkyTools, Millennium, and Luatools (install or update)' 18) -NoLog
    Write-Out (Format-MenuOptionLine '2' 'SteamTools' 'SteamTools, Millennium, and Luatools (optional STFixer)' 18) -NoLog
    Write-Out (Format-MenuOptionLine '3' 'OpenSteamTools' 'Defender exclusion and cleanup prompt' 18) -NoLog
    Write-Out (Format-MenuOptionLine '0' 'Back' 'Return to main menu' 18) -ForegroundColor DarkGray -NoLog
    Write-Host ''
}

function Write-ExportSubMenuHeader {
    Write-Host ''
    Write-Out '=== Export pack ===' -ForegroundColor Cyan -NoLog
    Write-Host ''
    Write-Out (Format-MenuOptionLine '1' 'SkyTools' 'SkyTools, Millennium, and Luatools' 18) -NoLog
    Write-Out (Format-MenuOptionLine '2' 'SteamTools' 'SteamTools, Millennium, and Luatools (optional STFixer)' 18) -NoLog
    Write-Out (Format-MenuOptionLine '3' 'OpenSteamTools' 'Defender exclusion and cleanup prompt' 18) -NoLog
    Write-Out (Format-MenuOptionLine '0' 'Back' 'Return to main menu' 18) -ForegroundColor DarkGray -NoLog
    Write-Host ''
}

function Write-ReinstallSubMenuHeader {
    Write-Host ''
    Write-Out '=== Clean Reinstall ===' -ForegroundColor Cyan -NoLog
    Write-Host ''
    Write-Out (Format-MenuOptionLine '1' 'SkyTools' 'Reinstall SkyTools stack' 18) -NoLog
    Write-Out (Format-MenuOptionLine '2' 'SteamTools' 'Reinstall SteamTools stack' 18) -NoLog
    Write-Out (Format-MenuOptionLine '3' 'OpenSteamTools' 'Reinstall OpenSteamTools' 18) -NoLog
    Write-Out (Format-MenuOptionLine '0' 'Back' 'Return to main menu' 18) -ForegroundColor DarkGray -NoLog
    Write-Host ''
}

function Read-InstallSubMenuChoice {
    $choice = Read-ManagerMenuKey -Prompt 'Select install option [1-3]' -ValidChoices @('0', '1', '2', '3') -InvalidChoiceMessage 'Invalid choice. Enter 1, 2, or 3.' -NoLog
    if (-not $choice) {
        Clear-Host
        return $null
    }
    switch ($choice) {
        '0' {
            Clear-Host
            return 'Back'
        }
        '1' { return 'InstallSkyTools' }
        '2' { return 'InstallSteamTools' }
        '3' { return 'InstallOpenSteamTools' }
    }
}

function Read-ExportSubMenuChoice {
    $choice = Read-ManagerMenuKey -Prompt 'Select export option [1-3]' -ValidChoices @('0', '1', '2', '3') -InvalidChoiceMessage 'Invalid choice. Enter 1, 2, or 3.' -NoLog
    if (-not $choice) {
        Clear-Host
        return $null
    }
    switch ($choice) {
        '0' {
            Clear-Host
            return 'Back'
        }
        '1' { return 'SkyTools' }
        '2' { return 'SteamTools' }
        '3' { return 'OpenSteamTools' }
    }
}

function Read-ReinstallSubMenuChoice {
    $choice = Read-ManagerMenuKey -Prompt 'Select reinstall option [1-3]' -ValidChoices @('0', '1', '2', '3') -InvalidChoiceMessage 'Invalid choice. Enter 1, 2, or 3.' -NoLog
    if (-not $choice) {
        Clear-Host
        return $null
    }
    switch ($choice) {
        '0' {
            Clear-Host
            return 'Back'
        }
        '1' { return 'ReinstallSkyTools' }
        '2' { return 'ReinstallSteamTools' }
        '3' { return 'ReinstallOpenSteamTools' }
    }
}

function Read-MenuChoice {
    $choice = Read-ManagerMenuKey -Prompt 'Select an option [1-5]' -ValidChoices @('0', '1', '2', '3', '4', '5') -InvalidChoiceMessage 'Invalid choice. Enter 1, 2, 3, 4, or 5.' -NoLog
    if (-not $choice) {
        Clear-Host
        return $null
    }
    switch ($choice) {
        '0' { return 'Exit' }
        '1' {
            while ($true) {
                Write-InstallSubMenuHeader
                $subChoice = Read-InstallSubMenuChoice
                if ($subChoice -eq 'Back') { return $null }
                if ($subChoice) { return $subChoice }
            }
        }
        '2' {
            while ($true) {
                Write-ReinstallSubMenuHeader
                $subChoice = Read-ReinstallSubMenuChoice
                if ($subChoice -eq 'Back') { return $null }
                if ($subChoice) { return $subChoice }
            }
        }
        '3' { return 'Uninstall' }
        '4' {
            while ($true) {
                Write-ExportSubMenuHeader
                $subChoice = Read-ExportSubMenuChoice
                if ($subChoice -eq 'Back') { return $null }
                if ($subChoice) {
                    $script:MenuIntegration = $subChoice
                    return 'ExportPack'
                }
            }
        }
        '5' { return 'LaunchSteam' }
    }
}

function Get-ExitArtLines {
    return @(
        '                   ...'
        '            ^jp%$$$$$$$$Bdx"'
        '         ,W$$$$$$$$$$$$$$$$$$&;'
        '       [%$$$$$$$$$$$$$$$$$$$$$$B1'
        '     ^M$$$$$J$$$$$$$$$$$&bdW@$$$$&"'
        '    {@$$$$$$$$$$$$$$@U        nB$$@('
        '   [@$$$$$$$$$$$$$$$" ,dWzvMhI ^M$$$|'
        '  ,$$$$$$$$$$$$$$$@l <B"     *} <$$$$>'
        '  h$$$$$$$$$$$$$$$8  Qc      nQ ^$$$$a'
        ' ;$$$$$$$$$$$$$$$a   ;B!     W? +$$$$$i'
        ' `CM@$$$$$$$$$$$C     .m@ak@b" ,%$$$$$['
        '     '']b$$$$$$$(              C@$$$$$$['
        '          ^d^ ''          jB@$$$$$$$$$$i'
        '              "x8{    +a$$$$$$$$M$$$$o'
        '  :%w,          ^@: CB$$$$$$$$$$$$$$$<'
        '   {@$$@k<      ;B^w$$$$$$$$$$$$$$$$/'
        '    {@$$$%I-#LXbMl{$$$$$$$$$$$$$$$@('
        '     "&$$$$%?"^,fB$$$$$$B$$$$$$$$8"'
        '       {B$$$$$$$$$$$$$$$$$$$$$$B|'
        '         ;W$$$$$$$$$$$$$$$$$$&!'
        '            "nbB$$$$$$$$$kv,'
        '                   ...'
    )
}

function Write-ExitArt {
    $lines = @(Get-ExitArtLines | Where-Object { $null -ne $_ })
    if (-not $lines.Count) { return $null }

    $artWidth = ($lines | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum

    Write-Host ''
    foreach ($line in $lines) {
        Write-Host $line -ForegroundColor Gray
    }
    Write-Host ''

    return @{ ArtWidth = $artWidth }
}

function Invoke-ManagerExit {
    if ([Console]::Out) {
        try { [Console]::Clear() } catch { Clear-Host }
    } else {
        Clear-Host
    }

    $layout = Write-ExitArt

    $message = 'Have fun!!'
    if ($layout) {
        $messagePad = [Math]::Max(0, [int](($layout.ArtWidth - $message.Length) / 2))
        Write-Host (' ' * $messagePad + $message) -ForegroundColor Green
    } else {
        Write-Host $message -ForegroundColor Green
    }
    Write-Host ''
    [Console]::Out.Flush()
}

function Wait-MenuReturn {
    Write-Host ''
    Write-Host 'Press any key to return to menu...' -ForegroundColor DarkGray
    if ($Host.UI -and $Host.UI.RawUI) {
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    } else {
        Read-Host | Out-Null
    }
    Clear-Host
}

function Test-LaunchSteamAfterInstall {
    $answer = Read-ManagerInput 'Launch Steam and exit? [Y/n]' -Question -NoLog
    if ($answer -and ($answer.Trim() -match '^n(o)?$')) {
        return $false
    }
    return $true
}

function Invoke-PostInstallSteamPrompt {
    param([switch]$FromMenu)

    if (-not (Test-LaunchSteamAfterInstall)) {
        if ($FromMenu) {
            Clear-Host
        }
        return 'Menu'
    }

    $steamPath = Get-SteamPath
    Start-SteamClient -SteamPath $steamPath
    Write-Log 'Session ended after install (Steam launched).'
    return 'Exit'
}

# Routes -Action values to install, uninstall, export, or launch handlers.

function Invoke-ManagerAction {
    param(
        [string]$Action,
        [string]$Integration,
        [switch]$Force,
        [string]$OutputDir,
        [switch]$PromptForOutputDir,
        [switch]$IncludeStPlugin
    )

    switch ($Action) {
        'InstallSkyTools' { Start-SkyToolsInstallFlow -Force:$Force }
        'InstallSteamTools' { Start-SteamtoolsInstallFlow -Force:$Force }
        'InstallOpenSteamTools' { Start-OpenSteamToolInstallFlow -Force:$Force }
        'ReinstallSkyTools' { Start-ReinstallFlow -Integration 'SkyTools' -Force:$Force }
        'ReinstallSteamTools' { Start-ReinstallFlow -Integration 'SteamTools' -Force:$Force }
        'ReinstallOpenSteamTools' { Start-ReinstallFlow -Integration 'OpenSteamTools' -Force:$Force }
        'Reinstall' {
            if ($Integration -in @('SkyTools', 'SteamTools', 'OpenSteamTools')) {
                Start-ReinstallFlow -Integration $Integration -Force:$Force
            } else {
                throw 'Reinstall requires -Integration SkyTools, SteamTools, or OpenSteamTools.'
            }
        }
        'Uninstall' { Start-UninstallFlow -Force:$Force }
        { $_ -in 'ExportPack', 'Download' } {
            # From the menu, PromptForOutputDir also gates the export follow-up prompts (scripts, zip, integration).
            $exportParams = @{
                OutputDir            = $OutputDir
                Integration          = $Integration
                PromptForStPlugin    = [bool]$PromptForOutputDir
                PromptForZip         = [bool]$PromptForOutputDir
                PromptForIntegration = [bool]$PromptForOutputDir -and -not ($Integration -in 'SkyTools', 'SteamTools', 'OpenSteamTools')
                PromptForOutputDir   = [bool]$PromptForOutputDir
            }
            if ($PSBoundParameters.ContainsKey('IncludeStPlugin')) {
                $exportParams.IncludeStPlugin = $IncludeStPlugin
            }
            Start-ExportPackFlow @exportParams
        }
        'LaunchSteam' {
            $steamPath = Get-SteamPath
            Start-SteamClient -SteamPath $steamPath
        }
    }
}

# =============================================================================
# Main — CLI action dispatch or interactive menu loop
# =============================================================================

try {
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        throw 'Interactive mode requires running from a script file. Use the README Quick start command instead of piping to Invoke-Expression.'
    }

    Clear-Host
    Set-ManagerWindowTitle
    Initialize-SessionLog
    $script:MenuIntegration = $null

    $fromMenu = [string]::IsNullOrWhiteSpace($Action)

    Invoke-StartupUpdateCheck -Interactive:$fromMenu

    # Normalize legacy/generic action names
    if ($Action -eq 'Download') { $Action = 'ExportPack' }
    if ($Action -eq 'Install') {
        if ($Integration -eq 'SkyTools') { $Action = 'InstallSkyTools' }
        elseif ($Integration -eq 'SteamTools') { $Action = 'InstallSteamTools' }
        elseif ($Integration -eq 'OpenSteamTools') { $Action = 'InstallOpenSteamTools' }
        else { throw 'Install requires -Integration SkyTools, SteamTools, or OpenSteamTools.' }
    }
    if ($Action -eq 'Reinstall') {
        if ($Integration -eq 'SkyTools') { $Action = 'ReinstallSkyTools' }
        elseif ($Integration -eq 'SteamTools') { $Action = 'ReinstallSteamTools' }
        elseif ($Integration -eq 'OpenSteamTools') { $Action = 'ReinstallOpenSteamTools' }
        else { throw 'Reinstall requires -Integration SkyTools, SteamTools, or OpenSteamTools.' }
    }

    if ($fromMenu) {
        while ($true) {
            try {
                Write-MenuHeader
                $menuChoice = Read-MenuChoice
                if (-not $menuChoice) { continue }
                if ($menuChoice -eq 'Exit') {
                    Invoke-ManagerExit
                    Write-Log 'Session ended by user.'
                    break
                }

                Write-ActionBoundary -Action $menuChoice -Phase start
                try {
                    $invokeParams = @{
                        Action             = $menuChoice
                        Force              = $Force
                        OutputDir          = $OutputDir
                        Integration        = if ($script:MenuIntegration) { $script:MenuIntegration } else { $Integration }
                        PromptForOutputDir = $true
                    }
                    if ($PSBoundParameters.ContainsKey('IncludeStPlugin')) {
                        $invokeParams.IncludeStPlugin = $IncludeStPlugin
                    }
                    Invoke-ManagerAction @invokeParams
                    $script:MenuIntegration = $null
                    Write-ActionBoundary -Action $menuChoice -Phase completed

                    if ($menuChoice -in 'InstallSkyTools', 'InstallSteamTools', 'InstallOpenSteamTools', 'ReinstallSkyTools', 'ReinstallSteamTools', 'ReinstallOpenSteamTools') {
                        $postInstall = Invoke-PostInstallSteamPrompt -FromMenu
                        if ($postInstall -eq 'Exit') {
                            Invoke-ManagerExit
                            break
                        }
                    } elseif ($menuChoice -eq 'LaunchSteam') {
                        Clear-Host
                    } else {
                        Wait-MenuReturn
                    }
                } catch [MainMenuReturnException] {
                    $script:MenuIntegration = $null
                    Write-ActionBoundary -Action $menuChoice -Phase cancelled
                    Clear-Host
                } catch {
                    Write-Step ERR $_.Exception.Message
                    Write-ActionBoundary -Action $menuChoice -Phase failed
                    if ($menuChoice -eq 'LaunchSteam') {
                        Clear-Host
                    } else {
                        Wait-MenuReturn
                    }
                }
            } catch [MainMenuReturnException] {
                $script:MenuIntegration = $null
                Clear-Host
            }
        }
    } else {
        Write-ActionBoundary -Action $Action -Phase start
        try {
            $invokeParams = @{
                Action      = $Action
                Force       = $Force
                OutputDir   = $OutputDir
                Integration = $Integration
            }
            if ($PSBoundParameters.ContainsKey('IncludeStPlugin')) {
                $invokeParams.IncludeStPlugin = $IncludeStPlugin
            }
            Invoke-ManagerAction @invokeParams
            Write-ActionBoundary -Action $Action -Phase completed

            if ($Action -in 'InstallSkyTools', 'InstallSteamTools', 'InstallOpenSteamTools', 'ReinstallSkyTools', 'ReinstallSteamTools', 'ReinstallOpenSteamTools') {
                $postInstall = Invoke-PostInstallSteamPrompt
                if ($postInstall -eq 'Exit') {
                    Invoke-ManagerExit
                    return
                }
            }
        } catch {
            Write-ActionBoundary -Action $Action -Phase failed
            throw
        }
    }
} catch [MainMenuReturnException] {
    if ($script:LogFile) {
        Write-Log 'Session ended: returned to main menu (non-interactive).'
    }
} catch {
    Write-Step ERR $_.Exception.Message
    if ($script:LogFile) {
        Write-Log "FATAL: $($_.Exception.Message)"
    }
    exit 1
}
