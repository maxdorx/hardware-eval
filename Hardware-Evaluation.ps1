#Requires -Version 5.1
<#
.SYNOPSIS
    Hardware Evaluation - System Assessment & Reporting Tool
.DESCRIPTION
    Automated hardware and software evaluation for any Windows system (laptop, desktop, server).
    Downloads required utilities on first run, runs a full diagnostic,
    and generates a self-contained HTML report with a buy/no-buy recommendation.
    Adaptive: detects laptop vs desktop, handles missing components gracefully.

.PARAMETER ConfigPath   Path to a JSON config file overriding built-in defaults.
.PARAMETER OutputPath   Directory where reports are saved. Default: Desktop\Reports
.PARAMETER FullAuto     Skip interactive menu; run all modules automatically.
.PARAMETER TestModules  Comma-separated: CPU,RAM,Storage,Battery,GPU,Network,Ports,Software,Thermal
.PARAMETER NoReport     Collect results without generating the HTML report.
.PARAMETER NoDownload   Skip automatic tool downloads; use built-in Windows tools only.

.EXAMPLE
    .\Hardware-Evaluation.ps1                              # interactive menu
    .\Hardware-Evaluation.ps1 -FullAuto                    # fully automatic
    .\Hardware-Evaluation.ps1 -TestModules "CPU,RAM"       # specific modules
    .\Hardware-Evaluation.ps1 -FullAuto -NoDownload        # no internet required

.NOTES
    Requires : Windows 10 Build 19041+ / Windows 11, PowerShell 5.1+
    Elevation: Run as Administrator for full functionality
#>

[CmdletBinding()]
param(
    [string]$ConfigPath  = '',
    [string]$OutputPath  = '',
    [switch]$FullAuto,
    [string]$TestModules = '',
    [switch]$NoReport,
    [switch]$NoDownload
)

$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'

# Capture the full script scriptblock here at top-level scope so the elevation
# re-launch can write the complete script to a temp file (inside Main, MyCommand
# ScriptBlock only contains Main's body, not the whole script).
$Script:SELF_SCRIPTBLOCK = $MyInvocation.MyCommand.ScriptBlock

# ============================================================
#  CONSTANTS & GLOBAL STATE
# ============================================================
$Script:STARTED    = Get-Date
$Script:SESSID     = [System.Guid]::NewGuid().ToString('N').Substring(0,8).ToUpper()
$Script:IS_ADMIN   = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$Script:RESULTS    = [ordered]@{}
$Script:TOOLS      = @{}
$Script:LOG        = [System.Collections.Generic.List[string]]::new()
$Script:ERR_COUNT  = 0
$Script:WARN_COUNT = 0
$Script:CAPS       = @{}   # device capability flags

# ============================================================
#  DEFAULT CONFIGURATION
# ============================================================
$Script:CFG = [ordered]@{
    Thresholds = [ordered]@{
        CPU_MinCores          = 4
        CPU_RecommendedCores  = 8
        CPU_MaxIdleTemp       = 65
        CPU_MaxLoadTemp       = 95
        CPU_WinSAT_Pass       = 6.5
        CPU_MinTier           = 'i5'    # i3 or lower = fail
        CPU_MinIntelGen       = 8       # 8th gen minimum (i5-8xxx and above)
        CPU_MinAMDGen         = 2       # Ryzen 2000 series minimum
        RAM_MinGB             = 8
        RAM_RecommendedGB     = 16
        RAM_MinSpeedMHz       = 2400
        RAM_RecommendedMHz    = 3200
        Storage_MinGB         = 256
        Storage_RecommendedGB = 512
        Storage_MinReadMBs    = 400
        Storage_MinWriteMBs   = 200
        Storage_NVMe_Bonus    = 15
        Battery_MinHealth     = 80
        Battery_WarnHealth    = 90
        Battery_MinCapWh      = 40
        GPU_MinVRAM_MB        = 512
        GPU_DiscreteBonus     = 20
        Network_MinDlMbps     = 50
        Network_MinUlMbps     = 20
        Network_MaxPingMs     = 50
        Network_WiFi6_Bonus   = 10
        Display_MinWidth      = 1920
        Display_MinRefreshHz  = 60
        Software_MaxFailed    = 2
        Software_MaxStartup   = 15
    }
    Tools = [ordered]@{
        WorkDir       = Join-Path $env:LOCALAPPDATA "HWEval\tools"
        DiskSpdUrl    = ''   # resolved dynamically via GitHub API
        SpeedtestUrl  = ''   # resolved dynamically via Ookla CDN
        SmartctlUrl   = ''   # resolved dynamically via GitHub API (smartmontools)
        TimeoutSec    = 120
    }
    Report = [ordered]@{
        OutputDir   = ''
        CompanyName = ''
        OpenAfter   = $true
        IncludeLog  = $true
    }
    # Module weights for the overall score (must sum to 100 for clean math).
    # Ports gets 0 because Test-Ports only checks presence (USB-C, HDMI, etc.) and
    # most laptops pass trivially - it's diagnostic info, not a deployment factor.
    # Software gets 10 because driver errors / BSODs / no AV are deployment blockers.
    Weights = [ordered]@{
        CPU      = 20
        RAM      = 15
        Storage  = 15
        Battery  = 10
        GPU      = 10
        Network  = 10
        Security = 10
        Ports    = 0
        Software = 10
    }
    Verdict = [ordered]@{
        Excellent = 90
        Good      = 78
        Adequate  = 65
        Marginal  = 50
    }
}

# Apply user config overrides
if ($ConfigPath -and (Test-Path $ConfigPath)) {
    try {
        $uc = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        foreach ($sec in $uc.PSObject.Properties) {
            if ($Script:CFG.Contains($sec.Name)) {
                foreach ($k in $sec.Value.PSObject.Properties) {
                    $Script:CFG[$sec.Name][$k.Name] = $k.Value
                }
            }
        }
        Write-Host "  Config loaded: $ConfigPath" -ForegroundColor Cyan
    } catch { Write-Warning "Could not parse config '$ConfigPath': $_" }
}

# Resolve output path
if ($OutputPath) {
    $Script:CFG.Report.OutputDir = $OutputPath
} else {
    $Script:CFG.Report.OutputDir = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Reports'
}

# ============================================================
#  UTILITY FUNCTIONS
# ============================================================

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts][$Level] $Msg"
    $Script:LOG.Add($line)
    switch ($Level) {
        'ERROR' {
            $Script:ERR_COUNT++
            Write-Host '  [' -NoNewline; Write-Host 'ERR' -ForegroundColor Red -NoNewline; Write-Host '] ' -NoNewline
            Write-Host $Msg -ForegroundColor Red
        }
        'WARN'  {
            $Script:WARN_COUNT++
            Write-Host '  [' -NoNewline; Write-Host 'WRN' -ForegroundColor Yellow -NoNewline; Write-Host '] ' -NoNewline
            Write-Host $Msg -ForegroundColor Yellow
        }
        'OK'    {
            Write-Host '  [' -NoNewline; Write-Host ' + ' -ForegroundColor Green -NoNewline; Write-Host '] ' -NoNewline
            Write-Host $Msg -ForegroundColor Green
        }
        'HEAD'  {
            Write-Host "  $Msg" -ForegroundColor Cyan
        }
        default {
            Write-Host '  [' -NoNewline; Write-Host ' . ' -ForegroundColor DarkGray -NoNewline; Write-Host '] ' -NoNewline
            # Highlight numbers and key=value pairs in cyan against the gray narrative
            $parts = [regex]::Split($Msg, '(\b\d+(?:\.\d+)?\s*(?:%|MB/s|MHz|GB|MB|Wh|C|ms|Mbps)?\b|\b(?:PASS|FAIL|WARN|OK|UNKNOWN|None)\b)')
            foreach ($p in $parts) {
                if ([string]::IsNullOrEmpty($p)) { continue }
                if ($p -match '^\d') {
                    Write-Host $p -ForegroundColor Cyan -NoNewline
                } elseif ($p -match '^(PASS|OK)$') {
                    Write-Host $p -ForegroundColor Green -NoNewline
                } elseif ($p -match '^(FAIL)$') {
                    Write-Host $p -ForegroundColor Red -NoNewline
                } elseif ($p -match '^(WARN|UNKNOWN)$') {
                    Write-Host $p -ForegroundColor Yellow -NoNewline
                } else {
                    Write-Host $p -ForegroundColor Gray -NoNewline
                }
            }
            Write-Host ''
        }
    }
}

function Write-Section {
    param([string]$Title)
    $line = '=' * 60
    Write-Host ''
    Write-Host "  $line" -ForegroundColor DarkCyan
    Write-Host "    $Title" -ForegroundColor White -BackgroundColor DarkBlue
    Write-Host "  $line" -ForegroundColor DarkCyan
}

function Write-Banner {
    Clear-Host
    Write-Host ''
    Write-Host '  +============================================================+' -ForegroundColor Cyan
    Write-Host '  |    HARDWARE EVALUATION                                     |' -ForegroundColor Cyan
    Write-Host '  |    System Assessment & Reporting Tool                       |' -ForegroundColor Cyan
    Write-Host '  +============================================================+' -ForegroundColor Cyan
    Write-Host ''
    $rights = if ($Script:IS_ADMIN) { 'Administrator' } else { 'LIMITED - some tests degraded' }
    Write-Host "  Host    : $Script:SESSID @ $env:COMPUTERNAME" -ForegroundColor Gray
    Write-Host "  Date    : $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Gray
    Write-Host "  Rights  : $rights" -ForegroundColor $(if ($Script:IS_ADMIN) { 'Green' } else { 'Yellow' })
    Write-Host ''
}

function New-Result {
    param(
        [string]$Module,
        [int]$Score       = 0,
        [string]$Status   = 'SKIP',
        [string]$Summary  = '',
        [hashtable]$Data  = @{},
        [string[]]$Issues = @(),
        [string[]]$Good   = @(),
        [object[]]$Fixables = @()
    )
    [PSCustomObject]@{
        Module    = $Module
        Score     = $Score
        Status    = $Status
        Summary   = $Summary
        Data      = $Data
        Issues    = [System.Collections.Generic.List[string]]($Issues)
        Good      = [System.Collections.Generic.List[string]]($Good)
        Fixables  = [System.Collections.Generic.List[object]]($Fixables)
        Timestamp = Get-Date
    }
}

# A FixableIssue represents a software/configuration problem the technician can fix
# in-place (no parts), then re-run this script to verify. Distinct from physical Repairs.
function New-FixableIssue {
    param(
        [string]$Title,
        [string]$Reason,
        [string]$FixCommand,         # PowerShell or CMD command to run (admin)
        [string]$EstimatedTime = '5 min',
        [bool]$RequiresReboot = $false
    )
    [ordered]@{
        Title          = $Title
        Reason         = $Reason
        FixCommand     = $FixCommand
        EstimatedTime  = $EstimatedTime
        RequiresReboot = $RequiresReboot
    }
}

function Clamp100 { param([double]$v); [int][Math]::Max(0, [Math]::Min(100, $v)) }
function ToGB    { param([long]$b); [Math]::Round($b / 1GB, 1) }
function ToMB    { param([long]$b); [Math]::Round($b / 1MB, 0) }

function Score-ToStatus {
    param([int]$s)
    if ($s -ge 80) { return 'PASS' }
    if ($s -ge 50) { return 'WARN' }
    return 'FAIL'
}

# ============================================================
#  ENVIRONMENT INITIALIZATION
# ============================================================

function Initialize-Env {
    Write-Section 'Initializing Environment'

    $wd = $Script:CFG.Tools.WorkDir
    if (-not (Test-Path $wd)) { New-Item -ItemType Directory $wd -Force | Out-Null }
    Write-Log "Work dir : $wd"

    $rd = $Script:CFG.Report.OutputDir
    if (-not (Test-Path $rd)) { New-Item -ItemType Directory $rd -Force | Out-Null }
    Write-Log "Report   : $rd"

    if (-not $Script:IS_ADMIN) {
        Write-Log 'Not running as Administrator. SMART, WinSAT, thermal monitoring may be limited.' 'WARN'
    }

    Detect-Capabilities
    Write-Log 'Environment ready.' 'OK'
}

function Detect-Capabilities {
    # Battery
    $b = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
    $Script:CAPS.HasBattery = ($null -ne $b) -and (@($b).Count -gt 0)

    # WiFi  -  Get-NetAdapter is reliable; Win32_NetworkAdapter AdapterTypeId varies by driver
    $w = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object {
        $_.InterfaceDescription -match 'Wireless|Wi-Fi|WiFi|802\.11|WLAN' -or
        $_.PhysicalMediaType -eq 'Native 802.11'
    }
    $Script:CAPS.HasWiFi = (@($w).Count -gt 0)

    # Discrete GPU  -  Intel/AMD integrated and Microsoft virtual adapters are NOT discrete
    $g = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue
    $Script:CAPS.HasDiscreteGPU = ($null -ne $g) -and (
        @($g | Where-Object {
            $_.Name -match 'NVIDIA|GeForce|Quadro|RTX|GTX|AMD\s+Radeon|RX\s+\d|FirePro|Radeon\s+Pro|Radeon\s+RX' -and
            $_.AdapterRAM -gt 512MB
        }).Count -gt 0
    )

    # Ethernet — exclude 'Not Present' ghost entries (unplugged USB dongles that Windows remembers)
    # and Disabled adapters, which on thin laptops usually means silicon with no physical port.
    # An adapter that is Up or Disconnected has a real port; Disabled/Not Present do not count.
    $eth = @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object {
        ($_.PhysicalMediaType -eq '802.3' -or
         ($_.InterfaceDescription -match 'Ethernet|LAN|Gigabit|10GbE' -and
          $_.InterfaceDescription -notmatch 'Wireless|Wi-Fi|WiFi|802\.11')) -and
        $_.Status -in @('Up', 'Disconnected')
    })
    $Script:CAPS.HasEthernet    = ($eth.Count -gt 0)
    # True only when a cable is (or was recently) live — used to distinguish real port vs always-dark silicon
    $Script:CAPS.EthernetActive = ($eth | Where-Object { $_.Status -eq 'Up' } | Measure-Object).Count -gt 0

    # Bluetooth
    $bt = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue
    $Script:CAPS.HasBluetooth = ($null -ne $bt) -and (@($bt).Count -gt 0)

    # Chassis type: 8-11,14,30-32 = laptop/notebook/tablet
    $chassis = (Get-CimInstance Win32_SystemEnclosure -ErrorAction SilentlyContinue).ChassisTypes
    $Script:CAPS.IsLaptop = ($null -ne ($chassis | Where-Object { $_ -in @(8,9,10,11,14,30,31,32) }))

    # VM / hypervisor detection
    $csVM = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $Script:CAPS.IsVM = (
        $csVM.Model        -match 'Virtual|VMware|VirtualBox|HyperV|KVM|QEMU|Bochs|Xen' -or
        ($csVM.Manufacturer -match 'VMware|innotek|QEMU|Xen|Microsoft Corporation' -and
            $csVM.Model    -match 'Virtual')
    )

    # Server detection: ProductType 2/3 = Domain Controller / Server
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $Script:CAPS.IsServer = ($os -and $os.ProductType -in @(2,3)) -or ($os -and $os.Caption -match 'Server')

    Write-Log ("Caps: Form=$(if ($Script:CAPS.IsLaptop){'Laptop'}elseif($Script:CAPS.IsServer){'Server'}else{'Desktop'}) " +
               "VM=$($Script:CAPS.IsVM) " +
               "Bat=$($Script:CAPS.HasBattery) WiFi=$($Script:CAPS.HasWiFi) ETH=$($Script:CAPS.HasEthernet) " +
               "dGPU=$($Script:CAPS.HasDiscreteGPU) BT=$($Script:CAPS.HasBluetooth)")

    # Loud warning for servers — this tool is calibrated for end-user laptops/desktops
    if ($Script:CAPS.IsServer) {
        Write-Host ''
        Write-Host '  +============================================================+' -ForegroundColor Yellow
        Write-Host '  |  SERVER OS DETECTED                                        |' -ForegroundColor Yellow
        Write-Host '  +============================================================+' -ForegroundColor Yellow
        Write-Host '  This tool is built for evaluating end-user laptops/desktops.'  -ForegroundColor Yellow
        Write-Host '  On a server, expect these inaccuracies in the report:'         -ForegroundColor Yellow
        Write-Host '    - SMART on RAID arrays returns UNKNOWN (controller hides it)' -ForegroundColor Gray
        Write-Host '    - GPU/Display tests are pointless (no real GPU)'              -ForegroundColor Gray
        Write-Host '    - Battery / Wi-Fi / Bluetooth absent (correctly skipped)'     -ForegroundColor Gray
        Write-Host '    - Use-case scoring assumes office user, not server workload'  -ForegroundColor Gray
        Write-Host '  Continuing anyway. Treat the report as advisory, not authoritative.' -ForegroundColor Yellow
        Write-Host ''
    }
}

# ============================================================
#  TOOL DOWNLOAD
# ============================================================

function Resolve-ToolUrls {
    # DiskSpd: query GitHub releases API for latest asset URL
    try {
        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add('User-Agent', 'HardwareEval')
        $json = $wc.DownloadString('https://api.github.com/repos/microsoft/diskspd/releases/latest')
        $rel  = $json | ConvertFrom-Json
        $asset = $rel.assets | Where-Object { $_.name -match '\.zip$' } | Select-Object -First 1
        if ($asset) {
            $Script:CFG.Tools.DiskSpdUrl = $asset.browser_download_url
            Write-Log "DiskSpd latest: $($asset.browser_download_url)" 'OK'
        } else {
            throw "no zip asset found"
        }
    } catch {
        Write-Log "DiskSpd GitHub API failed ($_), using pinned fallback" 'WARN'
        $Script:CFG.Tools.DiskSpdUrl = 'https://github.com/microsoft/diskspd/releases/download/v2.1/DiskSpd.ZIP'
    }

    # Speedtest CLI: scrape Ookla's CLI download page for the win64 zip link
    # (their JSON endpoint blocks WebClient with 403; scraping the HTML page works)
    try {
        $wc2 = New-Object System.Net.WebClient
        $wc2.Headers.Add('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')
        $html = $wc2.DownloadString('https://www.speedtest.net/apps/cli')
        # Page contains href links like: ookla-speedtest-1.2.0-win64.zip
        if ($html -match 'href="(https://[^"]*ookla-speedtest-[\d\.]+-win64\.zip)"') {
            $Script:CFG.Tools.SpeedtestUrl = $Matches[1]
            Write-Log "Speedtest latest: $($Matches[1])" 'OK'
        } else { throw "win64 zip link not found in page" }
    } catch {
        Write-Log "Speedtest page scrape failed ($_), using pinned fallback" 'WARN'
        $Script:CFG.Tools.SpeedtestUrl = 'https://install.speedtest.net/app/cli/ookla-speedtest-1.2.0-win64.zip'
    }

    # smartmontools (smartctl) - GitHub API: prefer zip, accept setup.exe if no zip available
    try {
        $wc3   = New-Object System.Net.WebClient
        $wc3.Headers.Add('User-Agent', 'HardwareEval')
        $json3 = $wc3.DownloadString('https://api.github.com/repos/smartmontools/smartmontools/releases/latest')
        $rel3  = $json3 | ConvertFrom-Json
        # Try zip assets first (portable), then fall back to win32 setup exe
        $asset3 = $rel3.assets | Where-Object { $_.name -match 'win64.*\.zip$' } | Select-Object -First 1
        if (-not $asset3) { $asset3 = $rel3.assets | Where-Object { $_.name -match 'win.*\.zip$' }   | Select-Object -First 1 }
        if (-not $asset3) { $asset3 = $rel3.assets | Where-Object { $_.name -match '\.zip$' }         | Select-Object -First 1 }
        if (-not $asset3) { $asset3 = $rel3.assets | Where-Object { $_.name -match 'win.*setup.*\.exe$' -or $_.name -match 'win.*\.exe$' } | Select-Object -First 1 }
        if ($asset3) {
            $Script:CFG.Tools.SmartctlUrl = $asset3.browser_download_url
            # tag whether the asset is a setup exe so Initialize-Tools can handle it
            $Script:CFG.Tools.SmartctlDirect = ($asset3.name -match '\.exe$')
            Write-Log "smartmontools latest: $($asset3.name)" 'OK'
        } else { throw "no usable asset found" }
    } catch {
        Write-Log "smartmontools GitHub API failed ($_), using pinned fallback" 'WARN'
        $Script:CFG.Tools.SmartctlUrl    = 'https://github.com/smartmontools/smartmontools/releases/download/RELEASE_7_4/smartmontools-7.4-1.win32-setup.exe'
        $Script:CFG.Tools.SmartctlDirect = $true
    }

}

function Get-Tool {
    param(
        [string]$Name,
        [string]$Url,
        [string]$ExeName,
        [switch]$Direct   # download URL is the executable itself (no ZIP extraction needed)
    )
    if ($NoDownload) { return $false }
    if (-not $Url)   { Write-Log "No URL for '$Name', skipping" 'WARN'; return $false }

    $wd = $Script:CFG.Tools.WorkDir

    # Check if already cached from a previous run in this session's workdir
    $already = Get-ChildItem $wd -Filter $ExeName -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($already) {
        $Script:TOOLS[$Name] = $already.FullName
        Write-Log "Tool '$Name' cached: $($already.FullName)" 'OK'
        return $true
    }

    Write-Log "Downloading $Name ..."
    try {
        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add('User-Agent', 'HardwareEval')

        if ($Direct) {
            # Direct EXE download - preserve the original filename from the URL
            $origName = [System.IO.Path]::GetFileName(($Url -split '\?' | Select-Object -First 1))
            if (-not $origName -or $origName -notmatch '\.exe$') { $origName = $ExeName }
            $destPath = Join-Path $wd $origName
            $wc.DownloadFile($Url, $destPath)
            if (-not (Test-Path $destPath) -or (Get-Item $destPath).Length -lt 65536) {
                throw "Downloaded file is too small or missing"
            }
            $Script:TOOLS[$Name] = $destPath
            Write-Log "Tool '$Name' ready: $destPath" 'OK'
            return $true
        } else {
            # ZIP download - extract and find exe
            $zipPath = Join-Path $wd "$Name.zip"
            $wc.DownloadFile($Url, $zipPath)
            if (-not (Test-Path $zipPath) -or (Get-Item $zipPath).Length -lt 10240) {
                throw "Downloaded file is too small or missing"
            }
            Expand-Archive -Path $zipPath -DestinationPath $wd -Force
            Remove-Item $zipPath -ErrorAction SilentlyContinue

            $found = Get-ChildItem $wd -Filter $ExeName -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($found) {
                $Script:TOOLS[$Name] = $found.FullName
                Write-Log "Tool '$Name' ready: $($found.FullName)" 'OK'
                return $true
            }
            throw "Executable '$ExeName' not found after extraction"
        }
    } catch {
        Write-Log "Cannot download '$Name': $_" 'WARN'
        return $false
    }
}

function Initialize-Tools {
    Write-Section 'Downloading Required Tools'
    $Script:TOOLS.DiskSpd   = $null
    $Script:TOOLS.Speedtest = $null
    $Script:TOOLS.Smartctl  = $null

    if ($NoDownload) {
        Write-Log '-NoDownload set: skipping tool downloads, using built-in Windows tools only.' 'WARN'
        return
    }

    # Resolve latest download URLs dynamically (no hardcoded version numbers)
    Resolve-ToolUrls

    # DiskSpd (Microsoft official disk benchmark)
    $ok = Get-Tool -Name 'DiskSpd' -Url $Script:CFG.Tools.DiskSpdUrl -ExeName 'diskspd.exe'
    if (-not $ok) {
        $sys = Get-Command 'diskspd.exe' -ErrorAction SilentlyContinue
        if ($sys) { $Script:TOOLS.DiskSpd = $sys.Source; Write-Log "DiskSpd on PATH: $($sys.Source)" 'OK' }
    }

    # Speedtest CLI (Ookla)
    $ok2 = Get-Tool -Name 'Speedtest' -Url $Script:CFG.Tools.SpeedtestUrl -ExeName 'speedtest.exe'
    if (-not $ok2) {
        $sys2 = Get-Command 'speedtest.exe' -ErrorAction SilentlyContinue
        if ($sys2) { $Script:TOOLS.Speedtest = $sys2.Source; Write-Log "Speedtest on PATH: $($sys2.Source)" 'OK' }
    }

    # smartmontools (smartctl) - real SMART attribute data
    # Latest release ships as a setup.exe (NSIS installer); run silently to extract into workdir
    $wd = $Script:CFG.Tools.WorkDir
    $isDirect = $Script:CFG.Tools.SmartctlDirect -eq $true
    $ok3 = $false
    if ($isDirect) {
        # Download setup.exe then extract silently
        $setupExe = Join-Path $wd 'smartctl-setup.exe'
        try {
            if (-not (Test-Path $setupExe)) {
                Write-Log 'Downloading smartmontools setup ...'
                $wc6 = New-Object System.Net.WebClient
                $wc6.Headers.Add('User-Agent', 'HardwareEval')
                $wc6.DownloadFile($Script:CFG.Tools.SmartctlUrl, $setupExe)
            }
            if ((Get-Item $setupExe -ErrorAction SilentlyContinue).Length -gt 100KB) {
                $smDir = Join-Path $wd 'smartmontools'
                New-Item $smDir -ItemType Directory -Force | Out-Null
                # NSIS /S silent install, /D sets the destination (must be absolute, no trailing slash)
                $proc = Start-Process $setupExe -ArgumentList "/S /D=$smDir" -Wait -PassThru -ErrorAction Stop
                $sc = Get-ChildItem $smDir -Filter 'smartctl.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($sc) {
                    $Script:TOOLS.Smartctl = $sc.FullName
                    Write-Log "smartctl ready: $($sc.FullName)" 'OK'
                    $ok3 = $true
                } else { Write-Log 'smartctl.exe not found after setup extraction' 'WARN' }
            } else { Write-Log 'smartmontools setup too small - skipping' 'WARN' }
        } catch { Write-Log "smartmontools setup failed: $_" 'WARN' }
    } else {
        $ok3 = Get-Tool -Name 'Smartctl' -Url $Script:CFG.Tools.SmartctlUrl -ExeName 'smartctl.exe'
    }
    if (-not $ok3) {
        $sys3 = Get-Command 'smartctl.exe' -ErrorAction SilentlyContinue
        if ($sys3) { $Script:TOOLS.Smartctl = $sys3.Source; Write-Log "smartctl on PATH: $($sys3.Source)" 'OK' }
    }

}

# ============================================================
#  SYSTEM INFO  (always runs, no scoring)
# ============================================================

function Get-SysInfo {
    Write-Section 'System Information'
    $cs  = Get-CimInstance Win32_ComputerSystem  -ErrorAction SilentlyContinue
    $os  = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $bio = Get-CimInstance Win32_BIOS            -ErrorAction SilentlyContinue
    $enc = Get-CimInstance Win32_SystemEnclosure -ErrorAction SilentlyContinue

    $upSec  = ((Get-Date) - $os.LastBootUpTime).TotalSeconds
    $upStr  = ('{0}d {1}h {2}m' -f [int]($upSec/86400), [int](($upSec % 86400)/3600), [int](($upSec % 3600)/60))

    $biosStr = ''
    if ($bio) { $biosStr = "$($bio.Manufacturer) $($bio.SMBIOSBIOSVersion)" }

    $d = [ordered]@{
        Manufacturer  = if ($cs)  { $cs.Manufacturer }  else { 'N/A' }
        Model         = if ($cs)  { $cs.Model }          else { 'N/A' }
        Serial        = if ($enc) { $enc.SerialNumber }  else { 'N/A' }
        BIOS          = $biosStr
        OS            = if ($os)  { $os.Caption }        else { 'N/A' }
        Build         = if ($os)  { $os.BuildNumber }    else { 'N/A' }
        Architecture  = if ($os)  { $os.OSArchitecture } else { 'N/A' }
        Hostname      = $env:COMPUTERNAME
        Domain        = if ($cs)  { $cs.Domain }         else { 'N/A' }
        LastBoot      = if ($os)  { $os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' }
        Uptime        = $upStr
        User          = "$env:USERDOMAIN\$env:USERNAME"
        FormFactor    = if ($Script:CAPS.IsLaptop) { 'Laptop/Notebook' } else { 'Desktop/Other' }
    }

    foreach ($k in $d.Keys) { Write-Log "$($k.PadRight(14)): $($d[$k])" }

    $Script:RESULTS.SysInfo = New-Result -Module SysInfo -Status PASS -Score 100 `
        -Summary "$($d.Manufacturer) $($d.Model) - $($d.OS)" -Data $d
}

# Single-keystroke Y/N prompt (Linux-style — no Enter needed).
# Loops until a valid Y or N is pressed.
function Read-YesNo {
    param([string]$Prompt, [ConsoleColor]$Color = 'Cyan')
    while ($true) {
        Write-Host "$Prompt [y/n]: " -ForegroundColor $Color -NoNewline
        $key = [Console]::ReadKey($true)
        $c   = "$($key.KeyChar)".ToUpper()
        if ($c -eq 'Y') { Write-Host 'y' -ForegroundColor Green; return $true }
        if ($c -eq 'N') { Write-Host 'n' -ForegroundColor Red;   return $false }
        Write-Host ''
        Write-Host '  Press y or n.' -ForegroundColor Yellow
    }
}

# Technician name + visual damage assessment.
# Collected before tests start so the technician can fill it in while waiting.
$Script:TechInfo = $null

function Get-TechnicianInfo {
    if ($Script:FullAuto) {
        $Script:TechInfo = [ordered]@{
            Technician        = "$env:USERDOMAIN\$env:USERNAME"
            DamageStatus      = 'Not assessed (FullAuto mode)'
            DamageDetails     = 'N/A'
            FinalVerdict      = 'N/A (FullAuto mode)'
            FinalVerdictNote  = ''
        }
        return
    }

    Write-Section 'Pre-Evaluation Info'
    Write-Host '  Your full name: ' -ForegroundColor Cyan -NoNewline
    $techName = (Read-Host).Trim()
    if (-not $techName) { $techName = "$env:USERDOMAIN\$env:USERNAME" }

    if (Read-YesNo '  Any visible physical damage on the laptop?') {
        Write-Host '  Describe the damage (cracks, dents, missing keys, etc.): ' -ForegroundColor Yellow -NoNewline
        $details = (Read-Host).Trim()
        if (-not $details) { $details = '(no description provided)' }
        $status = 'Damage reported'
    } else {
        $details = 'None'
        $status  = 'No physical damage reported'
    }

    $Script:TechInfo = [ordered]@{
        Technician        = $techName
        DamageStatus      = $status
        DamageDetails     = $details
        FinalVerdict      = ''
        FinalVerdictNote  = ''
    }
    Write-Log "Evaluator: $techName" 'OK'
    Write-Log "Damage check: $status" 'OK'
}

# Asks the technician for their own recommendation AFTER seeing the auto-computed score.
# The technician's expert call is shown alongside the script's verdict in the report.
function Get-TechnicianVerdict {
    if (-not $Script:TechInfo) { return }
    if ($Script:FullAuto) { return }

    $ov = $Script:RESULTS.Overall
    $name = $Script:TechInfo.Technician
    Write-Host ''
    Write-Section "$name's Final Verdict"
    Write-Host "  Auto score: $($ov.Score)/100 - $($ov.Verdict)" -ForegroundColor Gray

    # Surface any user-reported defects so the technician sees them before deciding
    $blockers = [System.Collections.Generic.List[string]]::new()
    if ($Script:TechInfo -and $Script:TechInfo.DamageStatus -eq 'Damage reported') {
        $blockers.Add("Physical damage: $($Script:TechInfo.DamageDetails)")
    }
    $gd = if ($Script:RESULTS.GPU) { $Script:RESULTS.GPU.Data } else { $null }
    if ($gd) {
        if ("$($gd.Dead_Pixels)"   -match 'Reported') { $blockers.Add("Dead/stuck pixels: $($gd.Dead_Pixels)") }
        if ("$($gd.Screen_Quality)" -match '^Issue')  { $blockers.Add("Screen quality: $($gd.Screen_Quality)") }
    }
    $id = if ($Script:RESULTS.Input) { $Script:RESULTS.Input.Data } else { $null }
    if ($id) {
        if ($id.Keyboard_Broken_Keys)                  { $blockers.Add("Broken keys: $($id.Keyboard_Broken_Keys)") }
        elseif ("$($id.Keyboard_Check)" -match '^FAIL'){ $blockers.Add("Keyboard: $($id.Keyboard_Check)") }
        if ("$($id.Touch_Screen_Check)" -match '^FAIL'){ $blockers.Add("Touch screen: $($id.Touch_Screen_Check)") }
    }
    if ($blockers.Count -gt 0) {
        Write-Host ''
        Write-Host '  *** YOU REPORTED THE FOLLOWING DEFECTS ***' -ForegroundColor Red
        foreach ($b in $blockers) { Write-Host "    - $b" -ForegroundColor Yellow }
        Write-Host '  These will be marked as DEPLOYMENT BLOCKERS in the report.' -ForegroundColor Red
        Write-Host '  This laptop should NOT be assigned until they are addressed.' -ForegroundColor Red
    }

    Write-Host ''
    $verdict = if (Read-YesNo '  Based on everything you observed, do you recommend deploying this laptop?') {
        "Recommended by $name"
    } else {
        "NOT recommended by $name"
    }

    Write-Host '  Comment / reasoning (optional, press Enter to skip): ' -ForegroundColor Cyan -NoNewline
    $note = (Read-Host).Trim()

    $Script:TechInfo.FinalVerdict     = $verdict
    $Script:TechInfo.FinalVerdictNote = $note
    Write-Log "$name's verdict: $verdict" 'OK'
    if ($note) { Write-Log "  Note: $note" 'OK' }
}

# ============================================================
#  MODULE: CPU
# ============================================================

function Get-CPUGenInfo {
    param([string]$Name)
    # Returns: @{ Vendor; Tier; TierNum; Gen; Pass; Reason }
    $r = @{ Vendor = 'Other'; Tier = 'Unknown'; TierNum = 0; Gen = 0; Pass = $true; Reason = '' }
    $t = $Script:CFG.Thresholds

    # Intel Core i3/i5/i7/i9
    # Model format: "i7-8650U"  (4-digit = gen is 1st digit)
    #               "i7-10750H" (5-digit = gen is 1st 2 digits)
    if ($Name -match 'Core.*i(\d)[\s-](\d{4,5})') {
        $tierNum  = [int]$Matches[1]        # 3, 5, 7, 9
        $modelNum = $Matches[2]
        $gen      = if ($modelNum.Length -ge 5) { [int]$modelNum.Substring(0,2) }
                    else { [int]$modelNum.Substring(0,1) }
        $r.Vendor  = 'Intel'
        $r.Tier    = "i$tierNum"
        $r.TierNum = $tierNum
        $r.Gen     = $gen

        $minTierNum = [int]($t.CPU_MinTier -replace 'i','')
        if ($tierNum -lt $minTierNum) {
            $r.Pass   = $false
            $r.Reason = "Intel Core i$tierNum is below minimum $($t.CPU_MinTier) tier"
        } elseif ($gen -lt $t.CPU_MinIntelGen) {
            $r.Pass   = $false
            $r.Reason = "$($gen)th gen Intel is below minimum $($t.CPU_MinIntelGen)th gen"
        } else {
            $r.Reason = "Intel Core i$tierNum $($gen)th gen meets requirements"
        }
    }
    # AMD Ryzen 3/5/7/9  -  model: "Ryzen 5 3600", "Ryzen 7 5800H", "Ryzen 5 PRO 7640U", "Ryzen 9 7950X"
    elseif ($Name -match 'Ryzen\s+(\d)(?:\s+PRO)?\s+(\d{4,5})') {
        $tierNum  = [int]$Matches[1]   # 3, 5, 7, 9
        $modelNum = $Matches[2]
        $gen      = [int]$modelNum.Substring(0,1)   # first digit = generation (1xxx=1, 5xxx=5, 7xxx=7, 9xxx=9)
        $isPRO    = $Name -match 'Ryzen\s+\d\s+PRO'
        $r.Vendor  = 'AMD'
        $r.Tier    = "Ryzen $tierNum$(if ($isPRO) {' PRO'})"
        $r.TierNum = $tierNum
        $r.Gen     = $gen

        $minTierNum = [int]($t.CPU_MinTier -replace 'i','')
        if ($tierNum -lt $minTierNum) {
            $r.Pass   = $false
            $r.Reason = "AMD Ryzen $tierNum is below minimum Ryzen $minTierNum tier"
        } elseif ($gen -lt $t.CPU_MinAMDGen) {
            $r.Pass   = $false
            $r.Reason = "Ryzen $($gen)000 series is below minimum Ryzen $($t.CPU_MinAMDGen)000 series"
        } else {
            $r.Reason = "AMD Ryzen $tierNum $($gen)000 series meets requirements"
        }
    }
    # AMD Ryzen AI (Zen 5 AI-branded laptops: "Ryzen AI 9 HX 370", "Ryzen AI 5 340")
    elseif ($Name -match 'Ryzen\s+AI\s+(\d+)') {
        $r.Vendor  = 'AMD'
        $r.Tier    = "Ryzen AI"
        $r.TierNum = 7   # treat as Ryzen 7 equivalent for scoring
        $r.Gen     = 9   # Ryzen AI = Zen 5 (2024+) = gen 9
        $r.Pass    = $true
        $r.Reason  = "AMD Ryzen AI - modern Zen 5 processor, exceeds all requirements"
    }
    # Intel Core Ultra (13th gen+ marketing name)
    elseif ($Name -match 'Core Ultra\s+(\d)\s+(\d{3})') {
        $r.Vendor  = 'Intel'
        $r.Tier    = "Ultra $($Matches[1])"
        $r.TierNum = [int]$Matches[1]
        $r.Gen     = 14   # Core Ultra starts at gen 14
        $r.Pass    = $true
        $r.Reason  = "Intel Core Ultra - exceeds all requirements"
    }

    return $r
}

function Test-CPU {
    Write-Section 'CPU Assessment'
    $issues = [System.Collections.Generic.List[string]]::new()
    $good   = [System.Collections.Generic.List[string]]::new()
    $d      = [ordered]@{}

    try {
        $cpus    = @(Get-CimInstance Win32_Processor -ErrorAction Stop)
        $cpu     = $cpus[0]
        $cores   = ($cpus | Measure-Object -Property NumberOfCores              -Sum).Sum
        $threads = ($cpus | Measure-Object -Property NumberOfLogicalProcessors  -Sum).Sum
        $mhz     = $cpu.MaxClockSpeed
        $t       = $Script:CFG.Thresholds

        $d.Name          = $cpu.Name.Trim()
        $d.Manufacturer  = $cpu.Manufacturer
        $d.Cores         = $cores
        $d.Threads       = $threads
        $d.Base_GHz      = [Math]::Round($mhz / 1000, 2)
        $d.CurrentLoad   = "$($cpu.LoadPercentage)%"
        $d.L2Cache_KB    = $cpu.L2CacheSize
        $d.L3Cache_MB    = [Math]::Round($cpu.L3CacheSize / 1024, 1)
        $d.Socket        = $cpu.SocketDesignation

        Write-Log "CPU: $($d.Name)"
        Write-Log "Cores $cores / Threads $threads / $($d.Base_GHz) GHz"

        # Idle temperature (read before bench so the bench doesn't skew it)
        $idleTemp = Read-CPUTemp
        if ($idleTemp -gt 0) {
            $d.IdleTemp_C = $idleTemp
            Write-Log "Idle temp: $($idleTemp) C"
        } else { $d.IdleTemp_C = 'N/A' }

        # Quick PS benchmark (~15s) - measures actual real-world CPU performance
        Write-Log 'Running quick CPU benchmark (~15s)...'
        $pts = Invoke-CPUBench -Seconds 15
        $d.BenchScore = $pts
        Write-Log "Bench: $pts points"

        # Score: cores 30%, real-world bench 40%, baseline 30
        # Reference: 18,000 points = 100% benchScore (typical 8th-gen i7 ultrabook).
        # Base clock is shown in $d.Base_GHz but no longer factored into the score -
        # WMI MaxClockSpeed reports a moving target (varies between base/boost/current),
        # while the bench measures actual sustained throughput on this machine.
        $coreScore  = Clamp100 (($cores / $t.CPU_RecommendedCores) * 100)
        $benchScore = Clamp100 (($pts / 18000.0) * 100)
        $score      = Clamp100 (($coreScore * 0.3) + ($benchScore * 0.4) + 30)
        $d.Score_Breakdown = "core:$([int]$coreScore) bench:$([int]$benchScore) (ref 18k pts)"

        if ($cores -lt $t.CPU_MinCores) {
            $issues.Add("$cores cores below minimum $($t.CPU_MinCores)")
        } else {
            $good.Add("$cores physical cores")
        }
        if ($cores -ge $t.CPU_RecommendedCores) {
            $good.Add("Meets recommended core count ($($t.CPU_RecommendedCores)+)")
        }
        if ($benchScore -ge 80) { $good.Add("Strong benchmark performance ($pts pts)") }

        # Generation and tier check (i5 8th gen minimum policy)
        $genInfo = Get-CPUGenInfo -Name $d.Name
        $d.CPU_Tier       = $genInfo.Tier
        $d.CPU_Generation = if ($genInfo.Gen -gt 0) { "$($genInfo.Gen)th gen" } else { 'Unknown' }
        $d.CPU_Vendor     = $genInfo.Vendor
        Write-Log "CPU tier: $($genInfo.Tier) | Gen: $($genInfo.Gen) | $($genInfo.Reason)"
        if (-not $genInfo.Pass) {
            $issues.Add("POLICY FAIL: $($genInfo.Reason)")
            $score = Clamp100 ($score - 30)    # hard penalty for failing minimum spec
        } else {
            $good.Add("$($genInfo.Reason)")
        }

        # WinSAT score from cached results
        $ws = Read-WinSATScore 'CPU'
        if ($ws -gt 0) {
            $d.WinSAT_CPU = $ws
            Write-Log "WinSAT CPU: $ws"
            if ($ws -ge $t.CPU_WinSAT_Pass) {
                $good.Add("WinSAT CPU score $ws (pass >= $($t.CPU_WinSAT_Pass))")
            } else {
                $issues.Add("WinSAT CPU $ws below threshold $($t.CPU_WinSAT_Pass)")
                $score = Clamp100 ($score - 10)
            }
        }

        # Apply idle-temp penalty (data already collected above)
        if ($idleTemp -gt 0) {
            if ($idleTemp -gt ($t.CPU_MaxIdleTemp + 2)) {
                $issues.Add("Idle temp $($idleTemp) C exceeds threshold $($t.CPU_MaxIdleTemp) C")
                $score = Clamp100 ($score - 15)
            } else {
                $good.Add("Idle temp $($idleTemp) C is healthy")
            }
        }

        $status = Score-ToStatus $score
        $Script:RESULTS.CPU = New-Result -Module CPU -Score $score -Status $status `
            -Summary "$($d.Name) | $cores C/$threads T | $($d.Base_GHz) GHz" `
            -Data $d -Issues $issues -Good $good

        Write-Log "CPU Score: $score ($status)" 'OK'
    } catch {
        Write-Log "CPU error: $_" 'ERROR'
        $Script:RESULTS.CPU = New-Result -Module CPU -Score 0 -Status ERROR -Summary "Error: $_"
    }
}

function Read-WinSATScore {
    param([string]$Comp)
    try {
        $dir = 'C:\Windows\Performance\WinSAT\DataStore'
        if (-not (Test-Path $dir)) { return 0 }
        $f = Get-ChildItem $dir -Filter '*Formal.Assessment*.xml' -ErrorAction SilentlyContinue |
             Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $f) { return 0 }
        # Ignore stale WinSAT data older than 12 months — scores from years-old assessments mislead scoring
        if ($f.LastWriteTime -lt (Get-Date).AddMonths(-12)) { return 0 }
        [xml]$x = Get-Content $f.FullName -ErrorAction Stop
        $v = switch ($Comp) {
            'CPU'      { $x.WinSAT.WinSPR.CpuScore }
            'Memory'   { $x.WinSAT.WinSPR.MemoryScore }
            'Disk'     { $x.WinSAT.WinSPR.DiskScore }
            'Graphics' { $x.WinSAT.WinSPR.GraphicsScore }
            default    { 0 }
        }
        return [double]$v
    } catch { return 0 }
}

function Read-CPUTemp {
    try {
        $zones = Get-CimInstance -Namespace root\wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop
        $temps = $zones | ForEach-Object { [Math]::Round(($_.CurrentTemperature / 10) - 273.15, 1) } |
                 Where-Object { $_ -gt 0 -and $_ -lt 120 }
        if ($temps) { return ($temps | Measure-Object -Maximum).Maximum }
    } catch {}
    return 0
}

function Invoke-CPUBench {
    param([int]$Seconds = 15)
    $n    = [Environment]::ProcessorCount
    $pool = [RunspaceFactory]::CreateRunspacePool(1, $n)
    $pool.Open()
    $end  = (Get-Date).AddSeconds($Seconds)

    $jobs = 1..$n | ForEach-Object {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript({
            param($e)
            $c = 0L
            while ((Get-Date) -lt $e) {
                $x = 0.0
                for ($i = 1; $i -le 5000; $i++) { $x += [Math]::Sqrt($i) * [Math]::PI }
                $c++
            }
            $c
        }).AddArgument($end)
        @{ P = $ps; H = $ps.BeginInvoke() }
    }

    $total = 0L
    foreach ($j in $jobs) {
        try { $total += [long]($j.P.EndInvoke($j.H))[0] } catch {}
        $j.P.Dispose()
    }
    $pool.Close(); $pool.Dispose()
    return [int]($total * 100 / $Seconds)
}

# ============================================================
#  MODULE: RAM
# ============================================================

function Test-RAM {
    Write-Section 'Memory (RAM) Assessment'
    $issues = [System.Collections.Generic.List[string]]::new()
    $good   = [System.Collections.Generic.List[string]]::new()
    $d      = [ordered]@{}

    try {
        $os        = Get-CimInstance Win32_OperatingSystem     -ErrorAction Stop
        $cs        = Get-CimInstance Win32_ComputerSystem      -ErrorAction Stop
        $sticks    = @(Get-CimInstance Win32_PhysicalMemory    -ErrorAction Stop)
        $memArrays = @(Get-CimInstance Win32_PhysicalMemoryArray -ErrorAction SilentlyContinue)
        $t         = $Script:CFG.Thresholds

        $totalGB = ToGB $cs.TotalPhysicalMemory
        # OS always reports slightly less than physical RAM (hardware-reserved pages).
        # Round up to the nearest standard size if within 6% below it.
        $stdRamSizes = @(4,6,8,12,16,24,32,48,64,96,128)
        $effectiveGB = $totalGB
        foreach ($s in $stdRamSizes) {
            if ($totalGB -ge ($s * 0.94) -and $totalGB -lt $s) { $effectiveGB = $s; break }
        }
        $availGB = ToGB ($os.FreePhysicalMemory * 1KB)
        $usedGB  = [Math]::Round($totalGB - $availGB, 1)
        $usedPct = if ($totalGB -gt 0) { [Math]::Round(($usedGB / $totalGB) * 100, 0) } else { 0 }

        $speeds = $sticks | Where-Object { $_.Speed -gt 0 } | ForEach-Object { $_.Speed }
        $maxMHz = if ($speeds) { ($speeds | Measure-Object -Maximum).Maximum } else { 0 }

        $memType = switch ($sticks[0].SMBIOSMemoryType) {
            20 { 'DDR' }; 21 { 'DDR2' }; 24 { 'DDR3' }
            26 { 'DDR4' }; 27 { 'LPDDR' }; 28 { 'LPDDR2' }
            29 { 'DDR4' }  # Lenovo/some BIOS report DDR4 SO-DIMMs as type 29
            30 { 'LPDDR4' }; 31 { 'DDR4E' }; 32 { 'LPDDR4X' }
            34 { 'DDR5' }; 35 { 'LPDDR5' }
            default { "Type-$($sticks[0].SMBIOSMemoryType)" }
        }
        $slotsUsed  = ($sticks | Where-Object { $_.Capacity -gt 0 }).Count
        # Win32_PhysicalMemory only enumerates populated slots; use Win32_PhysicalMemoryArray for the real total
        $totalSlots = if ($memArrays) { ($memArrays | Measure-Object -Property MemoryDevices -Sum).Sum } else { 0 }

        $d.TotalGB       = $effectiveGB
        $d.AvailableGB   = [Math]::Round($availGB, 1)
        $d.UsedGB        = $usedGB
        $d.UsedPct       = "$usedPct%"
        $d.SpeedMHz      = $maxMHz
        $d.Type          = $memType
        $d.SlotsUsed     = $slotsUsed
        $d.TotalSlots    = if ($totalSlots -gt 0) { $totalSlots } else { $sticks.Count }
        $d.Modules       = ($sticks | Where-Object { $_.Capacity -gt 0 } |
                            ForEach-Object { "$($_.DeviceLocator): $(ToGB $_.Capacity) GB @ $($_.Speed) MHz $($_.Manufacturer.Trim())" }) -join ' | '

        if ($effectiveGB -ne $totalGB) { $d.OSReportedGB = $totalGB }
        Write-Log "$($effectiveGB) GB $memType @ $($maxMHz) MHz | Used $($usedGB)/$($effectiveGB) GB ($($usedPct)%)"

        $capScore  = Clamp100 (($effectiveGB / $t.RAM_RecommendedGB) * 100)
        $spdScore  = if ($maxMHz -gt 0) { Clamp100 (($maxMHz / $t.RAM_RecommendedMHz) * 100) } else { 50 }
        $useScore  = if ($usedPct -lt 80) { 100 } elseif ($usedPct -lt 90) { 70 } else { 40 }
        $score     = Clamp100 (($capScore * 0.6) + ($spdScore * 0.3) + ($useScore * 0.1))

        $displayGB = if ($effectiveGB -ne $totalGB) { "$effectiveGB GB (reported: $totalGB GB)" } else { "$totalGB GB" }
        if ($effectiveGB -lt $t.RAM_MinGB) {
            $issues.Add("$displayGB is below minimum $($t.RAM_MinGB) GB")
            $score = Clamp100 ($score - 20)
        } elseif ($effectiveGB -ge $t.RAM_RecommendedGB) {
            $good.Add("$displayGB meets recommended $($t.RAM_RecommendedGB) GB")
        } else {
            $issues.Add("$displayGB is below recommended $($t.RAM_RecommendedGB) GB"  )
        }

        if ($maxMHz -gt 0 -and $maxMHz -lt $t.RAM_MinSpeedMHz) {
            $issues.Add("$($maxMHz) MHz below minimum $($t.RAM_MinSpeedMHz) MHz")
        } elseif ($maxMHz -ge $t.RAM_RecommendedMHz) {
            $good.Add("$($maxMHz) MHz meets recommended speed")
        }

        if ($usedPct -gt 85) {
            $issues.Add("High usage $($usedPct)% - may indicate bloatware or memory pressure")
        }

        # Single-channel detection: if there are unused slots AND only one stick is filled,
        # memory bandwidth is roughly halved - a real performance hit on iGPU laptops.
        # Soldered memory typically reports totalSlots = 0 via Win32_PhysicalMemoryArray, so we skip.
        if ($totalSlots -gt 1 -and $slotsUsed -eq 1) {
            $issues.Add("Running single-channel - only 1 of $totalSlots RAM slots populated; memory bandwidth ~50% of dual-channel")
            $score = Clamp100 ($score - 8)
        }

        $status = Score-ToStatus $score
        $Script:RESULTS.RAM = New-Result -Module RAM -Score $score -Status $status `
            -Summary "$($effectiveGB) GB $memType @ $($maxMHz) MHz | $slotsUsed slot(s)" `
            -Data $d -Issues $issues -Good $good

        Write-Log "RAM Score: $score ($status)" 'OK'
    } catch {
        Write-Log "RAM error: $_" 'ERROR'
        $Script:RESULTS.RAM = New-Result -Module RAM -Score 0 -Status ERROR -Summary "Error: $_"
    }
}

# ============================================================
#  MODULE: STORAGE
# ============================================================

function Test-Storage {
    Write-Section 'Storage Assessment'
    $issues    = [System.Collections.Generic.List[string]]::new()
    $good      = [System.Collections.Generic.List[string]]::new()
    $d         = [ordered]@{}
    $driveDocs = [System.Collections.Generic.List[object]]::new()
    $scores    = [System.Collections.Generic.List[int]]::new()

    try {
        $t         = $Script:CFG.Thresholds
        $diskDrvs  = @(Get-CimInstance Win32_DiskDrive -ErrorAction Stop)
        $physDisks = @(Get-PhysicalDisk -ErrorAction SilentlyContinue)
        $vols      = @(Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue)

        # Call SMART once for all drives; per-iteration calls returned the union result for every drive
        $smartAll = Read-DiskSMART
        if ($smartAll.Source -eq 'smartctl' -and $smartAll.Drives -and $smartAll.Drives.Count -gt 0) {
            foreach ($drv in $smartAll.Drives) {
                $lvl = if ($drv.Fail) { 'WARN' } else { 'OK' }
                Write-Log ("SMART [$($drv.Path)]: $($drv.Model) | Health: $($drv.Health) | " +
                           "Hours: $($drv.PowerOnHours) | Realloc: $($drv.Reallocated) Pending: $($drv.Pending)") $lvl
            }
        } else {
            Write-Log "SMART ($($smartAll.Source)): $($smartAll.Note)" $(if ($smartAll.Fail) { 'WARN' } else { 'INFO' })
        }

        $primaryDriveSeen = $false
        foreach ($dd in $diskDrvs) {
            $dd_d = [ordered]@{}
            $dd_d.Name       = $dd.Model.Trim()
            $dd_d.CapGB      = ToGB $dd.Size
            $dd_d.Interface  = $dd.InterfaceType
            $dd_d.Serial     = $dd.SerialNumber.Trim()
            $dd_d.Status     = $dd.Status

            # Get media/bus type from PhysicalDisk (more accurate)
            $modelPfx = $dd.Model.Trim()
            if ($modelPfx.Length -gt 15) { $modelPfx = $modelPfx.Substring(0,15) }
            $phys = $physDisks | Where-Object { $_.FriendlyName -like "*$modelPfx*" } | Select-Object -First 1
            if ($phys) {
                $dd_d.MediaType  = $phys.MediaType
                $dd_d.BusType    = $phys.BusType
                $dd_d.Health     = $phys.HealthStatus
            } else {
                $dd_d.BusType    = if ($dd.Model -match 'NVMe|NVME|M\.2') { 'NVMe' }
                                    elseif ($dd.InterfaceType -eq 'USB')   { 'USB' }
                                    else { $dd.InterfaceType }
                $dd_d.MediaType  = if ($dd.Model -match 'SSD|NVMe|Solid') { 'SSD' } else { 'HDD' }
                $dd_d.Health     = $dd.Status
            }

            # USB drives are removable and shouldn't influence the laptop's storage score
            if ($dd_d.BusType -eq 'USB') {
                $dd_d.SMART = 'N/A (USB removable)'
                $driveDocs.Add($dd_d)
                Write-Log "USB drive detected (excluded from score): $($dd_d.Name) - $($dd_d.CapGB) GB"
                continue
            }

            $isPrimary = -not $primaryDriveSeen
            if ($isPrimary) { $primaryDriveSeen = $true }

            # Convert OS-reported capacity (GiB) back to manufacturer GB (10^9 bytes / 2^30 = 1.0737)
            # Then snap to nearest standard drive size within 6% tolerance
            $manuGB       = [Math]::Round($dd_d.CapGB * 1.0737, 0)
            $stdStorSizes = @(64,120,128,240,256,480,512,960,1000,1024,2000,2048,4000,4096)
            $nearestStd   = $stdStorSizes | Sort-Object { [Math]::Abs($_ - $manuGB) } | Select-Object -First 1
            $effectiveCapGB = if ([Math]::Abs($nearestStd - $manuGB) -le ($nearestStd * 0.06)) { $nearestStd } else { $dd_d.CapGB }
            $dd_d.EffectiveCapGB = $effectiveCapGB

            Write-Log "Disk: $($dd_d.Name) - $($dd_d.CapGB) GB (effective: $effectiveCapGB GB) [$($dd_d.BusType)/$($dd_d.MediaType)]"

            # Per-drive SMART: match this WMI drive to its smartctl entry by serial then model prefix
            if ($smartAll.Source -eq 'smartctl' -and $smartAll.Drives -and $smartAll.Drives.Count -gt 0) {
                $ddSerial = $dd.SerialNumber.Trim()
                $ddModelL = $dd.Model.Trim().ToLower()
                $drvEntry = $smartAll.Drives | Where-Object {
                    ($ddSerial.Length -gt 3 -and $_.Serial -eq $ddSerial) -or
                    ($ddModelL.Length -gt 5 -and $_.Model.ToLower() -like "*$($ddModelL.Substring(0,[Math]::Min(15,$ddModelL.Length)))*")
                } | Select-Object -First 1
                $smart = if ($drvEntry) {
                    [ordered]@{ Fail=$drvEntry.Fail; Inconclusive=$drvEntry.Inconclusive
                                Note="$($drvEntry.Model) [$($drvEntry.Health)]"; Source='smartctl'; Drives=@($drvEntry) }
                } else {
                    [ordered]@{ Fail=$false; Inconclusive=$true
                                Note='Drive not matched in smartctl output (possible RAID alias)'; Source='smartctl'; Drives=@() }
                }
            } else {
                $smart = $smartAll
            }

            $dd_d.SMART = if ($smart.Fail) { 'FAILURE PREDICTED' }
                          elseif ($smart.Inconclusive) { 'UNKNOWN (SMART unreadable - likely behind RAID controller)' }
                          else { $smart.Note }
            $dd_d.SMART_Source = $smart.Source
            if ($smart.Fail) { $issues.Add("SMART failure predicted: $($dd_d.Name)") }
            elseif ($smart.Inconclusive) {
                $issues.Add("SMART unreadable on $($dd_d.Name) (likely behind RAID controller) - manually verify drive health")
            }

            # Per-disk score
            $ds = 70
            if ($dd_d.BusType -eq 'NVMe')                         { $ds += $t.Storage_NVMe_Bonus }
            if ($dd_d.MediaType -eq 'SSD')                         { $ds += 10 }
            if ($dd_d.Health -eq 'Healthy')                        { $ds += 5 }
            if ($smart.Fail)                                        { $ds -= 40 }
            if ($effectiveCapGB -ge $t.Storage_RecommendedGB)      { $ds += 5 }
            elseif ($effectiveCapGB -lt $t.Storage_MinGB)          { $ds -= 20 }

            # HDD primary drive - heavy penalty + repair recommendation (primary only; secondary HDDs are scored but not flagged).
            # Modern Windows + Office + Teams runs at acceptable speed only on SSD.
            if ($isPrimary -and ($dd_d.MediaType -eq 'HDD' -or $dd_d.MediaType -eq 'Unspecified')) {
                if ($dd_d.MediaType -eq 'HDD' -or $dd.Model -notmatch 'SSD|NVMe|Solid|M\.2') {
                    $ds -= 25
                    $issues.Add("Primary drive is HDD ($($dd_d.Name)) - Windows 11 + browser/office workload runs unacceptably slow on spinning disks")
                }
            }

            # Drive age via smartctl power-on hours - informational, not "good"
            if ($smart.Source -eq 'smartctl' -and $smart.Drives.Count -gt 0) {
                $maxHours = ($smart.Drives | ForEach-Object { [int]$_.PowerOnHours } | Measure-Object -Maximum).Maximum
                if ($maxHours -gt 80000) {
                    $ds -= 10
                    $issues.Add("Drive has $maxHours power-on hours (~$([Math]::Round($maxHours/8760,1)) yrs of 24/7 use) - heavily worn, plan replacement")
                } elseif ($maxHours -gt 50000) {
                    $issues.Add("Drive has $maxHours power-on hours (~$([Math]::Round($maxHours/8760,1)) yrs of 24/7 use) - aging, monitor for replacement")
                }
            }

            $scores.Add((Clamp100 $ds))

            $driveDocs.Add($dd_d)
        }

        $d.Drives = $driveDocs

        # Volume usage
        $volList = [System.Collections.Generic.List[string]]::new()
        foreach ($v in $vols) {
            $freeGB  = ToGB $v.FreeSpace
            $totGB   = ToGB $v.Size
            $usedPct = if ($v.Size -gt 0) { [Math]::Round((($v.Size - $v.FreeSpace) / $v.Size) * 100, 0) } else { 0 }
            $volList.Add("$($v.DeviceID) - $($totGB) GB ($($freeGB) GB free, $($usedPct)% used)")
            if ($usedPct -gt 90) { $issues.Add("Drive $($v.DeviceID) is $($usedPct)% full") }
        }
        $d.Volumes = $volList

        # DiskSpd benchmark if available (does not require admin)
        if ($Script:TOOLS.DiskSpd) {
            Write-Log 'Running DiskSpd benchmark (~30s)...'
            $bench = Invoke-DiskSpd
            if ($bench) {
                $d.ReadMBs  = $bench.Read
                $d.WriteMBs = $bench.Write
                Write-Log "DiskSpd Read $($bench.Read) MB/s | Write $($bench.Write) MB/s"

                if ($bench.Read -lt $t.Storage_MinReadMBs) {
                    $issues.Add("Read $($bench.Read) MB/s below threshold $($t.Storage_MinReadMBs) MB/s")
                    for ($i = 0; $i -lt $scores.Count; $i++) { $scores[$i] = Clamp100 ($scores[$i] - 10) }
                } else {
                    $good.Add("Read $($bench.Read) MB/s passes $($t.Storage_MinReadMBs) MB/s threshold")
                }
                if ($bench.Write -lt $t.Storage_MinWriteMBs) {
                    $issues.Add("Write $($bench.Write) MB/s below threshold $($t.Storage_MinWriteMBs) MB/s")
                } else {
                    $good.Add("Write $($bench.Write) MB/s passes $($t.Storage_MinWriteMBs) MB/s threshold")
                }
            }
        } else {
            $d.ReadMBs  = 'N/A'
            $d.WriteMBs = 'N/A'
            $ws = Read-WinSATScore 'Disk'
            if ($ws -gt 0) { $d.WinSAT_Disk = $ws; Write-Log "WinSAT Disk: $ws" }
        }

        $primary = $driveDocs | Where-Object { $_.BusType -ne 'USB' } | Select-Object -First 1
        $primStr = if ($primary) { "$($primary.Name) - $($primary.CapGB) GB [$($primary.BusType)/$($primary.MediaType)]" } else { 'No primary disk detected' }

        $primCap = if ($primary.EffectiveCapGB) { $primary.EffectiveCapGB } else { $primary.CapGB }
        if ($primary -and $primCap -ge $t.Storage_RecommendedGB) {
            $good.Add("Primary disk $primCap GB meets recommended $($t.Storage_RecommendedGB) GB")
        } elseif ($primary -and $primCap -lt $t.Storage_MinGB) {
            $issues.Add("Primary disk $primCap GB below minimum $($t.Storage_MinGB) GB")
        } else {
            if ($primary) { $good.Add("Primary disk $primCap GB meets minimum requirement") }
        }

        $avg = if ($scores.Count -gt 0) { Clamp100 (($scores | Measure-Object -Average).Average) } else { 0 }

        $status = Score-ToStatus $avg
        $Script:RESULTS.Storage = New-Result -Module Storage -Score $avg -Status $status `
            -Summary $primStr -Data $d -Issues $issues -Good $good

        Write-Log "Storage Score: $avg ($status)" 'OK'
    } catch {
        Write-Log "Storage error: $_" 'ERROR'
        $Script:RESULTS.Storage = New-Result -Module Storage -Score 0 -Status ERROR -Summary "Error: $_"
    }
}

function Read-DiskSMART {
    # Prefer smartctl for real attribute data; fall back to WMI prediction flag
    if ($Script:TOOLS.Smartctl) {
        return Read-DiskSMART-Smartctl
    }
    $r = @{ Fail = $false; Inconclusive = $false; Note = 'OK'; Source = 'WMI' }
    try {
        $p = Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction Stop
        foreach ($item in $p) {
            if ($item.PredictFailure) { $r.Fail = $true; $r.Note = "Reason code $($item.Reason)"; break }
        }
    } catch { $r.Note = 'N/A (requires admin)' }
    return $r
}

function Read-DiskSMART-Smartctl {
    $exe    = $Script:TOOLS.Smartctl
    $drives = [System.Collections.Generic.List[object]]::new()
    $anyFail = $false

    try {
        # Enumerate physical drives via smartctl --scan
        $scan = (& $exe --scan 2>&1) -join "`n"
        $drivePaths = @()
        foreach ($line in ($scan -split "`n")) {
            if ($line -match '^(/dev/\S+)') { $drivePaths += $Matches[1] }
        }
        # Fallback: try /dev/pd0 through /dev/pd3 if scan returned nothing
        if ($drivePaths.Count -eq 0) { $drivePaths = '/dev/pd0','/dev/pd1','/dev/pd2','/dev/pd3' }

        foreach ($dp in $drivePaths) {
            $out = (& $exe -a $dp 2>&1) -join "`n"
            if ($out -match 'No such device|Unable to detect|Smartctl open device') { continue }

            $info = [ordered]@{
                Path         = $dp
                Model        = ''
                Serial       = ''
                Health       = 'UNKNOWN'
                PowerOnHours = 0
                TempC        = 0
                Reallocated  = 0
                Pending      = 0
                Uncorrectable= 0
                Fail         = $false
            }

            # ATA/SATA model; NVMe uses "Model Number:"
            if ($out -match 'Device Model:\s+(.+)')       { $info.Model  = $Matches[1].Trim() }
            elseif ($out -match 'Model Number:\s+(.+)')   { $info.Model  = $Matches[1].Trim() }
            if ($out -match 'Serial Number:\s+(.+)')      { $info.Serial = $Matches[1].Trim() }
            if ($out -match 'overall-health.*?:\s+(\w+)') { $info.Health = $Matches[1].Trim() }
            # ATA attribute format: "Power_On_Hours  0x... 100 100 000 ... 5432"
            if ($out -match 'Power_On_Hours\s+\S+\s+\S+\s+\S+\s+\S+\s+(\d+)') {
                $info.PowerOnHours = [int]$Matches[1]
            }
            # NVMe log format: "Power On Hours:                     5432"
            elseif ($out -match 'Power On Hours:\s+(\d[\d,]+)') {
                $info.PowerOnHours = [int]($Matches[1] -replace ',','')
            }
            # ATA temp: "Temperature_Celsius  0x... 100 39 ..."  NVMe: "Temperature:  39 Celsius"
            if ($out -match 'Temperature_Celsius\s+\S+\s+(\d+)')  { $info.TempC = [int]$Matches[1] }
            elseif ($out -match 'Temperature:\s+(\d+) Celsius')   { $info.TempC = [int]$Matches[1] }
            if ($out -match 'Reallocated_Sector_Ct\s+\S+\s+\S+\s+\S+\s+\S+\s+(\d+)') { $info.Reallocated = [int]$Matches[1] }
            if ($out -match 'Current_Pending_Sector\s+\S+\s+\S+\s+\S+\s+\S+\s+(\d+)') { $info.Pending = [int]$Matches[1] }
            if ($out -match 'Offline_Uncorrectable\s+\S+\s+\S+\s+\S+\s+\S+\s+(\d+)')  { $info.Uncorrectable = [int]$Matches[1] }

            # Only flag a real failure: explicit FAILED health OR reallocated/pending/uncorrectable sectors > 0.
            # UNKNOWN (= unparseable, common behind RAID controllers) is NOT a failure - it's "we couldn't tell".
            $info.Fail         = ($info.Health -eq 'FAILED' -or $info.Health -match 'FAIL') -or
                                 ($info.Reallocated -gt 0) -or ($info.Pending -gt 0) -or ($info.Uncorrectable -gt 0)
            $info.Inconclusive = ($info.Health -eq 'UNKNOWN')
            if ($info.Fail) { $anyFail = $true }
            $drives.Add($info)
        }
    } catch {
        Write-Log "smartctl error: $_" 'WARN'
    }

    $summary = if ($drives.Count -gt 0) {
        ($drives | ForEach-Object {
            "$($_.Model) [$($_.Health)] Realloc:$($_.Reallocated) Pending:$($_.Pending) Hours:$($_.PowerOnHours)"
        }) -join '; '
    } else { 'No drives enumerated' }

    # Inconclusive = ALL parsed drives returned UNKNOWN (typical for hardware RAID controllers).
    # Distinct from Fail; we don't penalize the score, just note SMART couldn't be read.
    $allInconclusive = ($drives.Count -gt 0) -and -not ($drives | Where-Object { -not $_.Inconclusive })
    return @{
        Fail         = $anyFail
        Inconclusive = (-not $anyFail) -and $allInconclusive
        Note         = $summary
        Drives       = $drives
        Source       = 'smartctl'
    }
}

function Invoke-DiskSpd {
    try {
        $exe  = $Script:TOOLS.DiskSpd
        $file = Join-Path $Script:CFG.Tools.WorkDir 'dstest.dat'

        # Sequential read: 256MB file, 128KB blocks, 15s, 2 threads, 4 outstanding IOs
        # -c creates the test file; no -S flags to avoid compat issues across DiskSpd versions
        $rOut = (& $exe '-b128K' '-d15' '-o4' '-t2' '-c268435456' $file 2>&1) -join "`n"
        $read  = Parse-DiskSpdMBs $rOut

        # Sequential write: reuse existing file, 100% write workload
        $wOut  = (& $exe '-b128K' '-d15' '-o4' '-t2' '-w100' $file 2>&1) -join "`n"
        $write = Parse-DiskSpdMBs $wOut

        Remove-Item $file -ErrorAction SilentlyContinue
        return @{ Read = $read; Write = $write }
    } catch {
        Write-Log "DiskSpd failed: $_" 'WARN'
        return $null
    }
}

function Parse-DiskSpdMBs {
    param([string]$Out)
    # DiskSpd actual output format (pipe-separated columns):
    #   total:   BYTES | IOs | MiB/s | I/O per s | AvgLat | ...
    # The 3rd column (index 2) is MiB/s  -  no unit text on the line.
    foreach ($l in ($Out -split "`n")) {
        if ($l.Trim() -match '^total:') {
            $parts = $l -split '\|'
            if ($parts.Count -ge 3 -and $parts[2].Trim() -match '([\d]+\.[\d]*)') {
                return [int][Math]::Round([double]$Matches[1], 0)
            }
        }
    }
    return 0
}

# ============================================================
#  MODULE: BATTERY
# ============================================================

function Test-Battery {
    Write-Section 'Battery Assessment'

    if (-not $Script:CAPS.HasBattery) {
        Write-Log 'No battery - desktop or non-applicable device. Skipping.' 'INFO'
        $Script:RESULTS.Battery = New-Result -Module Battery -Score 100 -Status SKIP `
            -Summary 'No battery (desktop/server - not scored)'
        return
    }

    $issues = [System.Collections.Generic.List[string]]::new()
    $good   = [System.Collections.Generic.List[string]]::new()
    $d      = [ordered]@{}

    try {
        $t    = $Script:CFG.Thresholds
        $batt = Get-CimInstance Win32_Battery -ErrorAction Stop | Select-Object -First 1

        $chemStr = switch ($batt.Chemistry) {
            1 {'Other'}; 2 {'Unknown'}; 3 {'Lead Acid'}; 4 {'Nickel Cadmium'}
            5 {'Nickel Metal Hydride'}; 6 {'Lithium Ion'}; 7 {'Zinc Air'}; 8 {'Lithium Polymer'}
            default { "Type $($batt.Chemistry)" }
        }
        $statusStr = switch ($batt.BatteryStatus) {
            1 {'Discharging'}; 2 {'AC/Charging'}; 3 {'Fully Charged'}
            4 {'Low'}; 5 {'Critical'}; 7 {'Charging'}
            default { "Code $($batt.BatteryStatus)" }
        }

        $d.Name        = $batt.Name
        $d.Chemistry   = $chemStr
        $d.ChargedPct  = "$($batt.EstimatedChargeRemaining)%"
        $d.Status      = $statusStr
        $d.Manufacturer = if ($batt.Manufacturer) { $batt.Manufacturer.Trim() } else { 'Unknown' }

        Write-Log "Battery: $($d.Chemistry) | $($d.ChargedPct) | $($d.Status)"

        # powercfg battery report for health
        $health = Read-BatteryReport
        if ($health) {
            $d.DesignWh      = $health.DesignWh
            $d.FullChargeWh  = $health.FullWh
            $d.HealthPct     = "$($health.Health)%"
            $d.CycleCount    = $health.Cycles
            $d.MfgDate       = $health.MfgDate
            Write-Log "Health $($health.Health)% | Design $($health.DesignWh) Wh | Full $($health.FullWh) Wh | Cycles $($health.Cycles)"
        } else {
            $d.HealthPct  = 'N/A'
            $d.CycleCount = 'N/A'
        }

        $hp  = if ($health) { $health.Health } else { 80 }  # assume ok if cant read
        $cwh = if ($health -and $health.FullWh -gt 0) { $health.FullWh } else { 50 }

        $hScore = Clamp100 $hp
        $cScore = Clamp100 (($cwh / $t.Battery_MinCapWh) * 80)
        $score  = Clamp100 (($hScore * 0.65) + ($cScore * 0.35))

        if ($hp -lt $t.Battery_MinHealth) {
            $issues.Add("Health $($hp)% below minimum $($t.Battery_MinHealth)%")
        } elseif ($hp -lt $t.Battery_WarnHealth) {
            $issues.Add("Health $($hp)% degrading (warn < $($t.Battery_WarnHealth)%)")
        } else {
            $good.Add("Battery health $($hp)% is acceptable")
        }

        if ($health -and $health.FullWh -lt $t.Battery_MinCapWh) {
            $issues.Add("Full charge $($health.FullWh) Wh below minimum $($t.Battery_MinCapWh) Wh")
        }
        if ($health -and $health.Cycles -gt 500) {
            $issues.Add("High cycle count ($($health.Cycles)) - may need replacement soon")
        } elseif ($health -and $health.Cycles -le 200) {
            $good.Add("Low cycle count ($($health.Cycles))")
        }

        $status = Score-ToStatus $score
        $Script:RESULTS.Battery = New-Result -Module Battery -Score $score -Status $status `
            -Summary "$($d.Chemistry) | Health $($d.HealthPct) | $($d.ChargedPct) | $($d.Status)" `
            -Data $d -Issues $issues -Good $good

        Write-Log "Battery Score: $score ($status)" 'OK'
    } catch {
        Write-Log "Battery error: $_" 'ERROR'
        $Script:RESULTS.Battery = New-Result -Module Battery -Score 0 -Status ERROR -Summary "Error: $_"
    }
}

function Read-BatteryReport {
    try {
        $out = Join-Path $Script:CFG.Tools.WorkDir 'bat.xml'
        $null = & powercfg /batteryreport /output $out /xml 2>&1
        if (-not (Test-Path $out)) { return $null }
        [xml]$x = Get-Content $out -ErrorAction Stop
        $b = $x.BatteryReport.Batteries.Battery | Select-Object -First 1
        if (-not $b) { return $null }
        $designMwh = [long]$b.DesignCapacity
        $fullMwh   = [long]$b.FullChargeCapacity
        $cycles    = [int]$b.CycleCount
        return @{
            DesignWh = [Math]::Round($designMwh / 1000, 1)
            FullWh   = [Math]::Round($fullMwh   / 1000, 1)
            Health   = if ($designMwh -gt 0) { [int](($fullMwh / $designMwh) * 100) } else { 0 }
            Cycles   = $cycles
            MfgDate  = if ($b.ManufactureDate) { $b.ManufactureDate } else { 'Unknown' }
        }
    } catch {
        Write-Log "powercfg report failed: $_" 'WARN'
        return $null
    }
}

# ============================================================
#  MODULE: GPU & DISPLAY
# ============================================================

# Dead pixel test: full-screen WinForms color cycle on a guaranteed STA runspace.
# Modern PowerShell hosts (Windows Terminal etc.) can run in MTA - Form.ShowDialog()
# silently fails on MTA threads. Forcing a new STA runspace makes this work everywhere.
function Invoke-DeadPixelTest {
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $colors     = @(
            [System.Drawing.Color]::Black,
            [System.Drawing.Color]::White,
            [System.Drawing.Color]::Red,
            [System.Drawing.Color]::Lime,
            [System.Drawing.Color]::Blue
        )
        $colorNames = @('BLACK (stuck pixels appear bright)', 'WHITE (dead pixels appear dark)', 'RED', 'GREEN', 'BLUE')
        $script:idx = 0

        $form                 = New-Object System.Windows.Forms.Form
        $form.FormBorderStyle = 'None'
        $form.WindowState     = 'Maximized'
        $form.TopMost         = $true
        $form.BackColor       = $colors[0]
        $form.Cursor          = [System.Windows.Forms.Cursors]::Cross
        $form.KeyPreview      = $true

        $lbl                  = New-Object System.Windows.Forms.Label
        $lbl.Text             = "Color: $($colorNames[0]) (1/5)  |  SPACE = next color  |  ESC = done"
        $lbl.Font             = New-Object System.Drawing.Font('Consolas', 11)
        $lbl.ForeColor        = [System.Drawing.Color]::FromArgb(90, 90, 90)
        $lbl.BackColor        = [System.Drawing.Color]::Transparent
        $lbl.AutoSize         = $true
        $lbl.Location         = [System.Drawing.Point]::new(12, [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height - 36)
        $form.Controls.Add($lbl)

        $form.Add_KeyDown({
            param($s, $e)
            if     ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $s.Close() }
            elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Space)  {
                $script:idx = ($script:idx + 1) % 5
                $i = $script:idx
                $s.BackColor   = $colors[$i]
                $lbl.ForeColor = if ($i -in 0,4) { [System.Drawing.Color]::FromArgb(90,90,90) } else { [System.Drawing.Color]::FromArgb(60,60,60) }
                $lbl.Text      = "Color: $($colorNames[$i]) ($($i+1)/5)  |  SPACE = next color  |  ESC = done"
            }
        })

        $form.Add_Shown({ $form.Activate() })
        [void]$form.ShowDialog()
        $form.Dispose()
    })
    try { [void]$ps.Invoke() } catch { Write-Log "Dead-pixel UI error: $_" 'WARN' }
    $ps.Dispose()
    $rs.Close()
    $rs.Dispose()
}

# nvidia-smi: free CLI tool shipped with every NVIDIA driver - no external download needed.
# Returns GPU sensor data directly from the driver; works unattended, no GUI required.
# This is the only live-GPU sensor source the script uses.
function Get-NvidiaSmiData {
    $nvSmi = Get-Command 'nvidia-smi.exe' -ErrorAction SilentlyContinue
    if (-not $nvSmi) {
        $nvSmi = @(
            "$env:ProgramFiles\NVIDIA Corporation\NVSMI\nvidia-smi.exe",
            "$env:SystemRoot\System32\nvidia-smi.exe"
        ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    }
    if (-not $nvSmi) { return $null }
    $exe = if ($nvSmi -is [string]) { $nvSmi } else { $nvSmi.Source }
    try {
        $out = & $exe --query-gpu=temperature.gpu,utilization.gpu,clocks.current.graphics,clocks.current.memory,memory.used,memory.total --format=csv,noheader,nounits 2>$null
        if (-not $out -or $out -match 'error|FAILED') { return $null }
        $v = ($out.Trim() -split ',') | ForEach-Object { $_.Trim() }
        if ($v.Count -lt 6) { return $null }
        return @{ Temp=$([int]$v[0]); Load=$([int]$v[1]); Clock=$([int]$v[2]); MemClock=$([int]$v[3]); VRAMUsed=$([int]$v[4]); VRAMTotal=$([int]$v[5]); Tool='nvidia-smi' }
    } catch { return $null }
}

function Test-GPU {
    Write-Section 'GPU and Display Assessment'
    $issues = [System.Collections.Generic.List[string]]::new()
    $good   = [System.Collections.Generic.List[string]]::new()
    $d      = [ordered]@{}

    try {
        $t    = $Script:CFG.Thresholds
        $gpus = @(Get-CimInstance Win32_VideoController -ErrorAction Stop |
                  Where-Object { $_.Name -notmatch 'Microsoft|Basic Display' })

        $gpuList  = [System.Collections.Generic.List[object]]::new()
        $topScore = 0

        foreach ($g in $gpus) {
            $vram  = if ($g.AdapterRAM -gt 0) { ToMB $g.AdapterRAM } else { 0 }
            $isDisc = ($g.Name -match 'NVIDIA|AMD|Radeon|GeForce|Quadro') -and $vram -gt 512
            # Modern integrated GPU = sufficient for office work (Office, Teams, dual monitors,
            # 4K video playback). Don't penalize laptops just for not having a discrete GPU.
            $isModernIGPU = (-not $isDisc) -and (
                $g.Name -match 'UHD Graphics|Iris|HD Graphics 6\d{2}|HD Graphics [78]\d{2}|HD Graphics [5]\d{2}0' -or
                $g.Name -match 'Radeon\s+(?:RX )?Vega|Radeon Graphics|Radeon\(TM\) Graphics|AMD Graphics' -or
                $g.Name -match 'Apple M\d'
            )
            $dDate = 'Unknown'
            $dDays = 9999
            if ($g.DriverDate) {
                try {
                    $dDate = $g.DriverDate.ToString('yyyy-MM-dd')
                    $dDays = [int]((Get-Date) - $g.DriverDate).TotalDays
                } catch {}
            }

            # Discrete GPU starts at 60 + bonus. Modern iGPU starts at 75 (good enough for office).
            # Old/unknown integrated stays at 60.
            $gs = if ($isModernIGPU) { 75 } else { 60 }
            if ($isDisc)                          { $gs += $t.GPU_DiscreteBonus }
            if ($vram -ge 4096)                   { $gs += 10 }
            elseif ($vram -ge 2048)               { $gs += 5 }
            # VRAM penalty doesn't apply to integrated GPUs - they share system RAM dynamically
            if ($vram -lt $t.GPU_MinVRAM_MB -and -not ($isModernIGPU -or (-not $isDisc))) { $gs -= 10 }
            if ($dDays -gt 365)                   { $gs -= 5; $issues.Add("Driver for '$($g.Name)' is $([Math]::Round($dDays/365,1))y old") }
            elseif ($dDays -lt 180)               { $good.Add("Driver for '$($g.Name)' is recent ($dDays days)") }
            if ($isDisc)                          { $good.Add("Discrete GPU: $($g.Name)") }
            elseif ($isModernIGPU)                { $good.Add("Modern integrated GPU: $($g.Name) - sufficient for office workloads") }
            if ($gs -gt $topScore)                { $topScore = $gs }

            $gpuList.Add([ordered]@{
                Name       = $g.Name
                VRAM_MB    = $vram
                Driver     = $g.DriverVersion
                DriverDate = $dDate
                Status     = $g.Status
                Discrete   = $isDisc
                Modern_iGPU = $isModernIGPU
            })
        }

        if ($gpuList.Count -eq 0) {
            $issues.Add('No compatible video controller detected')
            $topScore = 30
        }
        $d.GPUs = $gpuList

        # GPU live sensor readings: nvidia-smi only (NVIDIA cards). Other GPUs are scored
        # purely on what WMI reports: card model, VRAM, driver date.
        $gpuSensors = $null
        $nvData = Get-NvidiaSmiData
        if ($nvData) {
            $gpuSensors = $nvData
            Write-Log ("nvidia-smi: Temp $($nvData.Temp)C  Load $($nvData.Load)%  Clock $($nvData.Clock)MHz") 'OK'
        }
        if ($gpuSensors) {
            if ($gpuSensors.Temp     -gt 0) { $d.GPU_Temp      = "$($gpuSensors.Temp) C" }
            if ($gpuSensors.Load     -ge 0) { $d.GPU_Load      = "$($gpuSensors.Load)%" }
            if ($gpuSensors.Clock    -gt 0) { $d.GPU_Clock     = "$($gpuSensors.Clock) MHz" }
            if ($gpuSensors.MemClock -gt 0) { $d.GPU_MemClock  = "$($gpuSensors.MemClock) MHz" }
            if ($gpuSensors.VRAMUsed -gt 0) { $d.GPU_VRAM_Used = "$($gpuSensors.VRAMUsed) MB" }
            $d.GPU_Sensor_Source = $gpuSensors.Tool
            if ($gpuSensors.Temp -gt 90) {
                $issues.Add("GPU idle temperature $($gpuSensors.Temp)C is critically high - thermal issue")
                $topScore -= 15
            }
        }

        # Physical resolution from Win32_VideoController (unaffected by DPI scaling).
        # System.Windows.Forms.Screen.Bounds returns logical pixels which shrink at >100% DPI.
        try {
            $resW = 0; $resH = 0
            if ($gpus.Count -gt 0 -and $gpus[0].CurrentHorizontalResolution -gt 0) {
                $resW = [int]$gpus[0].CurrentHorizontalResolution
                $resH = [int]$gpus[0].CurrentVerticalResolution
            }
            if ($resW -gt 0) {
                $resStr      = "${resW}x${resH}"
                $d.PrimaryRes = $resStr
                if ($resW -ge 3840) {
                    $good.Add("4K resolution ($resStr)")
                } elseif ($resW -ge 2560) {
                    $good.Add("QHD resolution ($resStr)")
                } elseif ($resW -ge 1920) {
                    $good.Add("Full HD resolution ($resStr)")
                } else {
                    $issues.Add("Resolution $resStr below 1920x1080")
                }
            }
            # Screen list for reference (device names via Forms, resolution from WMI above)
            try {
                Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
                $scrList = [System.Collections.Generic.List[string]]::new()
                foreach ($s in [System.Windows.Forms.Screen]::AllScreens) {
                    $scrList.Add("$($s.DeviceName) Primary:$($s.Primary)")
                }
                $d.Displays = $scrList
            } catch { $d.Displays = 'N/A' }
        } catch { $d.Displays = 'N/A' }

        # Refresh rate  -  snap to nearest standard rate within 1 Hz
        # (Windows reports 59 for 59.94 Hz panels marketed as 60 Hz)
        if ($gpus.Count -gt 0 -and $gpus[0].CurrentRefreshRate) {
            $hz  = [int]$gpus[0].CurrentRefreshRate
            $std = @(24,25,30,48,50,60,72,75,90,100,120,144,165,180,240,360)
            $nearest = $std | Sort-Object { [Math]::Abs($_ - $hz) } | Select-Object -First 1
            if ([Math]::Abs($nearest - $hz) -le 1) { $hz = $nearest }
            $d.RefreshHz = $hz
            if ($hz -ge 120) { $good.Add("High refresh rate display: $($hz) Hz") }
            elseif ($hz -lt $t.Display_MinRefreshHz) { $issues.Add("Refresh rate $($hz) Hz below minimum $($t.Display_MinRefreshHz) Hz") }
            else { $good.Add("Refresh rate $($hz) Hz") }
        }

        $ws = Read-WinSATScore 'Graphics'
        if ($ws -gt 0) { $d.WinSAT_Graphics = $ws }

        # Screen quality check: launch full-screen WinForms color cycle, then ask for results
        if (-not $Script:FullAuto) {
            Write-Host ''
            Write-Host '  === SCREEN QUALITY CHECK ===' -ForegroundColor Yellow
            Write-Host '  A full-screen color display will open.' -ForegroundColor Cyan
            Write-Host '  Cycle through all 5 colors (SPACE) looking for dead/stuck pixels, then close (ESC).' -ForegroundColor Cyan
            Write-Host '  Press Enter to start...' -ForegroundColor DarkGray -NoNewline
            [void](Read-Host)

            Invoke-DeadPixelTest

            if (Read-YesNo '  Did you see any dead or stuck pixels?') {
                Write-Host '     Approximate count: ' -ForegroundColor Yellow -NoNewline
                $dpCount = (Read-Host).Trim()
                $d.Dead_Pixels = "Reported: ~$(if($dpCount){$dpCount}else{'?'}) pixel(s)"
                $issues.Add("Screen: dead/stuck pixel(s) found during color test")
                $topScore -= 10
            } else {
                $d.Dead_Pixels = 'None (verified via color test)'
                $good.Add('Screen: no dead pixels found in color test')
            }

            if (Read-YesNo '  Was color/brightness even across the screen?') {
                $d.Screen_Quality = 'Good (user verified)'
                $good.Add('Screen: colors and brightness uniform')
            } else {
                Write-Host '     Describe the issue: ' -ForegroundColor Yellow -NoNewline
                $colorDesc = (Read-Host).Trim()
                $d.Screen_Quality = "Issue: $(if($colorDesc){$colorDesc}else{'reported'})"
                $issues.Add("Screen quality issue: $(if($colorDesc){$colorDesc}else{'reported by user'})")
                $topScore -= 5
            }
        }

        $score  = Clamp100 $topScore
        $status = Score-ToStatus $score
        $primStr = if ($gpuList.Count -gt 0) { $gpuList[0].Name } else { 'None' }

        $Script:RESULTS.GPU = New-Result -Module GPU -Score $score -Status $status `
            -Summary "$primStr | $($d.PrimaryRes)" -Data $d -Issues $issues -Good $good

        Write-Log "GPU Score: $score ($status)" 'OK'
    } catch {
        Write-Log "GPU error: $_" 'ERROR'
        $Script:RESULTS.GPU = New-Result -Module GPU -Score 0 -Status ERROR -Summary "Error: $_"
    }
}

# ============================================================
#  MODULE: NETWORK
# ============================================================

function Test-Network {
    Write-Section 'Network Assessment'
    $issues = [System.Collections.Generic.List[string]]::new()
    $good   = [System.Collections.Generic.List[string]]::new()
    $d      = [ordered]@{}

    try {
        $t = $Script:CFG.Thresholds

        $adapters = @(Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction Stop)
        $adapList = [System.Collections.Generic.List[string]]::new()
        foreach ($a in $adapters) {
            $ip = ($a.IPAddress | Where-Object { $_ -notmatch ':' }) -join ', '
            $adapList.Add("$($a.Description) [$ip]")
        }
        $d.ActiveAdapters = $adapList

        # WiFi info via netsh
        $wifi = Read-WiFiInfo
        if ($wifi) {
            $d.WiFi_SSID     = $wifi.SSID
            $d.WiFi_Standard = $wifi.Standard
            $d.WiFi_Signal   = if ($wifi.Signal) { "$($wifi.Signal)%" } else { 'N/A' }
            $d.WiFi_Speed    = if ($wifi.Speed)  { "$($wifi.Speed) Mbps" } else { 'N/A' }
            $wifiStd = if ($wifi.Standard) { " | $($wifi.Standard)" } else { '' }
            $wifiSig = if ($wifi.Signal)   { " | Signal $($wifi.Signal)%" } else { '' }
            Write-Log "WiFi: $($wifi.SSID)$wifiStd$wifiSig"
        }

        # Basic internet check
        $pingOk = $false
        $pingMs = 0
        try {
            $pings  = Test-Connection '8.8.8.8' -Count 3 -ErrorAction Stop
            $pingOk = $true
            $pingMs = [Math]::Round(($pings | Measure-Object ResponseTime -Average).Average, 0)
        } catch {}

        $d.InternetOK = $pingOk
        $d.PingMs     = if ($pingOk) { $pingMs } else { 'Failed' }
        if (-not $pingOk) { $issues.Add('No internet connectivity to 8.8.8.8') }
        else { Write-Log "Internet OK | Ping 8.8.8.8: $($pingMs) ms" }

        # Speedtest CLI
        if ($Script:TOOLS.Speedtest -and $pingOk) {
            Write-Log 'Running Speedtest (~45s)...'
            $st = Invoke-Speedtest
            if ($st) {
                $d.DL_Mbps = $st.DL
                $d.UL_Mbps = $st.UL
                $d.Ping_ms = $st.Ping
                $d.Server  = $st.Server
                Write-Log "Speedtest: DL $($st.DL) Mbps | UL $($st.UL) Mbps | Ping $($st.Ping) ms"
                if ($st.DL -lt $t.Network_MinDlMbps) {
                    $issues.Add("Download $($st.DL) Mbps below threshold $($t.Network_MinDlMbps) Mbps")
                } else {
                    $good.Add("Download $($st.DL) Mbps passes threshold")
                }
            }
        } elseif ($pingOk) {
            $est = Measure-QuickSpeed
            $d.DL_Mbps = "$est (estimated via CDN)"
            Write-Log "Estimated download: $est Mbps"
        }

        # Score
        $score = 70
        if (-not $pingOk) { $score = 20 }
        if ($wifi) {
            if ($wifi.Standard -match '802\.11ax|Wi-Fi 6') {
                $score += $t.Network_WiFi6_Bonus
                $good.Add("WiFi 6 (802.11ax) adapter")
            } elseif ($wifi.Standard -match '802\.11n($|[^a]|\sonly)|^n($|[^a])|Wi-Fi 4') {
                # 802.11n only = Wi-Fi 4 (capped at 150 Mbps single stream / 600 Mbps theoretical max).
                # Modern offices on Wi-Fi 6 networks see real-world bottlenecks here.
                $score -= 10
                $issues.Add("Wi-Fi 4 (802.11n) adapter - outdated wireless, can't fully use modern Wi-Fi networks")
            } elseif ($wifi.Standard -match '802\.11ac|Wi-Fi 5') {
                $good.Add("Wi-Fi 5 (802.11ac) adapter - adequate for most office workloads")
            }
            if ($wifi.Signal -and $wifi.Signal -ge 80)     { $good.Add("Strong signal $($wifi.Signal)%") }
            elseif ($wifi.Signal -and $wifi.Signal -lt 50) { $issues.Add("Weak WiFi signal $($wifi.Signal)%") }
        }
        if ($d.DL_Mbps -match '^\d') {
            $dlVal = [double]($d.DL_Mbps -replace '[^\d\.]')
            if ($dlVal -gt 0) { $score = Clamp100 ($score + [Math]::Min(25, ($dlVal / $t.Network_MinDlMbps) * 20)) }
        }
        if (-not $Script:CAPS.HasWiFi) { $issues.Add('No WiFi adapter detected') }
        if ($Script:CAPS.HasEthernet) {
            if ($Script:CAPS.EthernetActive -or -not $Script:CAPS.IsLaptop) {
                $good.Add('Ethernet adapter present')
            } else {
                $good.Add('Ethernet adapter detected (always disconnected on laptop - may be embedded silicon with no physical port)')
            }
        }

        $score  = Clamp100 $score
        $status = Score-ToStatus $score
        $wifiLabel = if ($wifi -and $wifi.SSID) { "WiFi: $($wifi.SSID)$(if($wifi.Standard){" ($($wifi.Standard))"}) | " }
                     elseif ($Script:CAPS.HasWiFi) { 'WiFi: Connected | ' }
                     else { 'WiFi: N/A | ' }
        $sumStr = "$wifiLabel`DL: $(if($d.DL_Mbps){"$($d.DL_Mbps) Mbps"}else{"N/A"}) | Internet: $pingOk"

        $Script:RESULTS.Network = New-Result -Module Network -Score $score -Status $status `
            -Summary $sumStr -Data $d -Issues $issues -Good $good

        Write-Log "Network Score: $score ($status)" 'OK'
    } catch {
        Write-Log "Network error: $_" 'ERROR'
        $Script:RESULTS.Network = New-Result -Module Network -Score 0 -Status ERROR -Summary "Error: $_"
    }
}

function Read-WiFiInfo {
    try {
        $out = & netsh wlan show interfaces 2>&1
        # Don't exit on $LASTEXITCODE  -  some drivers return non-zero even when connected
        $r = @{}
        foreach ($l in $out) {
            if ($l -match '^\s+SSID\s*:\s+(.+)$')          { $r.SSID     = $Matches[1].Trim() }
            if ($l -match 'Radio type\s*:\s*(.+)$')          { $r.Standard = $Matches[1].Trim() }
            if ($l -match 'Signal\s*:\s*(\d+)%')             { $r.Signal   = [int]$Matches[1] }
            if ($l -match 'Receive rate.*:\s*([\d\.]+)')      { $r.Speed    = [int]$Matches[1] }
            if ($l -match 'Authentication\s*:\s*(.+)$')       { $r.Auth     = $Matches[1].Trim() }
            if ($l -match 'Channel\s*:\s*(\d+)')              { $r.Channel  = [int]$Matches[1] }
        }
        # Fallback: Get-NetConnectionProfile gives SSID even when netsh parsing fails
        if (-not $r.SSID) {
            # Try adapter-specific lookup first (most reliable)
            $wifiAdapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
                Where-Object { $_.InterfaceDescription -match 'Wireless|Wi-Fi|WiFi|802\.11|WLAN' -or $_.PhysicalMediaType -eq 'Native 802.11' }
            foreach ($wAdapter in $wifiAdapters) {
                $prof = Get-NetConnectionProfile -InterfaceAlias $wAdapter.Name -ErrorAction SilentlyContinue
                if ($prof -and $prof.Name -and $prof.Name -ne 'Network') { $r.SSID = $prof.Name; break }
            }
        }
        if (-not $r.SSID) {
            $prof = Get-NetConnectionProfile -ErrorAction SilentlyContinue |
                    Where-Object { $_.IPv4Connectivity -eq 'Internet' -and $_.InterfaceAlias -notmatch 'Tailscale|Loopback|TAP|VPN' } |
                    Select-Object -First 1
            if (-not $prof) {
                $prof = Get-NetConnectionProfile -ErrorAction SilentlyContinue |
                        Where-Object { $_.InterfaceAlias -notmatch 'Tailscale|Loopback|TAP|VPN' } |
                        Select-Object -First 1
            }
            if ($prof -and $prof.Name -ne 'Network') { $r.SSID = $prof.Name }
        }
        if ($r.SSID) { return $r }
        return $null
    } catch { return $null }
}

function Invoke-Speedtest {
    try {
        $exe = $Script:TOOLS.Speedtest
        $raw = & $exe --accept-license --accept-gdpr --format=json 2>&1
        $json = ($raw | Where-Object { $_ -match '^\{' } | Select-Object -Last 1)
        if (-not $json) { return $null }
        $st = $json | ConvertFrom-Json
        return @{
            DL     = [Math]::Round($st.download.bandwidth * 8 / 1000000, 1)
            UL     = [Math]::Round($st.upload.bandwidth   * 8 / 1000000, 1)
            Ping   = [Math]::Round($st.ping.latency, 0)
            Server = "$($st.server.name), $($st.server.location)"
        }
    } catch { Write-Log "Speedtest parse failed: $_" 'WARN'; return $null }
}

function Measure-QuickSpeed {
    try {
        $url  = 'https://speed.cloudflare.com/__downbytes=5242880'
        $sw   = [System.Diagnostics.Stopwatch]::StartNew()
        $wc   = New-Object System.Net.WebClient
        $wc.Headers.Add('User-Agent', 'HardwareEval')
        $data = $wc.DownloadData($url)
        $sw.Stop()
        return [Math]::Round(($data.Length * 8) / ($sw.Elapsed.TotalSeconds * 1MB), 1)
    } catch { return 'Failed' }
}

# Capture 1.5s of mic audio via WinMM and return peak amplitude (0-32767).
# Returns -1 = no capture device, -2 = device open failed, -99 = Add-Type error.
function Measure-MicPeak {
    param([int]$DurationMs = 1500)
    try {
        if (-not ([System.Management.Automation.PSTypeName]'WinMMCapture').Type) {
            Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices; using System.Threading;
public class WinMMCapture {
    [StructLayout(LayoutKind.Sequential)] struct WAVEFORMATEX {
        public ushort wFormatTag, nChannels;
        public uint   nSamplesPerSec, nAvgBytesPerSec;
        public ushort nBlockAlign, wBitsPerSample, cbSize;
    }
    [StructLayout(LayoutKind.Sequential)] struct WAVEHDR {
        public IntPtr lpData;
        public uint   dwBufferLength, dwBytesRecorded;
        public IntPtr dwUser;
        public uint   dwFlags, dwLoops;
        public IntPtr lpNext, reserved;
    }
    [DllImport("winmm.dll")] static extern uint waveInGetNumDevs();
    [DllImport("winmm.dll")] static extern int waveInOpen(out IntPtr h, uint dev, ref WAVEFORMATEX f, IntPtr cb, IntPtr inst, uint fl);
    [DllImport("winmm.dll")] static extern int waveInPrepareHeader(IntPtr h, ref WAVEHDR hdr, uint sz);
    [DllImport("winmm.dll")] static extern int waveInAddBuffer(IntPtr h, ref WAVEHDR hdr, uint sz);
    [DllImport("winmm.dll")] static extern int waveInStart(IntPtr h);
    [DllImport("winmm.dll")] static extern int waveInStop(IntPtr h);
    [DllImport("winmm.dll")] static extern int waveInReset(IntPtr h);
    [DllImport("winmm.dll")] static extern int waveInUnprepareHeader(IntPtr h, ref WAVEHDR hdr, uint sz);
    [DllImport("winmm.dll")] static extern int waveInClose(IntPtr h);
    public static int MeasurePeak(int durationMs) {
        if (waveInGetNumDevs() == 0) return -1;
        var fmt = new WAVEFORMATEX { wFormatTag=1, nChannels=1, nSamplesPerSec=16000,
                                     nAvgBytesPerSec=32000, nBlockAlign=2, wBitsPerSample=16 };
        int bufBytes = (int)(fmt.nAvgBytesPerSec * durationMs / 1000);
        IntPtr hwi;
        if (waveInOpen(out hwi, 0xFFFFFFFF, ref fmt, IntPtr.Zero, IntPtr.Zero, 0) != 0) return -2;
        IntPtr buf = Marshal.AllocHGlobal(bufBytes);
        try {
            var hdr = new WAVEHDR { lpData = buf, dwBufferLength = (uint)bufBytes };
            uint hsz = (uint)Marshal.SizeOf(typeof(WAVEHDR));
            waveInPrepareHeader(hwi, ref hdr, hsz);
            waveInAddBuffer(hwi, ref hdr, hsz);
            waveInStart(hwi);
            Thread.Sleep(durationMs + 100);
            waveInStop(hwi); waveInReset(hwi);
            byte[] b = new byte[bufBytes]; Marshal.Copy(buf, b, 0, bufBytes);
            int peak = 0;
            for (int i = 0; i + 1 < b.Length; i += 2) {
                int s = Math.Abs((short)(b[i] | (b[i+1] << 8)));
                if (s > peak) peak = s;
            }
            waveInUnprepareHeader(hwi, ref hdr, hsz); waveInClose(hwi);
            return peak;
        } finally { Marshal.FreeHGlobal(buf); }
    }
}
'@ -ErrorAction Stop
        }
        return [WinMMCapture]::MeasurePeak($DurationMs)
    } catch { return -99 }
}

# ============================================================
#  MODULE: PORTS & PERIPHERALS
# ============================================================

function Test-Ports {
    Write-Section 'Ports and Peripherals'
    $issues = [System.Collections.Generic.List[string]]::new()
    $good   = [System.Collections.Generic.List[string]]::new()
    $d      = [ordered]@{}

    try {
        # USB controllers
        $usbCtrls = @(Get-CimInstance Win32_USBController -ErrorAction SilentlyContinue)
        $usb3     = @($usbCtrls | Where-Object { $_.Name -match 'USB 3|xHCI' })
        $usb2     = @($usbCtrls | Where-Object { $_.Name -match 'EHCI|USB 2' })
        $d.USB3_Controllers = $usb3.Count
        $d.USB2_Controllers = $usb2.Count
        if ($usb3.Count -ge 2) { $good.Add("USB 3.x present ($($usb3.Count) controller(s))") }
        elseif ($usb3.Count -eq 0) { $issues.Add('No USB 3.x controllers detected') }

        # USB Type-C (detect via PnP device names)
        $usbC = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            $_.FriendlyName -match 'USB.*Type.C|USB-C|UCM|USB Type C' -and $_.Status -eq 'OK'
        })
        $d.USB_TypeC = if ($usbC.Count -gt 0) { 'Present' } else { 'Not detected' }
        if ($usbC.Count -gt 0) { $good.Add('USB Type-C port present') }

        # Thunderbolt
        $tb = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object { $_.FriendlyName -match 'Thunderbolt' })
        $d.Thunderbolt = if ($tb.Count -gt 0) { $tb[0].FriendlyName } else { 'Not detected' }
        if ($tb.Count -gt 0) { $good.Add("Thunderbolt: $($tb[0].FriendlyName)") }

        # Display outputs (HDMI / DisplayPort)
        $dispOut = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            $_.FriendlyName -match 'HDMI|DisplayPort' -and $_.Status -eq 'OK'
        })
        $d.Display_Outputs = if ($dispOut.Count -gt 0) { ($dispOut | ForEach-Object { $_.FriendlyName }) -join ' | ' } else { 'Not directly detected' }

        # SD card reader
        $sdReader = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            $_.FriendlyName -match 'SD|Card Reader|MicroSD|Memory Card' -and $_.Status -eq 'OK'
        })
        $d.SD_CardReader = if ($sdReader.Count -gt 0) { $sdReader[0].FriendlyName } else { 'Not detected' }
        if ($sdReader.Count -gt 0) { $good.Add("SD card reader: $($sdReader[0].FriendlyName)") }

        # Audio
        $audio = @(Get-CimInstance Win32_SoundDevice -ErrorAction SilentlyContinue)
        $d.AudioDevices = ($audio | ForEach-Object { $_.Name }) -join ' | '
        if ($audio.Count -eq 0) { $issues.Add('No audio devices detected') }
        else { $good.Add("$($audio.Count) audio device(s)") }

        # Webcam  -  try to get model name and flag 1080p
        $cams = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            ($_.Class -eq 'Camera' -or $_.FriendlyName -match 'webcam|camera|imaging') -and
            $_.Status -eq 'OK' -and $_.FriendlyName -notmatch 'IR|infrared'
        })
        if ($cams.Count -gt 0) {
            $camName = $cams[0].FriendlyName
            $d.Webcam = $camName
            $camNote = if ($camName -match '1080|FHD') { 'FHD webcam' } elseif ($camName -match '720|HD') { 'HD webcam' } else { 'webcam present' }
            $good.Add("$camNote ($camName)")
        } else {
            $d.Webcam = 'Not detected'
            if ($Script:CAPS.IsLaptop) { $issues.Add('No webcam detected on laptop') }
        }

        # IR camera (Windows Hello facial)
        $irCam = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            ($_.Class -eq 'Camera' -or $_.FriendlyName -match 'camera') -and
            $_.FriendlyName -match 'IR|infrared' -and $_.Status -eq 'OK'
        })
        $d.IR_Camera = if ($irCam.Count -gt 0) { $irCam[0].FriendlyName } else { 'Not detected' }
        if ($irCam.Count -gt 0) { $good.Add("IR camera (Windows Hello facial recognition) present") }

        # Bluetooth radio  -  prefer devices with USB/PCI InstanceId (physical radio, not peripherals)
        $btAllOK = @(Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue | Where-Object { $_.Status -ne 'Error' })
        $btRadio = $btAllOK | Where-Object { $_.InstanceId -match '^USB\\|^PCI\\' } | Select-Object -First 1
        if (-not $btRadio) {
            # Fallback: any named radio/adapter device
            $btRadio = $btAllOK | Where-Object {
                $_.FriendlyName -match 'Intel|Broadcom|Realtek|Qualcomm|Atheros|MediaTek|Bluetooth Radio|Wireless Bluetooth'
            } | Select-Object -First 1
        }
        # Fallback: infer from WiFi adapter chip (e.g., Intel Wireless-AC 8265 → BT 4.2)
        $btVersionFromWifi = ''
        $wifiAdap = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
                    Where-Object { $_.InterfaceDescription -match 'Wireless|Wi-Fi|802\.11' } |
                    Select-Object -First 1
        if ($wifiAdap) {
            $wDesc = $wifiAdap.InterfaceDescription
            $btVersionFromWifi = if ($wDesc -match '(AX200|AX201|AX210|AX211|AX411|BE200|BE202)') { '5.2+ (Intel AX)' }
                                  elseif ($wDesc -match '(9260|9560|22260|AC 9)') { '5.0 (Intel AC 9xxx)' }
                                  elseif ($wDesc -match '(8265|8260|3165|3168|7265|7260)') { '4.2 (Intel AC 8xxx/7xxx)' }
                                  else { '' }
        }
        if ($Script:CAPS.HasBluetooth) {
            $btVer = 'Detected'
            if ($btRadio) {
                $btName = $btRadio.FriendlyName
                $btVer = if ($btName -match '5\.[2-9]') { '5.2+' }
                         elseif ($btName -match '5\.[01]') { '5.0/5.1' }
                         elseif ($btName -match '4\.2') { '4.2' }
                         elseif ($btName -match '4\.[01]') { '4.0/4.1' }
                         else { if ($btVersionFromWifi) { $btVersionFromWifi } else { 'Detected' } }
            } elseif ($btVersionFromWifi) { $btVer = $btVersionFromWifi }
            $d.Bluetooth = "Present ($btVer)"
            $good.Add("Bluetooth $btVer")
        } else {
            $d.Bluetooth = 'Not detected'
            if ($Script:CAPS.IsLaptop) { $issues.Add('No Bluetooth on laptop') }
        }

        # Fingerprint reader
        $fp = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            $_.FriendlyName -match 'Fingerprint|Biometric|WBDI' -and $_.Status -ne 'Error'
        })
        $d.Fingerprint = if ($fp.Count -gt 0) { $fp[0].FriendlyName } else { 'Not detected' }
        if ($fp.Count -gt 0) { $good.Add("Fingerprint reader: $($fp[0].FriendlyName)") }

        # HID devices
        $hid = @(Get-PnpDevice -Class HIDClass -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'OK' })
        $d.HID_Devices = $hid.Count

        # Problem devices  -  only flag actual Error status (Unknown = admin limitation artifact)
        $problemDevs  = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object { $_.Status -in @('Error','Unknown') })
        $errorDevsFull = @($problemDevs | Where-Object { $_.Status -eq 'Error' })
        if ($errorDevsFull.Count -gt 0) {
            $d.Problem_Devices = ($errorDevsFull | ForEach-Object { "$($_.FriendlyName) [$($_.Class)]" }) -join '; '
            $issues.Add("$($errorDevsFull.Count) device(s) in error state - check Device Manager")
        } else {
            $d.Problem_Devices = 'None'
            $good.Add('No device errors detected')
        }
        if (-not $Script:IS_ADMIN) {
            $d.Note = 'Some device statuses unverifiable without admin - run elevated for complete check'
        }

        # Score  -  only deduct for confirmed Error devices (not Unknown); $errorDevsFull already computed above
        $score = 70
        if ($usb3.Count -ge 2)         { $score += 10 }
        elseif ($usb3.Count -eq 0)     { $score -= 15 }
        if ($tb.Count -gt 0)           { $score += 10 }
        if ($usbC.Count -gt 0)         { $score += 5  }
        if ($cams.Count -gt 0)         { $score += 5  }
        if ($Script:CAPS.HasBluetooth) { $score += 5  }
        if ($sdReader.Count -gt 0)     { $score += 5  }
        $score -= [Math]::Min(20, $errorDevsFull.Count * 5)

        $score  = Clamp100 $score
        $status = Score-ToStatus $score
        $sumStr = "USB3:$($usb3.Count) TypeC:$($usbC.Count -gt 0) TB:$($tb.Count -gt 0) BT:$($Script:CAPS.HasBluetooth) Cam:$($cams.Count -gt 0) FP:$($fp.Count -gt 0)"

        $Script:RESULTS.Ports = New-Result -Module Ports -Score $score -Status $status `
            -Summary $sumStr -Data $d -Issues $issues -Good $good

        Write-Log "Ports Score: $score ($status)" 'OK'
    } catch {
        Write-Log "Ports error: $_" 'ERROR'
        $Script:RESULTS.Ports = New-Result -Module Ports -Score 0 -Status ERROR -Summary "Error: $_"
    }
}

# ============================================================
#  MODULE: AUDIO / WEBCAM
# ============================================================

function Test-Audio {
    Write-Section 'Audio & Webcam Assessment'
    $issues = [System.Collections.Generic.List[string]]::new()
    $good   = [System.Collections.Generic.List[string]]::new()
    $d      = [ordered]@{}

    try {
        # ── Audio controllers ──────────────────────────────────────────────
        $soundDevs = @(Get-CimInstance Win32_SoundDevice -ErrorAction SilentlyContinue)
        $d.AudioControllers = if ($soundDevs.Count -gt 0) { ($soundDevs | ForEach-Object { $_.Name.Trim() }) -join '; ' } else { 'None detected' }
        Write-Log "Audio controllers: $($d.AudioControllers)"
        if ($soundDevs.Count -eq 0) { $issues.Add('No audio controller detected - no sound output or input possible') }
        else { $good.Add("$($soundDevs.Count) audio controller(s) present") }

        # ── Windows Audio service ──────────────────────────────────────────
        $audioSvc = Get-Service -Name 'AudioSrv' -ErrorAction SilentlyContinue
        $d.AudioService = if ($audioSvc) { $audioSvc.Status.ToString() } else { 'Not found' }
        if ($audioSvc -and $audioSvc.Status -ne 'Running') { $issues.Add("Windows Audio service is $($audioSvc.Status) - audio will not work") }
        elseif ($audioSvc) { $good.Add('Windows Audio service is running') }

        # ── Audio endpoints: speakers + microphones ────────────────────────
        $endpoints = @(Get-PnpDevice -Class AudioEndpoint -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'OK' })
        $speakers  = @($endpoints | Where-Object { $_.FriendlyName -notmatch 'Microphone|Mic\b' })
        $mics      = @($endpoints | Where-Object { $_.FriendlyName -match  'Microphone|Mic\b' })

        $d.Playback_Devices  = if ($speakers.Count -gt 0) { ($speakers | Select-Object -First 4 | ForEach-Object { $_.FriendlyName }) -join '; ' } else { 'None detected' }
        $d.Recording_Devices = if ($mics.Count -gt 0)     { ($mics     | Select-Object -First 4 | ForEach-Object { $_.FriendlyName }) -join '; ' } else { 'None detected' }
        Write-Log "Playback : $($d.Playback_Devices)"
        Write-Log "Recording: $($d.Recording_Devices)"

        if ($speakers.Count -gt 0) { $good.Add("$($speakers.Count) playback device(s): $($d.Playback_Devices)") }
        else { $issues.Add('No playback device detected (speakers / headphones)') }
        if ($mics.Count -gt 0) { $good.Add("$($mics.Count) microphone(s): $($d.Recording_Devices)") }
        else { $issues.Add('No microphone detected - Teams/Zoom calls will have no audio input') }

        # ── Speaker functional test: play tone through audio engine ────────
        # Tests the full rendering pipeline (driver -> mixer -> endpoint).
        # Cannot distinguish physical speaker vs headphone without a human listener.
        $d.Speaker_Test = 'Not tested (no playback device)'
        if ($speakers.Count -gt 0 -and $audioSvc -and $audioSvc.Status -eq 'Running') {
            try {
                [System.Media.SystemSounds]::Beep.Play()
                Start-Sleep -Milliseconds 600
                [System.Media.SystemSounds]::Beep.Play()
                $d.Speaker_Test = 'Test tones sent to audio engine - confirm audible output'
                $good.Add('Speaker: test tones played through audio engine (confirm you heard them)')
            } catch {
                $d.Speaker_Test = "Playback call failed: $_"
                $issues.Add('Speaker: audio engine returned error during test tone playback')
            }
        }

        # ── Microphone functional test: capture 1.5s, measure peak level ──
        # Peak 0-32767 (16-bit PCM). A working mic in any normal environment
        # will show ambient noise of at least 100-500. Below 50 = likely muted.
        $micPrivKey = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone'
        $micPrivacy = (Get-ItemProperty $micPrivKey -Name 'Value' -ErrorAction SilentlyContinue).Value
        $d.Mic_Privacy = if ($micPrivacy) { $micPrivacy } else { 'Unknown' }
        if ($micPrivacy -eq 'Deny') { $issues.Add('Microphone blocked by Windows privacy settings - Settings > Privacy > Microphone') }
        elseif ($micPrivacy -eq 'Allow') { $good.Add('Microphone access allowed in privacy settings') }

        $d.Mic_LevelTest = 'Not tested'
        $micTooQuiet = $false
        if ($mics.Count -gt 0 -and $micPrivacy -ne 'Deny') {
            Write-Log 'Measuring microphone level (1.5s capture)...'
            $peak = Measure-MicPeak -DurationMs 1500
            if ($peak -ge 0) {
                $pct = [Math]::Round($peak / 327.67, 1)
                $d.Mic_LevelTest = "Peak $peak / 32767 ($pct% of full scale, 1.5s sample)"
                Write-Log "Mic peak: $peak ($pct%)"
                if ($peak -lt 50) {
                    $issues.Add("Microphone peak $pct% - hardware may be muted or faulty (test in quiet room; speak during test for accurate reading)")
                    $micTooQuiet = $true
                } elseif ($peak -lt 300) {
                    $good.Add("Microphone responsive - ambient level $pct% (speak during test for full reading)")
                } else {
                    $good.Add("Microphone functional - good signal level $pct%")
                }
            } elseif ($peak -eq -1) { $d.Mic_LevelTest = 'No WinMM capture device found'
            } elseif ($peak -eq -2) { $d.Mic_LevelTest = 'Could not open capture device'
            } else                  { $d.Mic_LevelTest = "Test error (code $peak)" }
        }

        # ── Audio driver errors ────────────────────────────────────────────
        $audioErrDevs = @(Get-PnpDevice -Class MEDIA -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Error' })
        if ($audioErrDevs.Count -gt 0) {
            $d.Driver_Errors = ($audioErrDevs | ForEach-Object { $_.FriendlyName }) -join '; '
            $issues.Add("$($audioErrDevs.Count) audio device(s) with driver errors: $($d.Driver_Errors)")
        } else { $good.Add('No audio driver errors') }

        # ── Webcam / Camera ────────────────────────────────────────────────
        $cameras = @(Get-PnpDevice -Class Camera -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'OK' })
        if ($cameras.Count -eq 0) {
            $cameras = @(Get-PnpDevice -Class Image -ErrorAction SilentlyContinue |
                         Where-Object { $_.Status -eq 'OK' -and $_.FriendlyName -match 'Camera|Webcam|Cam\b|HD\s+\d+p' })
        }
        $rgbCameras = @($cameras | Where-Object { $_.FriendlyName -notmatch '\bIR\b|Hello|Face Auth|Windows Hello' })
        $irCameras  = @($cameras | Where-Object { $_.FriendlyName -match  '\bIR\b|Hello|Face Auth|Windows Hello' })

        $d.Webcam = if ($rgbCameras.Count -gt 0) { ($rgbCameras | ForEach-Object { $_.FriendlyName }) -join '; ' } else { 'Not detected' }
        Write-Log "Webcam: $($d.Webcam)"

        if ($rgbCameras.Count -gt 0) { $good.Add("Webcam driver OK: $($rgbCameras[0].FriendlyName)") }
        elseif ($Script:CAPS.IsLaptop) { $issues.Add('No webcam detected - video calls will have no camera feed') }

        if ($irCameras.Count -gt 0) {
            $d.IR_Camera = ($irCameras | ForEach-Object { $_.FriendlyName }) -join '; '
            $good.Add("Windows Hello IR camera present: $($d.IR_Camera)")
        }

        # Camera WinRT init test - verifies hardware + driver stack beyond PnP presence
        $d.Camera_Test = 'Not tested'
        if ($rgbCameras.Count -gt 0) {
            try {
                $null = [Windows.Media.Devices.MediaDevice,Windows.Media.Devices,ContentType=WindowsRuntime]
                $sel  = [Windows.Media.Devices.MediaDevice]::GetVideoCaptureSelector()
                Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction SilentlyContinue
                $task = [System.WindowsRuntimeSystemExtensions]::AsTask(
                    [Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($sel))
                if ($task.Wait(4000) -and $task.Result.Count -gt 0) {
                    $d.Camera_Test = "WinRT init OK - $($task.Result.Count) device(s) accessible"
                    $good.Add("Camera WinRT stack initialised successfully ($($task.Result.Count) device(s))")
                } else {
                    $d.Camera_Test = 'WinRT enumeration returned no devices'
                }
            } catch { $d.Camera_Test = 'WinRT check skipped (PS 5.1 limitation)' }
        }

        # Camera privacy
        $camPrivKey = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam'
        $camPrivacy = (Get-ItemProperty $camPrivKey -Name 'Value' -ErrorAction SilentlyContinue).Value
        $d.Camera_Privacy = if ($camPrivacy) { $camPrivacy } else { 'Unknown' }
        if ($camPrivacy -eq 'Deny' -and $cameras.Count -gt 0) {
            $issues.Add('Camera blocked by Windows privacy settings - Settings > Privacy > Camera')
        }

        # ── Score ──────────────────────────────────────────────────────────
        $score = 55
        if ($soundDevs.Count -eq 0) { $score = 10 } else {
            if ($speakers.Count -gt 0)                          { $score += 15 }
            if ($mics.Count -gt 0)                              { $score += 15 }
            if ($rgbCameras.Count -gt 0)                        { $score += 10 }
            if ($irCameras.Count -gt 0)                         { $score += 5  }
            if ($audioErrDevs.Count -gt 0)                      { $score -= 20 }
            if ($audioSvc -and $audioSvc.Status -ne 'Running')  { $score  = [Math]::Min($score, 20) }
            if ($micPrivacy -eq 'Deny')                         { $score -= 10 }
            if ($camPrivacy -eq 'Deny')                         { $score -= 5  }
            if ($micTooQuiet)                                   { $score -= 10 }
        }

        $score  = Clamp100 $score
        $status = Score-ToStatus $score

        $sumParts = [System.Collections.Generic.List[string]]::new()
        $sumParts.Add($(if ($speakers.Count -gt 0) { "Speakers: OK" } else { "Speakers: None" }))
        $sumParts.Add($(if ($mics.Count -gt 0)     { "Mic: OK"      } else { "Mic: None" }))
        if ($rgbCameras.Count -gt 0) { $sumParts.Add("Webcam: OK") }
        if ($irCameras.Count -gt 0)  { $sumParts.Add("Hello IR: OK") }

        $Script:RESULTS.Audio = New-Result -Module Audio -Score $score -Status $status `
            -Summary ($sumParts -join ' | ') -Data $d -Issues $issues -Good $good

        Write-Log "Audio Score: $score ($status)" 'OK'
    } catch {
        Write-Log "Audio error: $_" 'ERROR'
        $Script:RESULTS.Audio = New-Result -Module Audio -Score 0 -Status ERROR -Summary "Error: $_"
    }
}

# ============================================================
#  MODULE: INPUT DEVICES
# ============================================================

#  Keyboard test using a WinForms KeyDown capture window (STA runspace).
#  WinForms catches more keys than [Console]::ReadKey: modifiers as standalone keys,
#  media keys (the codes Lenovo Fn-locked F-keys actually emit), arrows, nav cluster.
#  No auto-timeout - only finishes on all-pressed OR double-Escape.
#  Returns @{ Pressed=@(...); NotPressed=@(...); Total=N; Detected=N }
function Invoke-KeyboardTest {
    # keyDef: ConsoleKey-style identifier -> @(displayName, [accepted KeyCode strings...])
    # Multiple accepted KeyCodes lets us treat Lenovo Fn-locked F-keys (which emit
    # VolumeMute / VolumeDown / VolumeUp etc.) as F1/F2/F3 hits.
    $keyDef = [ordered]@{}
    foreach ($c in [char[]]'QWERTYUIOPASDFGHJKLZXCVBNM') { $keyDef["$c"] = @("$c", "$c") }
    for ($i = 0; $i -le 9; $i++)                          { $keyDef["D$i"] = @("$i", "D$i") }
    $keyDef['F1']  = @('F1',  'F1', 'VolumeMute')
    $keyDef['F2']  = @('F2',  'F2', 'VolumeDown')
    $keyDef['F3']  = @('F3',  'F3', 'VolumeUp')
    $keyDef['F4']  = @('F4',  'F4')
    $keyDef['F5']  = @('F5',  'F5')
    $keyDef['F6']  = @('F6',  'F6')
    $keyDef['F7']  = @('F7',  'F7')
    $keyDef['F8']  = @('F8',  'F8')
    $keyDef['F9']  = @('F9',  'F9')
    $keyDef['F10'] = @('F10', 'F10')
    $keyDef['F11'] = @('F11', 'F11')
    $keyDef['F12'] = @('F12', 'F12', 'LaunchApplication2')
    $keyDef['LShiftKey']   = @('L-Shift', 'LShiftKey', 'ShiftKey')
    $keyDef['RShiftKey']   = @('R-Shift', 'RShiftKey')
    $keyDef['LControlKey'] = @('L-Ctrl',  'LControlKey', 'ControlKey')
    $keyDef['RControlKey'] = @('R-Ctrl',  'RControlKey')
    $keyDef['LMenu']       = @('L-Alt',   'LMenu', 'Menu')
    $keyDef['RMenu']       = @('R-Alt',   'RMenu')
    $keyDef['LWin']        = @('Win',     'LWin', 'RWin')
    $keyDef['Apps']        = @('Menu',    'Apps')
    $keyDef['Oem3']        = @('` ~',     'Oem3', 'Oemtilde')
    $keyDef['OemMinus']    = @('-',       'OemMinus')
    $keyDef['Oemplus']     = @('=',       'Oemplus', 'OemPlus')
    $keyDef['Oem4']        = @('[',       'Oem4', 'OemOpenBrackets')
    $keyDef['Oem6']        = @(']',       'Oem6', 'OemCloseBrackets')
    $keyDef['Oem5']        = @('\',       'Oem5', 'OemPipe', 'OemBackslash')
    $keyDef['Oem1']        = @(';',       'Oem1', 'OemSemicolon')
    $keyDef['Oem7']        = @("'",       'Oem7', 'OemQuotes')
    $keyDef['Oemcomma']    = @(',',       'Oemcomma', 'OemComma')
    $keyDef['OemPeriod']   = @('.',       'OemPeriod')
    $keyDef['OemQuestion'] = @('/',       'OemQuestion', 'OemSlash')
    $keyDef['Space']       = @('Space',   'Space')
    $keyDef['Enter']       = @('Enter',   'Enter', 'Return')
    $keyDef['Back']        = @('BkSp',    'Back', 'Backspace')
    $keyDef['Tab']         = @('Tab',     'Tab')
    $keyDef['CapsLock']    = @('CapsLk',  'CapsLock', 'Capital')
    $keyDef['Escape']      = @('Esc',     'Escape')
    $keyDef['Insert']      = @('Ins',     'Insert')
    $keyDef['Delete']      = @('Del',     'Delete')
    $keyDef['Home']        = @('Home',    'Home')
    $keyDef['End']         = @('End',     'End')
    $keyDef['PageUp']      = @('PgUp',    'PageUp', 'Prior')
    $keyDef['PageDown']    = @('PgDn',    'PageDown', 'Next')
    $keyDef['Up']          = @('Up',      'Up')
    $keyDef['Down']        = @('Down',    'Down')
    $keyDef['Left']        = @('Left',    'Left')
    $keyDef['Right']       = @('Right',   'Right')

    # Synchronized state shared with the runspace's KeyDown handler
    $state = [hashtable]::Synchronized(@{
        Pressed  = @{}
        EscCount = 0
        KeyDef   = $keyDef
    })
    foreach ($k in $keyDef.Keys) { $state.Pressed[$k] = $false }

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('state', $state)

    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $form                 = New-Object System.Windows.Forms.Form
        $form.Text            = 'Keyboard Test'
        $form.Size            = New-Object System.Drawing.Size(960, 620)
        $form.StartPosition   = 'CenterScreen'
        $form.KeyPreview      = $true
        $form.TopMost         = $true
        $form.MaximizeBox     = $false
        $form.FormBorderStyle = 'FixedSingle'
        $form.BackColor       = [System.Drawing.Color]::FromArgb(20,20,20)
        $form.ForeColor       = [System.Drawing.Color]::WhiteSmoke

        $lblTop               = New-Object System.Windows.Forms.Label
        $lblTop.Dock          = 'Top'
        $lblTop.Height        = 56
        $lblTop.Font          = New-Object System.Drawing.Font('Consolas', 13, [System.Drawing.FontStyle]::Bold)
        $lblTop.ForeColor     = [System.Drawing.Color]::Lime
        $lblTop.TextAlign     = 'MiddleCenter'
        $form.Controls.Add($lblTop)

        $rtb                  = New-Object System.Windows.Forms.RichTextBox
        $rtb.Dock             = 'Fill'
        $rtb.Font             = New-Object System.Drawing.Font('Consolas', 11)
        $rtb.BackColor        = [System.Drawing.Color]::FromArgb(20,20,20)
        $rtb.ForeColor        = [System.Drawing.Color]::Gainsboro
        $rtb.ReadOnly         = $true
        $rtb.BorderStyle      = 'None'
        $form.Controls.Add($rtb)

        $lblBottom            = New-Object System.Windows.Forms.Label
        $lblBottom.Dock       = 'Bottom'
        $lblBottom.Height     = 36
        $lblBottom.Font       = New-Object System.Drawing.Font('Consolas', 10)
        $lblBottom.ForeColor  = [System.Drawing.Color]::Gray
        $lblBottom.TextAlign  = 'MiddleCenter'
        $lblBottom.Text       = 'Press every key. Closes when all keys pressed OR Escape pressed twice. Lenovo F-keys: try Fn+F-key or use the locked function (mute/vol).'
        $form.Controls.Add($lblBottom)

        $script:render = {
            $sb = New-Object System.Text.StringBuilder
            $cnt = 0; $i = 0
            foreach ($k in $state.KeyDef.Keys) {
                $disp = $state.KeyDef[$k][0]
                if ($state.Pressed[$k]) { [void]$sb.Append("[+] $($disp.PadRight(8))"); $cnt++ }
                else                    { [void]$sb.Append("[ ] $($disp.PadRight(8))") }
                $i++
                if ($i % 8 -eq 0) { [void]$sb.AppendLine() }
            }
            $rtb.Text = $sb.ToString()
            $lblTop.Text = "$cnt / $($state.KeyDef.Count) keys pressed"
        }
        & $script:render

        $form.Add_KeyDown({
            param($s, $e)
            $kc = $e.KeyCode.ToString()
            if ($kc -eq 'Escape') { $state.EscCount++ } else { $state.EscCount = 0 }
            foreach ($k in $state.KeyDef.Keys) {
                $accepted = $state.KeyDef[$k]
                for ($j = 1; $j -lt $accepted.Count; $j++) {
                    if ($kc -eq $accepted[$j]) { $state.Pressed[$k] = $true; break }
                }
            }
            & $script:render
            $remaining = ($state.Pressed.Values | Where-Object { -not $_ } | Measure-Object).Count
            if ($remaining -eq 0 -or $state.EscCount -ge 2) { $s.Close() }
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        })

        $form.Add_Shown({ $form.Activate() })
        [void]$form.ShowDialog()
        $form.Dispose()
    })
    try { [void]$ps.Invoke() } catch { Write-Log "Keyboard test UI error: $_" 'WARN' }
    $ps.Dispose()
    $rs.Close()
    $rs.Dispose()

    $pressedNames    = @($state.Pressed.GetEnumerator() | Where-Object { $_.Value }      | ForEach-Object { $keyDef[$_.Key][0] })
    $notPressedNames = @($state.Pressed.GetEnumerator() | Where-Object { -not $_.Value } | ForEach-Object { $keyDef[$_.Key][0] })
    return @{
        Pressed    = $pressedNames
        NotPressed = $notPressedNames
        Total      = $keyDef.Count
        Detected   = $pressedNames.Count
    }
}

function Test-Input {
    Write-Section 'Input Devices Assessment'
    $issues = [System.Collections.Generic.List[string]]::new()
    $good   = [System.Collections.Generic.List[string]]::new()
    $d      = [ordered]@{}

    try {
        # ── Keyboard ──────────────────────────────────────────────────────
        $kbDevs = @(Get-CimInstance Win32_Keyboard -ErrorAction SilentlyContinue)
        $kbPnP  = @(Get-PnpDevice -Class Keyboard -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'OK' })
        $kbErr  = @(Get-PnpDevice -Class Keyboard -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Error' })

        $d.Keyboard = if ($kbDevs.Count -gt 0) { ($kbDevs | ForEach-Object { $_.Description }) -join '; ' } else { 'Not detected' }
        $d.Keyboard_Status = if ($kbErr.Count -gt 0) { "Driver error on $($kbErr.Count) device(s)" } else { 'OK' }
        Write-Log "Keyboard: $($d.Keyboard)"

        if ($kbPnP.Count -gt 0 -and $kbErr.Count -eq 0) { $good.Add("Keyboard driver OK: $($kbPnP.Count) device(s) detected") }
        elseif ($kbErr.Count -gt 0) { $issues.Add("Keyboard driver error - $($kbErr.Count) device(s) showing fault") }
        else { $issues.Add('No keyboard detected') }

        # Keyboard test: open keyboardchecker.com in default browser. Web-based testers
        # catch ALL keys reliably (browser bypasses console limitations and even handles
        # most special-function keys). User tests there, then reports back here.
        if (-not $Script:FullAuto -and $kbPnP.Count -gt 0) {
            Write-Host ''
            Write-Host '  === KEYBOARD TEST ===' -ForegroundColor Yellow
            Write-Host '  Opening keyboardchecker.com in your browser.' -ForegroundColor Cyan
            Write-Host '  1) Click the page once to give it focus' -ForegroundColor Cyan
            Write-Host '  2) Press every key on your keyboard - including F1-F12' -ForegroundColor Cyan
            Write-Host '     (try Fn+F-key on Lenovo if F-keys are locked to brightness/volume)' -ForegroundColor DarkGray
            Write-Host '  3) Note any keys that DID NOT light up on the on-screen keyboard' -ForegroundColor Cyan
            Write-Host '  4) Come back here when done' -ForegroundColor Cyan
            Write-Host ''
            try { Start-Process 'https://keyboardchecker.com/' -ErrorAction Stop } catch {
                Write-Host '  (Could not auto-open browser. Manually open: https://keyboardchecker.com/)' -ForegroundColor Yellow
            }
            Write-Host '  Press Enter when you have finished testing the keyboard...' -ForegroundColor DarkGray -NoNewline
            [void](Read-Host)

            if (Read-YesNo '  Did EVERY key on your keyboard work?') {
                $d.Keyboard_Check = 'PASS - user verified all keys via keyboardchecker.com'
                $d.Keyboard_Test_Method = 'keyboardchecker.com (browser)'
                $good.Add('Keyboard: all keys verified by user')
            } else {
                Write-Host '  List the broken keys (comma-separated, e.g. "F4, Caps, Right Shift"): ' -ForegroundColor Yellow -NoNewline
                $brokenKeys = (Read-Host).Trim()
                if (-not $brokenKeys) { $brokenKeys = '(unspecified)' }
                $d.Keyboard_Check       = "FAIL - broken keys: $brokenKeys"
                $d.Keyboard_Broken_Keys = $brokenKeys
                $d.Keyboard_Test_Method = 'keyboardchecker.com (browser)'
                $issues.Add("Keyboard: broken keys reported - $brokenKeys")
            }
            Write-Log "Keyboard: $($d.Keyboard_Check)" 'OK'
        } else {
            $d.Keyboard_Note = if ($Script:FullAuto) { 'Key test skipped (FullAuto mode)' } else { 'No keyboard - skipped' }
        }

        # ── Touchpad / Pointing device ─────────────────────────────────────
        $allPointing = @(Get-CimInstance Win32_PointingDevice -ErrorAction SilentlyContinue)
        $touchpads   = @($allPointing | Where-Object {
            $_.Name -match 'TouchPad|Synaptics|ELAN|Alps|Precision|Cirque|HID.*(Touch|Touchpad)' -or
            $_.DeviceID -match 'ACPI.*SYN|ACPI.*ELAN|ACPI.*ALPS'
        })
        $d.Pointing_Devices = if ($allPointing.Count -gt 0) { ($allPointing | ForEach-Object { $_.Name }) -join '; ' } else { 'None' }

        # Precision Touchpad: kernel-level multi-gesture support
        $isPrecision = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PrecisionTouchPad'
        $d.Precision_Touchpad = if ($isPrecision) { 'Yes (Windows Precision Touchpad)' } else { 'No (vendor driver)' }

        Write-Log "Pointing: $($d.Pointing_Devices) | Precision: $isPrecision"

        if ($touchpads.Count -gt 0 -or $isPrecision) {
            $precStr = if ($isPrecision) { ' (Precision Touchpad - multi-gesture supported)' } else { '' }
            $good.Add("Touchpad detected$precStr")
        } elseif ($Script:CAPS.IsLaptop -and $allPointing.Count -eq 0) {
            $issues.Add('No pointing device detected - touchpad may be disabled or driver missing')
        }
        if ($allPointing.Count -gt 0) { $d.Touchpad_Note = 'Gesture and multi-touch testing requires manual verification' }

        # ── Touch screen ──────────────────────────────────────────────────
        $touchScreen = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            $_.Status -eq 'OK' -and $_.FriendlyName -match 'HID-Compliant Touch Screen|Touch Digitizer|Wacom|Pen Digitizer'
        })
        $d.Touch_Screen = if ($touchScreen.Count -gt 0) { ($touchScreen | ForEach-Object { $_.FriendlyName }) -join '; ' } else { 'Not detected' }
        Write-Log "Touch screen: $($d.Touch_Screen)"
        if ($touchScreen.Count -gt 0) { $good.Add("Touch screen driver OK: $($touchScreen[0].FriendlyName)") }

        # Interactive touchscreen verification
        if ($touchScreen.Count -gt 0 -and -not $Script:FullAuto) {
            Write-Host ''
            Write-Host '  === TOUCH SCREEN CHECK ===' -ForegroundColor Yellow
            $tabcal = Join-Path $env:SystemRoot 'System32\tabcal.exe'
            if (Test-Path $tabcal) {
                Write-Host '  Launching Windows Tablet Calibration (tabcal.exe).' -ForegroundColor Cyan
                Write-Host '  A full white screen will open with a crosshair.' -ForegroundColor Cyan
                Write-Host '    1) Tap each crosshair as it appears at different screen positions' -ForegroundColor Gray
                Write-Host '    2) Right-click to return to the previous point if needed' -ForegroundColor Gray
                Write-Host '    3) Press ESC any time to close the tool' -ForegroundColor Gray
                Write-Host '    4) DO NOT rotate the screen during calibration' -ForegroundColor Gray
                Write-Host '  Calibration data is saved to the registry on completion.' -ForegroundColor DarkGray
                Write-Host '  Press Enter to launch tabcal, or Ctrl+C to skip...' -ForegroundColor DarkGray -NoNewline
                [void](Read-Host)
                try {
                    $proc = Start-Process -FilePath $tabcal -PassThru -Wait -ErrorAction Stop
                    Write-Host "  tabcal.exe exited (code $($proc.ExitCode))." -ForegroundColor Gray
                } catch {
                    Write-Log "Could not launch tabcal: $_" 'WARN'
                }
                if (Read-YesNo '  Did the calibration complete and touch respond accurately at every crosshair?') {
                    $d.Touch_Screen_Check = 'PASS (verified via tabcal.exe calibration)'
                    $d.Touch_Calibration  = 'Completed successfully'
                    $good.Add('Touch screen: calibration completed and verified')
                } else {
                    $d.Touch_Screen_Check = 'FAIL (calibration could not complete or touch was inaccurate)'
                    $d.Touch_Calibration  = 'Failed or aborted'
                    $issues.Add('Touch screen: calibration via tabcal.exe failed or touch was inaccurate')
                }
            } else {
                # Fallback when tabcal.exe is not installed (some Server SKUs / Tablet PC feature stripped)
                Write-Host '  tabcal.exe not available - falling back to manual check.' -ForegroundColor DarkGray
                Write-Host '  Please tap the screen a few times and try a swipe gesture.' -ForegroundColor Cyan
                if (Read-YesNo '  Did touch input respond correctly?') {
                    $d.Touch_Screen_Check = 'PASS (user verified)'
                    $good.Add('Touch screen: confirmed working by user')
                } else {
                    $d.Touch_Screen_Check = 'FAIL (user reported non-functional)'
                    $issues.Add('Touch screen: user reported touch input not working')
                }
            }
        }

        # ── Fingerprint reader ─────────────────────────────────────────────
        # Biometric class includes both fingerprint sensors AND facial-recognition IR cameras.
        # Exclude facial-recognition devices so we only report actual fingerprint hardware.
        $allBio  = @(Get-PnpDevice -Class Biometric -ErrorAction SilentlyContinue |
                     Where-Object { $_.FriendlyName -notmatch 'Facial Recognition|Face Recognition|IR Camera|Face Auth|Camera' })
        $fpsOK   = @($allBio | Where-Object { $_.Status -eq 'OK'    })
        $fpsErr  = @($allBio | Where-Object { $_.Status -eq 'Error' })

        if ($fpsErr.Count -gt 0) {
            $d.Fingerprint_Reader = "DRIVER ERROR: $($fpsErr[0].FriendlyName)"
            $issues.Add("Fingerprint sensor driver error - sensor may be physically damaged or disconnected")
        } elseif ($fpsOK.Count -gt 0) {
            $d.Fingerprint_Reader = ($fpsOK | ForEach-Object { $_.FriendlyName }) -join '; '
            $d.Fingerprint_Note   = 'Driver present. Functional verification requires a Windows Hello enrollment attempt.'
            $good.Add("Fingerprint reader driver OK: $($fpsOK[0].FriendlyName)")
        } else {
            $d.Fingerprint_Reader = 'Not detected'
        }
        Write-Log "Fingerprint: $($d.Fingerprint_Reader)"

        # ── NFC ───────────────────────────────────────────────────────────
        $nfc = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            $_.Status -eq 'OK' -and $_.FriendlyName -match 'NFC|Near Field Communication|Contactless'
        })
        $d.NFC = if ($nfc.Count -gt 0) { ($nfc | ForEach-Object { $_.FriendlyName }) -join '; ' } else { 'Not detected' }
        if ($nfc.Count -gt 0) { $good.Add("NFC reader present: $($nfc[0].FriendlyName)") }

        # ── SD / Memory card reader ────────────────────────────────────────
        $sdReader = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            $_.Status -eq 'OK' -and $_.FriendlyName -match 'SD|Memory Card|Card Reader|MMC|SDXC|SDHC'
        })
        $d.SD_Card_Reader = if ($sdReader.Count -gt 0) { ($sdReader | Select-Object -First 2 | ForEach-Object { $_.FriendlyName }) -join '; ' } else { 'Not detected' }
        if ($sdReader.Count -gt 0) { $good.Add("SD/Memory card reader: $($sdReader[0].FriendlyName)") }

        # ── Keyboard backlight ─────────────────────────────────────────────
        $kbLight = Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            $_.Status -eq 'OK' -and $_.FriendlyName -match 'Keyboard.*Light|Backlight.*Keyboard|LED.*Keyboard'
        } | Select-Object -First 1
        # Also try ACPI method
        $kbLightAcpi = try {
            $acpi = Get-CimInstance -Namespace root\wmi -ClassName LENOVO_LIGHTING_METHOD -ErrorAction Stop
            if ($acpi) { 'Detected (Lenovo ACPI)' } else { $null }
        } catch { $null }
        $d.Keyboard_Backlight = if ($kbLight) { $kbLight.FriendlyName } elseif ($kbLightAcpi) { $kbLightAcpi } else { 'Not detected / vendor-specific' }

        # ── Score ──────────────────────────────────────────────────────────
        $score = 60
        if ($kbPnP.Count -gt 0 -and $kbErr.Count -eq 0) { $score += 20 }
        elseif ($kbErr.Count -gt 0)                      { $score -= 20 }
        if ($Script:CAPS.IsLaptop) {
            if ($touchpads.Count -gt 0 -or $isPrecision) { $score += 10 }
            else                                          { $score -= 10 }
            if ($touchScreen.Count -gt 0) { $score += 5 }
        }
        if ($fpsOK.Count -gt 0) { $score += 5 }
        elseif ($fpsErr.Count -gt 0) { $score -= 5 }
        if ($d.Keyboard_Check    -like 'FAIL*') { $score -= 15 }
        if ($d.Touch_Screen_Check -like 'FAIL*') { $score -= 10 }

        $score  = Clamp100 $score
        $status = Score-ToStatus $score

        $sumParts = [System.Collections.Generic.List[string]]::new()
        $kbCheckStr = if ($d.Keyboard_Check -like 'PASS*') { ' [Verified]' } elseif ($d.Keyboard_Check -like 'FAIL*') { ' [FAIL]' } else { '' }
        $sumParts.Add($(if ($kbPnP.Count -gt 0) { "Keyboard: OK$kbCheckStr" } else { "Keyboard: Issue" }))
        if ($Script:CAPS.IsLaptop) {
            $sumParts.Add($(if ($touchpads.Count -gt 0 -or $isPrecision) { "Touchpad: OK$(if($isPrecision){' (Precision)'})" } else { "Touchpad: None" }))
        }
        if ($touchScreen.Count -gt 0) {
            $tsStr = if ($d.Touch_Screen_Check -like 'PASS*') { ' [Verified]' } elseif ($d.Touch_Screen_Check -like 'FAIL*') { ' [FAIL]' } else { '' }
            $sumParts.Add("TouchScreen: Yes$tsStr")
        }
        if ($fpsOK.Count -gt 0)          { $sumParts.Add("Fingerprint: Yes") }
        elseif ($fpsErr.Count -gt 0)    { $sumParts.Add("Fingerprint: ERROR") }

        $Script:RESULTS.Input = New-Result -Module Input -Score $score -Status $status `
            -Summary ($sumParts -join ' | ') -Data $d -Issues $issues -Good $good

        Write-Log "Input Score: $score ($status)" 'OK'
    } catch {
        Write-Log "Input error: $_" 'ERROR'
        $Script:RESULTS.Input = New-Result -Module Input -Score 0 -Status ERROR -Summary "Error: $_"
    }
}

# ============================================================
#  MODULE: SOFTWARE HEALTH
# ============================================================

# Map of common Windows bugcheck codes -> friendly name + likely cause.
# Source: Microsoft docs (Bug Check Code Reference). Origin is rough - many bugchecks
# can be either, but flagging "Hardware" ones (especially WHEA / MCE) points the
# technician at a real component failure rather than a driver/software fix.
$Script:BugcheckMap = @{
    '0x0000001E' = @{ Name='KMODE_EXCEPTION_NOT_HANDLED'; Origin='Driver/Software' }
    '0x00000019' = @{ Name='BAD_POOL_HEADER';             Origin='Driver/Software' }
    '0x0000001A' = @{ Name='MEMORY_MANAGEMENT';           Origin='Hardware (RAM) or Driver' }
    '0x0000003B' = @{ Name='SYSTEM_SERVICE_EXCEPTION';    Origin='Driver/Software' }
    '0x0000003F' = @{ Name='NO_MORE_SYSTEM_PTES';         Origin='Driver/Software' }
    '0x00000050' = @{ Name='PAGE_FAULT_IN_NONPAGED_AREA'; Origin='Hardware (RAM) or Driver' }
    '0x0000007E' = @{ Name='SYSTEM_THREAD_EXCEPTION_NOT_HANDLED'; Origin='Driver/Software' }
    '0x0000007F' = @{ Name='UNEXPECTED_KERNEL_MODE_TRAP'; Origin='Hardware or Driver' }
    '0x0000009C' = @{ Name='MACHINE_CHECK_EXCEPTION';     Origin='Hardware (CPU/MB)' }
    '0x0000009F' = @{ Name='DRIVER_POWER_STATE_FAILURE';  Origin='Driver/Software' }
    '0x000000C2' = @{ Name='BAD_POOL_CALLER';             Origin='Driver/Software' }
    '0x000000D1' = @{ Name='DRIVER_IRQL_NOT_LESS_OR_EQUAL'; Origin='Driver/Software' }
    '0x000000EF' = @{ Name='CRITICAL_PROCESS_DIED';       Origin='Software (Windows file corruption)' }
    '0x000000F4' = @{ Name='CRITICAL_OBJECT_TERMINATION'; Origin='Software (Windows / disk)' }
    '0x000000F7' = @{ Name='DRIVER_OVERRAN_STACK_BUFFER'; Origin='Driver/Software' }
    '0x00000109' = @{ Name='CRITICAL_STRUCTURE_CORRUPTION'; Origin='Hardware or Malware/Rootkit' }
    '0x00000124' = @{ Name='WHEA_UNCORRECTABLE_ERROR';    Origin='Hardware (CPU/RAM/MB)' }
    '0x00000133' = @{ Name='DPC_WATCHDOG_VIOLATION';      Origin='Driver (often storage)' }
    '0x00000139' = @{ Name='KERNEL_SECURITY_CHECK_FAILURE'; Origin='Driver/Software' }
    '0x0000014F' = @{ Name='PDC_WATCHDOG_TIMEOUT';        Origin='Driver/Hardware' }
    '0x000001E3' = @{ Name='WHEA_INTERNAL_ERROR';         Origin='Hardware' }
}

function Get-BSODDetails {
    # Returns a list of {Time, BugcheckCode, BugcheckName, Origin, Source} for crashes
    # in the last 30 days. Source is parsed from Event 1001 messages (BugCheck event).
    $since   = (Get-Date).AddDays(-30)
    $results = [System.Collections.Generic.List[object]]::new()
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='System'; Id=1001; StartTime=$since} -ErrorAction Stop
        foreach ($e in $events) {
            $msg = "$($e.Message)"
            $code = $null
            if ($msg -match '0x([0-9A-Fa-f]{8,16})')                 { $code = '0x' + $Matches[1].PadLeft(8,'0').ToUpper() }
            elseif ($msg -match 'bugcheck.*?(0x[0-9A-Fa-f]+)')       { $code = $Matches[1].ToUpper() }
            $info = if ($code -and $Script:BugcheckMap.ContainsKey($code)) { $Script:BugcheckMap[$code] }
                    else { @{ Name='UNKNOWN'; Origin='Unknown - check minidump' } }
            $results.Add([PSCustomObject]@{
                Time         = $e.TimeCreated.ToString('yyyy-MM-dd HH:mm')
                BugcheckCode = if ($code) { $code } else { 'unparsed' }
                BugcheckName = $info.Name
                Origin       = $info.Origin
            })
        }
    } catch {}
    return $results
}

function Test-Software {
    Write-Section 'Windows and Software Health'
    $issues = [System.Collections.Generic.List[string]]::new()
    $good   = [System.Collections.Generic.List[string]]::new()
    $d      = [ordered]@{}

    try {
        $t  = $Script:CFG.Thresholds
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop

        $d.OS       = $os.Caption
        $d.Build    = $os.BuildNumber
        $d.LastBoot = $os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm')
        $upDays     = [Math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1)
        $d.Uptime   = "$upDays days"

        # Windows activation — slmgr can hang for minutes on KMS-configured boxes whose KMS host is unreachable
        $slmgrJob = Start-Job { & cscript //nologo "$env:SystemRoot\System32\slmgr.vbs" /dli 2>&1 }
        $slmgr    = if ($slmgrJob | Wait-Job -Timeout 15) { $slmgrJob | Receive-Job } else { '' }
        Remove-Job $slmgrJob -Force
        $activated = ($slmgr -join ' ') -match 'License Status: Licensed'
        $d.WindowsActivation = if ($activated) { 'Licensed' } else { 'CHECK REQUIRED' }
        if ($activated) { $good.Add('Windows is properly licensed') }
        else { $issues.Add('Windows activation could not be confirmed - check license') }

        # OS build
        if ([int]$os.BuildNumber -ge 22000) { $good.Add("Windows 11 (Build $($os.BuildNumber))") }
        elseif ([int]$os.BuildNumber -ge 19041) { $good.Add("Windows 10 current (Build $($os.BuildNumber))") }
        else { $issues.Add("Windows build $($os.BuildNumber) is outdated") }

        # Pending updates
        $updCount = Read-PendingUpdates
        if ($updCount -lt 0) {
            $d.PendingUpdates = 'Timed out - check manually'
            $updCount = 0   # treat as unknown for scoring (no penalty, no bonus)
        } else {
            $d.PendingUpdates = $updCount
            if ($updCount -gt 20)   { $issues.Add("$updCount pending updates - significant backlog") }
            elseif ($updCount -gt 5){ $issues.Add("$updCount pending Windows updates") }
            else                    { $good.Add("Updates current ($updCount pending)") }
        }

        # Driver health  -  only Status='Error' is a real driver problem
        # Status='Unknown' from Get-PnpDevice is a WMI permission artifact, NOT an actual broken device
        # Device Manager is the authoritative source; if it shows no yellow bangs, there are no driver errors
        $allDev  = @(Get-PnpDevice -ErrorAction SilentlyContinue)
        # ConfigManagerErrorCode=22 means user-disabled in Device Manager — not a broken driver
        $failDev = @($allDev | Where-Object { $_.Status -eq 'Error' -and $_.ConfigManagerErrorCode -ne 22 })
        $d.TotalDevices = $allDev.Count
        $d.ErrorDrivers = if ($failDev.Count -gt 0) { ($failDev | ForEach-Object { "$($_.FriendlyName) [$($_.Class)]" }) -join '; ' } else { 'None' }
        if ($failDev.Count -gt 0) { $issues.Add("$($failDev.Count) device(s) with driver errors (yellow bang in Device Manager)") }
        else                      { $good.Add('No driver errors detected') }

        # Antivirus
        $avStr = Read-AVStatus
        $d.Antivirus = $avStr
        if ($avStr -match 'Active|Enabled') { $good.Add("Antivirus: $avStr") }
        else { $issues.Add("Antivirus: $avStr") }

        # Firewall
        $fwOff = @(Get-NetFirewallProfile -ErrorAction SilentlyContinue | Where-Object { -not $_.Enabled })
        $d.Firewall = if ($fwOff.Count -eq 0) { 'Enabled all profiles' } else { "Disabled: $(($fwOff.Name) -join ', ')" }
        if ($fwOff.Count -gt 0) { $issues.Add("Firewall disabled on: $(($fwOff.Name) -join ', ')") }
        else { $good.Add('Windows Firewall enabled on all profiles') }

        # Startup programs
        $su = 0
        foreach ($k in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
                          'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run')) {
            $p = Get-ItemProperty $k -ErrorAction SilentlyContinue
            if ($p) { $su += ($p.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' }).Count }
        }
        $d.StartupPrograms = $su
        if ($su -gt $t.Software_MaxStartup) { $issues.Add("$su startup programs may slow boot") }

        # Dirty volumes — only local fixed drives (DriveType=3); network/removable drives can hang fsutil
        $dirty = [System.Collections.Generic.List[string]]::new()
        $localFixed = @(Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue)
        foreach ($vol in $localFixed) {
            $r = & fsutil dirty query $vol.DeviceID 2>&1
            if ($r -match 'is Dirty') { $dirty.Add($vol.DeviceID) }
        }
        $d.DirtyVolumes = if ($dirty.Count -gt 0) { $dirty -join ', ' } else { 'None' }
        if ($dirty.Count -gt 0) { $issues.Add("Dirty bit set on: $($dirty -join ', ') - CHKDSK pending") }

        # BSOD / crash details (last 30 days) - extract bugcheck codes so the technician
        # can see WHAT crashed, not just how many times. Hardware-origin bugchecks
        # (WHEA, MCE, MEMORY_MANAGEMENT) point to component failure; driver-origin can
        # often be fixed by reinstalling the offending driver and re-running.
        try {
            $bsodList = Get-BSODDetails
            $d.BSOD_Last30Days = $bsodList.Count
            if ($bsodList.Count -gt 0) {
                $d.BSOD_Details = ($bsodList | ForEach-Object { "$($_.Time)  $($_.BugcheckCode) $($_.BugcheckName) [$($_.Origin)]" }) -join ' || '
                $hwBugs = @($bsodList | Where-Object { $_.Origin -match 'Hardware' })
                $swBugs = @($bsodList | Where-Object { $_.Origin -match 'Driver|Software' })
                if ($hwBugs.Count -gt 0) {
                    $unique = ($hwBugs | ForEach-Object { $_.BugcheckName } | Sort-Object -Unique) -join ', '
                    $issues.Add("$($hwBugs.Count) hardware-origin crash(es) in last 30 days: $unique - possible RAM/CPU/board failure, run hardware diagnostics")
                }
                if ($swBugs.Count -gt 0) {
                    $unique = ($swBugs | ForEach-Object { $_.BugcheckName } | Sort-Object -Unique) -join ', '
                    $issues.Add("$($swBugs.Count) driver-origin crash(es) in last 30 days: $unique - review/reinstall problematic drivers, then re-test")
                }
            } else { $good.Add('No system crashes in last 30 days') }
        } catch { $d.BSOD_Last30Days = 0; $d.BSOD_Details = ''; $good.Add('No crash events detected') }

        # Windows Reliability Index (0-10, higher is better)
        try {
            $rel = Get-CimInstance -ClassName Win32_ReliabilityStabilityMetrics -ErrorAction Stop |
                   Sort-Object StartMeasurementDate -Descending | Select-Object -First 1
            if ($rel) {
                $ri = [Math]::Round($rel.SystemStabilityIndex, 1)
                $d.ReliabilityIndex = "$ri / 10"
                if ($ri -ge 8)     { $good.Add("Reliability index $ri/10 - stable system") }
                elseif ($ri -ge 5) { $issues.Add("Reliability index $ri/10 - some instability") }
                else               { $issues.Add("Reliability index $ri/10 - significant instability") }
            }
        } catch { $d.ReliabilityIndex = 'N/A' }

        # Score  -  only penalize confirmed Error devices, not Unknown (Unknown = admin limitation artifact)
        $score = 100
        $score -= ($failDev.Count * 8)
        if ($updCount -gt 20)    { $score -= 20 }
        elseif ($updCount -gt 5) { $score -= 10 }
        if (-not $activated)     { $score -= 20 }
        if ($fwOff.Count -gt 0)  { $score -= 10 }
        if ($dirty.Count -gt 0)  { $score -= 10 }
        $bsodCount = if ($d.BSOD_Last30Days -is [int]) { $d.BSOD_Last30Days } else { 0 }
        if ($bsodCount -gt 3)    { $score -= 20 }
        elseif ($bsodCount -gt 0){ $score -= 10 }

        $score  = Clamp100 $score
        $status = Score-ToStatus $score

        # Build fixable issues - things the technician can resolve and re-test
        $fixables = [System.Collections.Generic.List[object]]::new()
        if ($updCount -gt 5) {
            $fixables.Add( (New-FixableIssue -Title "Install $updCount pending Windows updates" `
                -Reason "Backlog of updates can carry security patches and driver fixes" `
                -FixCommand "UsoClient StartScan ; UsoClient StartDownload ; UsoClient StartInstall  # or open Settings > Windows Update" `
                -EstimatedTime '15-60 min' -RequiresReboot $true) )
        }
        if (-not $activated) {
            $fixables.Add( (New-FixableIssue -Title 'Resolve Windows activation' `
                -Reason 'Windows shows as not licensed - check with licensing team' `
                -FixCommand "slmgr /dlv  # check status / contact licensing" `
                -EstimatedTime '5 min') )
        }
        if ($fwOff.Count -gt 0) {
            $fixables.Add( (New-FixableIssue -Title "Re-enable Windows Firewall on: $(($fwOff.Name) -join ', ')" `
                -Reason "Firewall disabled - required for managed deployment" `
                -FixCommand "Set-NetFirewallProfile -Profile $(($fwOff.Name) -join ',') -Enabled True" `
                -EstimatedTime '1 min') )
        }
        if ($dirty.Count -gt 0) {
            $fixables.Add( (New-FixableIssue -Title "Run CHKDSK on dirty volume(s): $($dirty -join ', ')" `
                -Reason "Dirty bit set - file system needs verification" `
                -FixCommand "chkdsk /f $($dirty[0])  # repeat per volume; reboot may be required" `
                -EstimatedTime '5-30 min' -RequiresReboot $true) )
        }
        if ($failDev.Count -gt 0) {
            $fixables.Add( (New-FixableIssue -Title "Fix $($failDev.Count) device(s) with driver errors" `
                -Reason "Devices showing yellow bang in Device Manager - re-install or update drivers from manufacturer" `
                -FixCommand "Open Device Manager, right-click each device with a yellow icon, choose 'Update driver' or 'Uninstall device' then reboot. Failing devices: $($d.ErrorDrivers)" `
                -EstimatedTime '10-30 min' -RequiresReboot $true) )
        }
        if ($avStr -notmatch 'Active|Enabled') {
            $fixables.Add( (New-FixableIssue -Title 'Enable antivirus protection' `
                -Reason "Antivirus not active: $avStr" `
                -FixCommand "Set-MpPreference -DisableRealtimeMonitoring `$false  # for Defender" `
                -EstimatedTime '2 min') )
        }

        $Script:RESULTS.Software = New-Result -Module Software -Score $score -Status $status `
            -Summary "$($d.OS) (Build $($d.Build)) | Drivers: $($failDev.Count) err | Updates: $updCount | Crashes: $bsodCount" `
            -Data $d -Issues $issues -Good $good -Fixables $fixables

        Write-Log "Software Score: $score ($status)" 'OK'
    } catch {
        Write-Log "Software error: $_" 'ERROR'
        $Script:RESULTS.Software = New-Result -Module Software -Score 0 -Status ERROR -Summary "Error: $_"
    }
}

function Read-PendingUpdates {
    # The Microsoft.Update.Session COM call can hang for minutes against WSUS on servers
    # or air-gapped boxes. Run it in a job with a 30s ceiling. If it doesn't return in time,
    # we report 'unknown' and move on rather than freezing the whole evaluation.
    $job = Start-Job -ScriptBlock {
        try {
            $s = New-Object -ComObject Microsoft.Update.Session -ErrorAction Stop
            $r = $s.CreateUpdateSearcher().Search("IsInstalled=0 and Type='Software'")
            return $r.Updates.Count
        } catch { return 0 }
    }
    if (Wait-Job $job -Timeout 30) {
        $count = Receive-Job $job
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        return $count
    } else {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        Write-Log 'Pending updates check timed out (>30s) - skipping' 'WARN'
        return -1
    }
}

function Read-AVStatus {
    try {
        $av = Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction Stop
        if (-not $av) { return 'Not detected' }
        $h    = '{0:X6}' -f [int]($av | Select-Object -First 1).productState
        $en   = $h.Substring(2,2) -ne '00'
        $utd  = $h.Substring(4,2) -eq '00'
        $name = ($av | Select-Object -First 1).displayName
        return "$name - $(if ($en){'Active'}else{'Inactive'}), $(if ($utd){'Up-to-date'}else{'Out-of-date'})"
    } catch { return 'Status unavailable' }
}

# ============================================================
#  MODULE: SECURITY
# ============================================================

function Test-Security {
    Write-Section 'Security Assessment'
    $issues = [System.Collections.Generic.List[string]]::new()
    $good   = [System.Collections.Generic.List[string]]::new()
    $d      = [ordered]@{}
    $score  = 70

    try {
        # TPM version
        try {
            $tpmWmi = Get-WmiObject -Namespace 'root\cimv2\security\microsofttpm' -Class Win32_Tpm -ErrorAction Stop | Select-Object -First 1
            if ($tpmWmi -and $tpmWmi.IsEnabled_InitialValue) {
                $tpmVerStr = $tpmWmi.SpecVersion
                $tpmVer = if ($tpmVerStr -match '2\.0') { '2.0' }
                          elseif ($tpmVerStr -match '1\.2') { '1.2' }
                          else { $tpmVerStr }
                $d.TPM_Version = $tpmVer
                $d.TPM_Enabled = $true
                if ($tpmVer -eq '2.0') {
                    $good.Add('TPM 2.0 present and enabled (Windows 11 / BitLocker ready)')
                    $score += 10
                } else {
                    $issues.Add("TPM $tpmVer detected - TPM 2.0 required for Windows 11")
                    $score -= 10
                }
            } else {
                $d.TPM_Version = 'Disabled or absent'
                $d.TPM_Enabled = $false
                $issues.Add('TPM not enabled - required for Windows 11 and BitLocker')
                $score -= 20
            }
        } catch {
            # Fallback 1: Get-Tpm cmdlet
            $tpmBasic = Get-Tpm -ErrorAction SilentlyContinue
            # Fallback 2: PnP device enumeration (works without admin)
            # Status may be 'Unknown' without admin  -  don't filter by OK, just find the device
            # Select-First 1 ensures scalar even when Intel PTT + discrete TPM both appear
            $tpmPnP = Get-PnpDevice -Class SecurityDevices -ErrorAction SilentlyContinue |
                      Where-Object { $_.FriendlyName -match 'TPM|Trusted Platform' -and $_.Status -ne 'Error' } |
                      Select-Object -First 1
            if ($tpmBasic -and $tpmBasic.TpmPresent -and $tpmBasic.TpmEnabled) {
                $tpmVerPnP = if ($tpmPnP -and $tpmPnP.FriendlyName -match '2\.0') { '2.0' } elseif ($tpmPnP -and $tpmPnP.FriendlyName -match '1\.2') { '1.2' } else { 'Detected' }
                $d.TPM_Version = $tpmVerPnP
                $d.TPM_Enabled = $true
                if ($tpmVerPnP -eq '2.0') {
                    $good.Add('TPM 2.0 present and enabled')
                    $score += 10
                } else {
                    $good.Add("TPM $tpmVerPnP present and enabled")
                    $score += 5
                }
            } elseif ($tpmPnP) {
                $tpmVerPnP = if ($tpmPnP.FriendlyName -match '2\.0') { '2.0' } elseif ($tpmPnP.FriendlyName -match '1\.2') { '1.2' } else { 'Detected' }
                $d.TPM_Version = $tpmVerPnP
                $d.TPM_Enabled = $true
                $good.Add("TPM $tpmVerPnP detected via device manager")
                if ($tpmVerPnP -eq '2.0') { $score += 10 } else { $score += 5 }
            } else {
                $d.TPM_Version = 'Not detected'
                $d.TPM_Enabled = $false
                $issues.Add('TPM not detected - required for Windows 11 and BitLocker')
                $score -= 15
            }
        }

        # Secure Boot
        try {
            $sb = Confirm-SecureBootUEFI -ErrorAction Stop
            $d.SecureBoot = if ($sb) { 'Enabled' } else { 'Disabled' }
            if ($sb) { $good.Add('Secure Boot enabled'); $score += 5 }
            else     { $issues.Add('Secure Boot is disabled'); $score -= 10 }
        } catch { $d.SecureBoot = 'Cannot determine (BIOS mode or non-UEFI)' }

        # BitLocker
        try {
            $bl = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
            $d.BitLocker = $bl.ProtectionStatus
            $d.BitLocker_Volume = $bl.VolumeStatus
            if ($bl.ProtectionStatus -eq 'On') {
                $good.Add("BitLocker encryption active on $env:SystemDrive")
                $score += 10
            } else {
                $issues.Add("BitLocker not active on $env:SystemDrive - drive is unencrypted")
                $score -= 5
            }
        } catch { $d.BitLocker = 'Cannot determine' }

        # BIOS/firmware age
        $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
        if ($bios -and $bios.ReleaseDate) {
            $biosAgeDays = [int]((Get-Date) - $bios.ReleaseDate).TotalDays
            $biosAgeYr   = [Math]::Round($biosAgeDays / 365, 1)
            $d.BIOS_Version  = $bios.SMBIOSBIOSVersion
            $d.BIOS_Date     = $bios.ReleaseDate.ToString('yyyy-MM-dd')
            $d.BIOS_Age      = "$biosAgeYr years"
            if ($biosAgeDays -gt 1095) {
                $issues.Add("BIOS firmware is $biosAgeYr years old - check manufacturer for updates")
                $score -= 5
            } else { $good.Add("BIOS firmware is $biosAgeYr years old") }
        }

        # Virtualization (VT-x / AMD-V)  -  WMI property unreliable without admin; only penalize if admin confirms disabled
        $cpuVirt = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cpuVirt) {
            $virtEnabled = $cpuVirt.VirtualizationFirmwareEnabled
            if ($virtEnabled) {
                $d.Virtualization = 'Enabled'
                $good.Add('CPU virtualization (VT-x / AMD-V) enabled - supports Hyper-V / Docker')
                $score += 5
            } elseif ($Script:IS_ADMIN) {
                $d.Virtualization = 'Disabled (confirmed)'
                $issues.Add('CPU virtualization disabled in BIOS/UEFI - Hyper-V and Docker will not work')
                $score -= 5
            } else {
                $d.Virtualization = 'Unknown (run as admin to confirm)'
            }
        }

        # Intel vPro (remote management)
        $cpuName = if ($cpuVirt) { $cpuVirt.Name } else { '' }
        $isVPro = $cpuName -match 'vPro'
        $d.Intel_vPro = if ($isVPro) { 'Detected' } else { 'Not detected' }
        if ($isVPro) { $good.Add('Intel vPro detected - remote management (AMT) capable') }

        # Hyper-V feature
        try {
            $hv = Get-WindowsOptionalFeature -Online -FeatureName 'Microsoft-Hyper-V' -ErrorAction Stop
            $d.HyperV = $hv.State
            if ($hv.State -eq 'Enabled') { $good.Add('Hyper-V feature enabled') }
        } catch { $d.HyperV = 'N/A' }

        $score  = Clamp100 $score
        $status = Score-ToStatus $score
        $Script:RESULTS.Security = New-Result -Module Security -Score $score -Status $status `
            -Summary "TPM: $($d.TPM_Version) | SecureBoot: $($d.SecureBoot) | BitLocker: $($d.BitLocker)" `
            -Data $d -Issues $issues -Good $good

        Write-Log "Security Score: $score ($status)" 'OK'
    } catch {
        Write-Log "Security error: $_" 'ERROR'
        $Script:RESULTS.Security = New-Result -Module Security -Score 0 -Status ERROR -Summary "Error: $_"
    }
}

# ============================================================
#  MODULE: THERMAL STRESS
# ============================================================
# Scoring is now based on observable signals that don't require kernel-level temperature
# probes: CPU clock-drop under load (throttling) and ACPI thermal-zone delta. These are
# enough to flag a laptop with failing cooling without needing third-party DLLs.

function Test-Thermal {
    Write-Section 'Thermal Stress Test'
    $issues = [System.Collections.Generic.List[string]]::new()
    $good   = [System.Collections.Generic.List[string]]::new()
    $d      = [ordered]@{}

    try {
        $t = $Script:CFG.Thresholds

        Write-Log 'Sampling idle temperature (5s)...'
        Start-Sleep -Seconds 5
        $idle = Read-CPUTemp
        $d.IdleTemp = if ($idle -gt 0) { "$($idle) C (ACPI)" } else { 'N/A' }
        if ($idle -gt 0) { Write-Log "Idle: $($idle) C (ACPI thermal zone)" }

        # Baseline max clock before stress
        $cpuPre      = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        $maxClockMHz = if ($cpuPre) { $cpuPre.MaxClockSpeed } else { 0 }

        Write-Log 'Stress test starting (30 seconds, all cores)...'
        $stress = Start-StressJob -Seconds 30

        $peaks  = [System.Collections.Generic.List[int]]::new()
        $loads  = [System.Collections.Generic.List[int]]::new()
        $clocks = [System.Collections.Generic.List[int]]::new()
        for ($i = 0; $i -lt 10; $i++) {
            Start-Sleep -Seconds 3
            $tmp  = Read-CPUTemp
            $cpuS = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
            $load = if ($cpuS) { [int]$cpuS.LoadPercentage } else { 0 }
            $clock= if ($cpuS -and $cpuS.CurrentClockSpeed -gt 0) { [int]$cpuS.CurrentClockSpeed } else { 0 }
            if ($tmp   -gt 0) { $peaks.Add($tmp) }
            if ($clock -gt 0) { $clocks.Add($clock) }
            $loads.Add($load)
            Write-Log ("  t+$([int](($i+1)*3))s  Temp: $($tmp) C  Load: $($load)%  Clock: $($clock) MHz")
        }
        Stop-StressJob $stress
        Write-Log 'Stress test complete.'

        $peak    = if ($peaks.Count  -gt 0) { ($peaks  | Measure-Object -Maximum).Maximum } else { 0 }
        $minClk  = if ($clocks.Count -gt 0) { ($clocks | Measure-Object -Minimum).Minimum } else { 0 }
        $avgLoad = if ($loads.Count  -gt 0) { [Math]::Round(($loads | Measure-Object -Average).Average, 0) } else { 0 }
        $rise    = if ($idle -gt 0 -and $peak -gt 0) { $peak - $idle } else { 0 }

        # Throttling: minimum observed clock dropped >20% below the rated max during stress
        $throttled = $maxClockMHz -gt 0 -and $minClk -gt 0 -and $minClk -lt ($maxClockMHz * 0.80)
        $d.Throttling = if ($throttled)         { "YES - dropped to $($minClk) MHz (max $($maxClockMHz) MHz)" }
                        elseif ($minClk -gt 0)  { "No (min $($minClk) MHz / max $($maxClockMHz) MHz)" }
                        else                    { 'N/A' }
        $d.PeakTemp = if ($peak -gt 0) { "$($peak) C (ACPI)" } else { 'N/A' }
        $d.TempRise = if ($rise -gt 0) { "$($rise) C" }        else { 'N/A' }
        $d.AvgLoad  = "$($avgLoad)%"

        # Score is built primarily from CLOCK BEHAVIOUR — that signal is reliable on every
        # machine. Temperature is only used as a secondary check because ACPI thermal zones
        # are often stuck on laptops (single-value sensor, doesn't move during stress).
        $score = 80

        # Throttling = the single most important thermal failure signal
        if ($throttled) {
            $issues.Add("CPU throttling detected under load ($($d.Throttling)) - cooling or power limit issue")
            $score -= 25
        } elseif ($minClk -gt 0) {
            $good.Add("No throttling detected - sustained $($minClk) MHz under full load")
        }

        # Sustained-load sanity: stress should pin all cores high. If average load stayed low,
        # the stress job didn't actually run (rare, but worth flagging).
        if ($avgLoad -lt 50 -and $loads.Count -gt 0) {
            $issues.Add("Stress test only reached $($avgLoad)% average load - thermal headroom not validated")
            $score -= 10
        } elseif ($avgLoad -ge 80) {
            $good.Add("Stress test held $($avgLoad)% average load across all samples")
        }

        # Temperature checks — informational when sensor is stuck, real when it moves
        if ($peak -gt 0 -and $idle -gt 0 -and $rise -ge 5) {
            # Sensor is moving — trust it
            if ($peak -gt $t.CPU_MaxLoadTemp) {
                $issues.Add("Peak $($peak) C exceeds critical threshold $($t.CPU_MaxLoadTemp) C")
                $score -= 20
            } elseif ($peak -gt ($t.CPU_MaxLoadTemp - 5)) {
                $issues.Add("Peak $($peak) C is dangerously close to throttle threshold")
                $score -= 10
            } else {
                $good.Add("Peak $($peak) C within safe limits under full load")
            }
            if ($idle -gt ($t.CPU_MaxIdleTemp + 2)) {
                $issues.Add("Idle $($idle) C exceeds threshold $($t.CPU_MaxIdleTemp) C - check thermal paste/cooling")
                $score -= 10
            }
            if ($rise -gt 40) {
                $issues.Add("Temperature rise $($rise) C - cooling may be undersized")
            }
        } elseif ($peak -gt 0) {
            $d.Note = "ACPI sensor returned a static value ($peak C) for the whole run - typical of laptops with thermal zones reporting one number; throttling check is the authoritative signal here"
            Write-Log $d.Note 'WARN'
        } else {
            $d.Note = 'No CPU temperature available (no ACPI thermal zone exposed)'
            Write-Log 'No CPU temperature available - throttling check is the authoritative signal' 'WARN'
        }

        $score  = Clamp100 $score
        $status = Score-ToStatus $score
        $Script:RESULTS.Thermal = New-Result -Module Thermal -Score $score -Status $status `
            -Summary "Idle: $($d.IdleTemp) | Peak: $($d.PeakTemp) | Rise: $($d.TempRise)" `
            -Data $d -Issues $issues -Good $good

        Write-Log "Thermal Score: $score ($status)" 'OK'
    } catch {
        Write-Log "Thermal error: $_" 'ERROR'
        $Script:RESULTS.Thermal = New-Result -Module Thermal -Score 0 -Status ERROR -Summary "Error: $_"
    }
}

function Start-StressJob {
    param([int]$Seconds = 30)
    $n    = [Environment]::ProcessorCount
    $pool = [RunspaceFactory]::CreateRunspacePool(1, $n)
    $pool.Open()
    $end  = (Get-Date).AddSeconds($Seconds)

    $jobs = 1..$n | ForEach-Object {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript({
            param($e)
            while ((Get-Date) -lt $e) {
                $x = [Math]::PI
                for ($i = 0; $i -lt 100000; $i++) { $x = [Math]::Sqrt([Math]::Abs($x * $i + 1.23456)) }
            }
        }).AddArgument($end)
        @{ P = $ps; H = $ps.BeginInvoke() }
    }
    return @{ Pool = $pool; Jobs = $jobs }
}

function Stop-StressJob {
    param($j)
    foreach ($item in $j.Jobs) {
        try { $item.P.EndInvoke($item.H) } catch {}
        $item.P.Dispose()
    }
    try { $j.Pool.Close(); $j.Pool.Dispose() } catch {}
}

# ============================================================
#  SCORING ENGINE
# ============================================================

function Invoke-Scoring {
    Write-Section 'Calculating Overall Score'

    # Manual copy  -  OrderedDictionary has no .Clone() in PS 5.1
    $weights = [ordered]@{}
    foreach ($k in $Script:CFG.Weights.Keys) { $weights[$k] = $Script:CFG.Weights[$k] }

    # Redistribute battery weight if no battery (desktop)
    if (-not $Script:CAPS.HasBattery) {
        $bw = $weights.Battery
        $weights.Battery  = 0
        $weights.Security = $weights.Security + [int]($bw / 2)
        $weights.Storage  = $weights.Storage  + ($bw - [int]($bw / 2))
    }

    $wsum    = 0.0
    $wtotal  = 0
    $details = [ordered]@{}

    foreach ($mod in $weights.Keys) {
        $w = $weights[$mod]
        if ($w -le 0) { continue }
        $r = $Script:RESULTS[$mod]
        if (-not $r -or $r.Status -eq 'SKIP') { continue }
        $s = $r.Score
        $wsum   += $s * $w
        $wtotal += $w
        $details[$mod] = [PSCustomObject]@{
            Score  = $s
            Weight = $w
            Contribution = [Math]::Round($s * $w / 100, 1)
        }
        Write-Log ("{0,-12} {1,3}/100  w:{2,3}%  contrib:{3,5:F1}" -f $mod, $s, $w, ($s * $w / 100))
    }

    $overall = if ($wtotal -gt 0) { Clamp100 ($wsum / $wtotal) } else { 0 }

    # ── Veto rules: catastrophic single-component failures cap the verdict regardless of average ──
    $vetos = [System.Collections.Generic.List[string]]::new()
    $batt  = $Script:RESULTS['Battery']
    if ($batt -and $batt.Status -ne 'SKIP' -and $batt.Data) {
        $hpRaw = "$($batt.Data.HealthPct)" -replace '[^0-9]',''
        $fwRaw = "$($batt.Data.FullChargeWh)"
        if ($hpRaw -and ([int]$hpRaw) -lt 50)              { $vetos.Add("Battery health $hpRaw% - laptop cannot run unplugged reliably") }
        if ($fwRaw -and ([double]$fwRaw) -gt 0 -and ([double]$fwRaw) -lt 25) { $vetos.Add("Battery full charge under 25 Wh - laptop is essentially desktop-bound") }
    }
    $stor = $Script:RESULTS['Storage']
    if ($stor -and $stor.Data) {
        # SMART failure
        if ($stor.Data.Drives) {
            foreach ($drv in $stor.Data.Drives) {
                if ("$($drv.SMART)" -match 'FAILURE PREDICTED|FAIL') { $vetos.Add("SSD/HDD SMART failing on $($drv.Name) - data loss imminent") }
            }
        }
        # Primary drive is HDD - Windows 11 effectively unusable
        $primaryDrive = $stor.Data.Drives | Where-Object { $_.BusType -ne 'USB' } | Select-Object -First 1
        if ($primaryDrive -and $primaryDrive.MediaType -eq 'HDD') {
            $vetos.Add("Primary drive is HDD ($($primaryDrive.Name)) - must be replaced with SSD before deployment")
        }
    }
    # Software-side vetos: driver errors and frequent crashes
    $sw = $Script:RESULTS['Software']
    if ($sw -and $sw.Data) {
        $bsodCount = if ($sw.Data.BSOD_Last30Days -is [int]) { $sw.Data.BSOD_Last30Days } else { 0 }
        if ($bsodCount -gt 5) { $vetos.Add("$bsodCount system crashes in last 30 days - hardware/driver instability, do not deploy until root cause is fixed") }
        if ($sw.Data.ErrorDrivers -and "$($sw.Data.ErrorDrivers)" -ne 'None') {
            # Count semicolon separators to count error devices
            $errCount = ([regex]::Matches("$($sw.Data.ErrorDrivers)", ';')).Count + 1
            if ($errCount -gt 5) { $vetos.Add("$errCount devices with driver errors - investigate / repair drivers before deployment") }
        }
        $relStr = "$($sw.Data.ReliabilityIndex)"
        if ($relStr -match '^([\d\.]+)') {
            $rel = [double]$Matches[1]
            if ($rel -gt 0 -and $rel -lt 4) { $vetos.Add("Windows Reliability Index $rel/10 - system is unstable, investigate before deployment") }
        }
    }

    $v = $Script:CFG.Verdict

    # As-is verdict (cap at Adequate if any veto rules tripped)
    $vText = if ($overall -ge $v.Excellent) { 'Highly Recommended' }
             elseif ($overall -ge $v.Good)      { 'Recommended' }
             elseif ($overall -ge $v.Adequate)  { 'Conditionally Recommended' }
             elseif ($overall -ge $v.Marginal)  { 'Use Case Dependent' }
             else { 'Not Recommended' }
    if ($vetos.Count -gt 0 -and $overall -ge $v.Good) {
        $vText = 'Conditionally Recommended (veto: critical component issue)'
    }

    $vColor = if ($overall -ge $v.Good -and $vetos.Count -eq 0) { '#22c55e' }
              elseif ($overall -ge $v.Adequate) { '#f59e0b' } else { '#ef4444' }

    $crits = @($Script:RESULTS.Values | Where-Object { $null -ne $_ -and $_.Status -in @('FAIL','ERROR') } |
               ForEach-Object { $_.Module })

    $Script:RESULTS.Overall = [PSCustomObject]@{
        Score   = $overall
        Verdict = $vText
        VColor  = $vColor
        Details = $details
        Crits   = $crits
        Vetos   = $vetos
    }

    $col = if ($overall -ge $v.Good -and $vetos.Count -eq 0) { 'Green' } elseif ($overall -ge $v.Adequate) { 'Yellow' } else { 'Red' }
    Write-Host ''
    Write-Host "  +----------------------------------------------+" -ForegroundColor White
    Write-Host ("  |  OVERALL: {0,3}/100   {1,-35}|" -f $overall, $vText) -ForegroundColor $col
    Write-Host "  +----------------------------------------------+" -ForegroundColor White
    Write-Host ''
    foreach ($vt in $vetos) { Write-Log "VETO: $vt" 'WARN' }
    Write-Log "Overall: $overall/100 - $vText" 'OK'
}

# ============================================================
#  DEPARTMENT / USE-CASE SCORING
# ============================================================

function Get-DepartmentScores {
    function priv-gs([string]$m) {
        if ($Script:RESULTS.Contains($m) -and $Script:RESULTS[$m]) { return [int]$Script:RESULTS[$m].Score }
        return 50
    }
    function priv-gd([string]$m,[string]$k) {
        if ($Script:RESULTS.Contains($m) -and $Script:RESULTS[$m] -and $Script:RESULTS[$m].Data) {
            $v = $Script:RESULTS[$m].Data[$k]; if ($null -ne $v) { return "$v" }
        }
        return ''
    }

    $ms = @{
        CPU=priv-gs 'CPU'; RAM=priv-gs 'RAM'; Storage=priv-gs 'Storage'; Battery=priv-gs 'Battery'
        GPU=priv-gs 'GPU'; Network=priv-gs 'Network'; Security=priv-gs 'Security'
        Ports=priv-gs 'Ports'; Software=priv-gs 'Software'
    }

    $ramGB     = try { [int][double](priv-gd 'RAM' 'TotalGB') } catch { 0 }
    $cpuTier   = priv-gd 'CPU' 'CPU_Tier'
    $cpuGenStr = priv-gd 'CPU' 'CPU_Generation'
    $cpuGen    = if ($cpuGenStr -match '(\d+)') { [int]$Matches[1] } else { 0 }
    $gpuScore  = $ms['GPU']
    $hasDGPU   = $Script:CAPS.HasDiscreteGPU
    $storIsNVMe = $false; $storIsSSD = $false; $storCapGB = 0
    try {
        $driveList = $Script:RESULTS['Storage'].Data['Drives']
        if ($driveList) {
            foreach ($drv in @($driveList)) {
                $bus = "$($drv['BusType'])"; $med = "$($drv['MediaType'])"
                if ($bus -match 'NVMe')                       { $storIsNVMe = $true }
                if ($bus -match 'NVMe' -or $med -match 'SSD') { $storIsSSD  = $true }
                $cap = try { [int]"$($drv['EffectiveCapGB'])" } catch { 0 }
                if ($cap -gt $storCapGB) { $storCapGB = $cap }
            }
        }
    } catch {}
    $readMBs   = try { [double](priv-gd 'Storage' 'ReadMBs') } catch { 0 }
    $resStr    = priv-gd 'GPU' 'PrimaryRes'
    $resW      = if ($resStr -match '^(\d+)x') { [int]$Matches[1] } else { 0 }
    $tpmVer    = priv-gd 'Security' 'TPM_Version'
    $hasTPM2   = $tpmVer -match '^2\.'
    $sBoot     = (priv-gd 'Security' 'SecureBoot') -match 'True|Enabled'
    $bLock     = (priv-gd 'Security' 'BitLocker')  -match '^On$|Encrypted|Protected'
    $dlRaw     = (priv-gd 'Network' 'DL_Mbps') -replace '[^\d\.]'
    $dlMbps    = if ($dlRaw) { try { [double]$dlRaw } catch { 0 } } else { 0 }
    $fwOK      = -not ((priv-gd 'Software' 'Firewall') -match 'Disabled')

    function priv-calc([hashtable]$w) {
        $t=0; $s=0
        foreach ($k in $w.Keys) { $wv=$w[$k]; if ($wv -gt 0) { $t += $ms[$k]*$wv; $s += $wv } }
        if ($s -eq 0) { return 50 }
        return [Math]::Round($t / $s, 0)
    }
    function priv-note([string]$n) {
        # Returns 'ok' if the note describes a positive/passing state, 'ng' otherwise
        if ($n -match 'comfortable|good|suitable|fast|smooth|present|active|enabled|handles|confirmed|adequate|meets|sufficient|NVMe|TPM 2\.0|BitLocker active|Dedicated GPU') { return 'ok' }
        return 'ng'
    }

    $out = [ordered]@{}

    # ── 1. General Use (includes minimum viability) ────────────────────────────
    # This is the baseline: email, Teams calls, browser, MS 365, video calls
    $p = @{Name='General Use';Score=priv-calc @{CPU=20;RAM=25;Storage=10;Battery=20;Network=15;GPU=5;Security=5}
           Notes=[System.Collections.Generic.List[string]]::new()}
    if     ($ramGB -lt 8)  { $p.Score=[Math]::Max(0,$p.Score-35); $p.Notes.Add("Only ${ramGB} GB RAM  -  cannot reliably run Teams + browser + Office at the same time") }
    elseif ($ramGB -ge 16) { $p.Notes.Add("${ramGB} GB RAM  -  comfortable for everyday multitasking") }
    else                   { $p.Notes.Add("8 GB RAM  -  adequate for basic use; avoid many open browser tabs") }
    if (-not $storIsSSD)   { $p.Score=[Math]::Max(0,$p.Score-25); $p.Notes.Add("HDD  -  boot, login, and app launch are noticeably slow for daily use") }
    if ($dlMbps -gt 0 -and $dlMbps -lt 10) { $p.Score=[Math]::Max(0,$p.Score-15); $p.Notes.Add("Network under 10 Mbps  -  Teams/Zoom calls will be poor quality") }
    elseif ($dlMbps -ge 25){ $p.Notes.Add("Network good for HD video calls and cloud apps") }
    if ($cpuGen -gt 0 -and $cpuGen -lt 6) { $p.Score=[Math]::Max(0,$p.Score-15); $p.Notes.Add("$cpuGenStr  -  too old; modern browsers and Teams will feel sluggish") }
    $out['General Use'] = $p

    # ── 2. Software Development ────────────────────────────────────────────────
    $p = @{Name='Software Dev';Score=priv-calc @{CPU=30;RAM=30;Storage=25;Network=10;Security=5}
           Notes=[System.Collections.Generic.List[string]]::new()}
    if     ($ramGB -lt 16) { $p.Score=[Math]::Max(0,$p.Score-20); $p.Notes.Add("Under 16 GB RAM  -  IDE + containers + browser will compete for memory") }
    elseif ($ramGB -ge 32) { $p.Score=[Math]::Min(100,$p.Score+5); $p.Notes.Add("${ramGB} GB RAM  -  comfortable for VMs and multiple containers") }
    else                   { $p.Score=[Math]::Max(0,$p.Score-10); $p.Notes.Add("16 GB RAM  -  tight for Docker + IDE + browser simultaneously; 32 GB recommended") }
    if ($storIsNVMe -and $readMBs -ge 2000) { $p.Notes.Add("NVMe ${readMBs} MB/s  -  fast builds and container operations") }
    elseif (-not $storIsNVMe) { $p.Score=[Math]::Max(0,$p.Score-10); $p.Notes.Add("Non-NVMe storage  -  build times and package installs noticeably slower") }
    if ($storCapGB -gt 0 -and $storCapGB -lt 256) { $p.Score=[Math]::Max(0,$p.Score-15); $p.Notes.Add("Under 256 GB  -  tight for repos, node_modules, and container images") }
    elseif ($storCapGB -ge 512) { $p.Notes.Add("${storCapGB} GB  -  adequate space for dev environments") }
    if ($cpuTier -match 'i7|i9|Ryzen 7|Ryzen 9') { $p.Score=[Math]::Min(100,$p.Score+5); $p.Notes.Add("$cpuTier  -  handles parallel compilation well") }
    $out['Software Dev'] = $p

    # ── 3. Graphics / Design ──────────────────────────────────────────────────
    # GPU quality matters most here  -  integrated GPU is a real limitation, not just a missing bonus
    $p = @{Name='Graphics / Design';Score=priv-calc @{GPU=35;CPU=20;RAM=25;Storage=15;Battery=5}
           Notes=[System.Collections.Generic.List[string]]::new()}
    # GPU: integrated GPU is a hard limitation  -  penalize significantly
    if ($hasDGPU) {
        if     ($gpuScore -ge 80) { $p.Score=[Math]::Min(100,$p.Score+10); $p.Notes.Add("Dedicated GPU  -  GPU-accelerated rendering and effects available") }
        elseif ($gpuScore -ge 60) { $p.Notes.Add("Dedicated GPU  -  adequate for most design tools, not heavy 3D") }
        else                      { $p.Score=[Math]::Max(0,$p.Score-5);    $p.Notes.Add("Dedicated GPU present but low-end  -  limited GPU acceleration benefit") }
    } else {
        $p.Score=[Math]::Max(0,$p.Score-25)
        $p.Notes.Add("Integrated GPU  -  GPU-accelerated filters, 3D previews, and export will be slow")
    }
    # Display resolution
    if     ($resW -ge 3840) { $p.Score=[Math]::Min(100,$p.Score+10); $p.Notes.Add("4K display  -  excellent for colour-accurate detail work") }
    elseif ($resW -ge 2560) { $p.Score=[Math]::Min(100,$p.Score+5);  $p.Notes.Add("QHD display  -  good canvas size for design work") }
    elseif ($resW -gt 0 -and $resW -lt 1920) { $p.Score=[Math]::Max(0,$p.Score-15); $p.Notes.Add("Below 1080p  -  limited workspace for design applications") }
    # RAM
    if     ($ramGB -lt 16)  { $p.Score=[Math]::Max(0,$p.Score-15); $p.Notes.Add("Under 16 GB RAM  -  large files and multi-app workflows will struggle") }
    elseif ($ramGB -ge 32)  { $p.Score=[Math]::Min(100,$p.Score+5); $p.Notes.Add("${ramGB} GB RAM  -  handles large layered files comfortably") }
    else                    { $p.Score=[Math]::Max(0,$p.Score-10); $p.Notes.Add("16 GB RAM  -  workable but 32 GB recommended for large Photoshop/Illustrator/After Effects files") }
    $out['Graphics / Design'] = $p

    # ── 4. Video Editing ──────────────────────────────────────────────────────
    $p = @{Name='Video Editing';Score=priv-calc @{CPU=30;RAM=25;Storage=25;GPU=20}
           Notes=[System.Collections.Generic.List[string]]::new()}
    if     ($ramGB -lt 16) { $p.Score=[Math]::Max(0,$p.Score-25); $p.Notes.Add("Under 16 GB RAM  -  video editing severely limited, constant proxies required") }
    elseif ($ramGB -lt 32) { $p.Score=[Math]::Max(0,$p.Score-10); $p.Notes.Add("16 GB RAM  -  workable for 1080p; 4K multicam will struggle") }
    else                   { $p.Notes.Add("${ramGB} GB RAM  -  comfortable for 4K and multicam timelines") }
    if ($storCapGB -gt 0 -and $storCapGB -lt 512) { $p.Score=[Math]::Max(0,$p.Score-10); $p.Notes.Add("Under 512 GB  -  very limited space for video project files") }
    if ($readMBs -gt 0 -and $readMBs -lt 500) { $p.Score=[Math]::Max(0,$p.Score-15); $p.Notes.Add("Slow storage  -  scrubbing high-bitrate footage will lag") }
    elseif ($readMBs -ge 2000) { $p.Notes.Add("NVMe ${readMBs} MB/s  -  smooth scrubbing even with high-bitrate RAW") }
    if ($hasDGPU) {
        if ($gpuScore -ge 70) { $p.Score=[Math]::Min(100,$p.Score+10); $p.Notes.Add("Dedicated GPU  -  hardware encode/decode acceleration available") }
        else                  { $p.Notes.Add("Dedicated GPU present but limited  -  some hardware acceleration available") }
    } else {
        $p.Score=[Math]::Max(0,$p.Score-15); $p.Notes.Add("No dedicated GPU  -  export times will be CPU-only, significantly slower")
    }
    if ($cpuGen -gt 0 -and $cpuGen -lt 10) { $p.Score=[Math]::Max(0,$p.Score-10); $p.Notes.Add("$cpuGenStr  -  older CPU; limited modern encoder support (AV1, HEVC HW)") }
    $out['Video Editing'] = $p

    # ── 5. IT Department ──────────────────────────────────────────────────────
    $p = @{Name='IT Department';Score=priv-calc @{Security=30;Network=20;CPU=15;RAM=15;Battery=10;Ports=10}
           Notes=[System.Collections.Generic.List[string]]::new()}
    if     ($ramGB -ge 32) { $p.Score=[Math]::Min(100,$p.Score+5); $p.Notes.Add("${ramGB} GB RAM  -  comfortable for VMs and multiple RDP/remote sessions") }
    elseif ($ramGB -lt 16) { $p.Score=[Math]::Max(0,$p.Score-10); $p.Notes.Add("Under 16 GB RAM  -  limited for running VMs and concurrent remote management tasks") }
    else                   { $p.Notes.Add("16 GB RAM  -  adequate for most IT tasks; 32 GB recommended if running VMs or Hyper-V") }
    if ($hasTPM2)                            { $p.Score=[Math]::Min(100,$p.Score+10); $p.Notes.Add("TPM 2.0  -  BitLocker, Windows Hello, and MDM attestation supported") }
    elseif (-not ($tpmVer -match '1\.|2\.')) { $p.Score=[Math]::Max(0,$p.Score-20);  $p.Notes.Add("No TPM  -  cannot fully enroll in Intune/MDM with security compliance") }
    if ($sBoot)  { $p.Score=[Math]::Min(100,$p.Score+5);  $p.Notes.Add("Secure Boot enabled  -  firmware integrity protected") }
    else         { $p.Score=[Math]::Max(0,$p.Score-5);    $p.Notes.Add("Secure Boot disabled  -  does not meet standard managed-device policy") }
    if ($bLock)  { $p.Score=[Math]::Min(100,$p.Score+5);  $p.Notes.Add("BitLocker active  -  data-at-rest encryption enforced") }
    if (-not $fwOK) { $p.Score=[Math]::Max(0,$p.Score-10); $p.Notes.Add("Firewall disabled on all profiles  -  non-compliant for managed deployment") }
    if ($dlMbps -ge 50) { $p.Notes.Add("Network suitable for RDP, VPN, and remote management") }
    $out['IT Department'] = $p

    # Clamp and assign verdicts
    foreach ($k in $out.Keys) {
        $p = $out[$k]
        $p.Score = [Math]::Max(0,[Math]::Min(100,[int]$p.Score))
        $p.Verdict = if     ($p.Score -ge 85) { 'Ideal' }
                     elseif ($p.Score -ge 70) { 'Recommended' }
                     elseif ($p.Score -ge 55) { 'Marginal' }
                     elseif ($p.Score -ge 40) { 'Limited' }
                     else                     { 'Not Recommended' }
        $p.VColor  = switch ($p.Verdict) {
            'Ideal'           { '#22c55e' }
            'Recommended'     { '#4ade80' }
            'Marginal'        { '#f59e0b' }
            'Limited'         { '#f97316' }
            default           { '#ef4444' }
        }
    }
    return $out
}

# ============================================================
#  HTML REPORT
# ============================================================

function New-Report {
    param([string]$File)

    $ov       = $Script:RESULTS.Overall
    $si       = $Script:RESULTS.SysInfo
    $score    = $ov.Score
    $vText    = $ov.Verdict
    $vColor   = $ov.VColor
    $sColor   = $vColor
    $genTime  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $duration = [Math]::Round(((Get-Date) - $Script:STARTED).TotalMinutes, 1)

    $sysModel   = if ($si) { "$($si.Data.Manufacturer) $($si.Data.Model)" } else { $env:COMPUTERNAME }
    $sysOS      = if ($si) { "$($si.Data.OS) (Build $($si.Data.Build))" } else { 'Unknown' }
    $sysSerial  = if ($si -and $si.Data.Serial -and $si.Data.Serial -ne 'N/A') { $si.Data.Serial } else { 'Unknown' }
    $formFact   = if ($Script:CAPS.IsLaptop) { 'Laptop/Notebook' } else { 'Desktop/Other' }
    $isVMNote   = if ($Script:CAPS.IsVM) { ' (Virtual Machine)' } else { '' }

    # Build module cards
    $cards = ''
    $mods  = @('CPU','RAM','Storage','Battery','GPU','Network','Security','Ports','Software','Thermal','Audio','Input')
    foreach ($m in $mods) {
        $r = $Script:RESULTS[$m]
        if (-not $r) { continue }

        $mc = switch ($r.Status) {
            'PASS'  { '#22c55e' }
            'WARN'  { '#f59e0b' }
            'FAIL'  { '#ef4444' }
            'SKIP'  { '#6b7280' }
            'ERROR' { '#ef4444' }
            default { '#6b7280' }
        }

        $gHtml = ''
        if ($r.Good -and $r.Good.Count -gt 0) {
            $items = ($r.Good | ForEach-Object { "<li>$([System.Net.WebUtility]::HtmlEncode($_))</li>" }) -join ''
            $gHtml = "<ul class='g-list'>$items</ul>"
        }
        $iHtml = ''
        if ($r.Issues -and $r.Issues.Count -gt 0) {
            $items = ($r.Issues | ForEach-Object { "<li>$([System.Net.WebUtility]::HtmlEncode($_))</li>" }) -join ''
            $iHtml = "<ul class='i-list'>$items</ul>"
        }

        $tRows = ''
        if ($r.Data) {
            foreach ($k in $r.Data.Keys) {
                $v = $r.Data[$k]
                if ($null -eq $v) { continue }
                # Serialize complex types to readable strings
                if ($v -is [System.Collections.IEnumerable] -and -not ($v -is [string])) {
                    $parts = foreach ($item in @($v)) {
                        if ($item -is [System.Collections.IDictionary]) {
                            # Ordered dict (drive/GPU entry) → "Key:Value | Key:Value"
                            ($item.Keys | ForEach-Object { "$_`: $($item[$_])" }) -join ' | '
                        } else {
                            [string]$item
                        }
                    }
                    $v = $parts -join '<br>'
                } elseif ($v -is [System.Collections.IDictionary]) {
                    $v = ($v.Keys | ForEach-Object { "$_`: $($v[$_])" }) -join ' | '
                }
                $vStr = [System.Net.WebUtility]::HtmlEncode([string]$v)
                $tRows += "<tr><td class='dk'>$k</td><td>$vStr</td></tr>"
            }
        }
        $tHtml = if ($tRows) { "<table class='dt'>$tRows</table>" } else { '' }

        $cards += @"
<div class="card">
  <div class="ch" onclick="tog(this)">
    <span class="dot" style="background:$mc"></span>
    <span class="ct">$m</span>
    <span class="cs">$([System.Net.WebUtility]::HtmlEncode($r.Summary))</span>
    <span class="sb" style="background:$mc">$($r.Score)</span>
    <span class="cv">&#9660;</span>
  </div>
  <div class="cb">$gHtml$iHtml$tHtml</div>
</div>
"@
    }

    # Score table
    $stRows = ''
    foreach ($m in $ov.Details.Keys) {
        $det = $ov.Details[$m]
        $c = if ($det.Score -ge 80) { '#22c55e' } elseif ($det.Score -ge 50) { '#f59e0b' } else { '#ef4444' }
        $stRows += "<tr><td>$m</td><td style='color:$c;font-weight:700'>$($det.Score)</td><td>$($det.Weight)%</td><td>$($det.Contribution)</td></tr>"
    }

    # Log
    $logHtml = ''
    if ($Script:CFG.Report.IncludeLog) {
        $lines = $Script:LOG | ForEach-Object {
            $cls = if ($_ -match '\[ERROR\]') { 'le' } elseif ($_ -match '\[WARN\]') { 'lw' } else { 'li' }
            $enc = [System.Net.WebUtility]::HtmlEncode($_)
            "<div class='$cls'>$enc</div>"
        }
        $logHtml = "<h2 class='sh'>Diagnostic Log</h2><div class='lc'>$($lines -join '')</div>"
    }

    $barW = "$($score)%"

    $html = '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    $html += "<title>Hardware Eval - $sysModel</title>"
    $html += '<style>'
    $html += ':root{--bg:#0d0d0d;--bg2:#1a1a1a;--bg3:#2b2b2b;--tx:#e5e5e5;--tx2:#888888;--bd:#2b2b2b}'
    $html += '*{box-sizing:border-box;margin:0;padding:0}'
    $html += 'body{background:var(--bg);color:var(--tx);font-family:"Segoe UI",system-ui,sans-serif;font-size:14px}'
    $html += '.wrap{max-width:1100px;margin:0 auto;padding:24px 16px}'
    $html += '.hdr{background:linear-gradient(135deg,#1a1a1a,#0d0d0d);border-radius:12px;padding:28px 32px;margin-bottom:24px;border:1px solid #333333}'
    $html += '.hdr h1{font-size:22px;font-weight:700;color:#ffffff;margin-bottom:4px}'
    $html += '.hdr .sub{color:var(--tx2);font-size:12px}'
    $html += '.mg{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:8px;margin-top:18px}'
    $html += '.mi{background:rgba(255,255,255,.05);border-radius:8px;padding:10px 14px}'
    $html += '.mi .lbl{color:var(--tx2);font-size:11px;text-transform:uppercase;letter-spacing:.5px}'
    $html += '.mi .val{font-weight:600;margin-top:2px;font-size:13px}'
    $html += ".vbox{border-radius:12px;padding:28px;margin-bottom:24px;text-align:center;border:2px solid $sColor;background:rgba(20,20,20,.9)}"
    $html += ".vbox .sc{font-size:72px;font-weight:800;color:$sColor;line-height:1}"
    $html += '.vbox .slbl{font-size:12px;color:var(--tx2);margin-top:4px}'
    $html += ".vbox .vt{font-size:24px;font-weight:700;color:$sColor;margin-top:12px}"
    $html += '.vbox .vsub{color:var(--tx2);margin-top:8px;font-size:13px}'
    $html += '.bar-wrap{background:var(--bg3);border-radius:99px;height:8px;margin:12px auto;max-width:400px}'
    $html += ".bar-fill{height:8px;border-radius:99px;background:$sColor;width:$barW}"
    $html += '.st{width:100%;border-collapse:collapse;margin-bottom:24px}'
    $html += '.st th,.st td{padding:10px 14px;text-align:left;border-bottom:1px solid var(--bd)}'
    $html += '.st th{background:var(--bg3);font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--tx2)}'
    $html += '.st tr:hover td{background:rgba(255,255,255,.03)}'
    $html += '.card{background:var(--bg2);border:1px solid var(--bd);border-radius:10px;margin-bottom:10px;overflow:hidden}'
    $html += '.ch{display:flex;align-items:center;gap:12px;padding:14px 18px;cursor:pointer;user-select:none}'
    $html += '.ch:hover{background:rgba(255,255,255,.04)}'
    $html += '.dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}'
    $html += '.ct{font-weight:700;font-size:15px;min-width:80px}'
    $html += '.cs{color:var(--tx2);flex:1;font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
    $html += '.sb{border-radius:6px;padding:3px 10px;font-weight:700;font-size:13px;color:#fff;min-width:36px;text-align:center}'
    $html += '.cv{color:var(--tx2);font-size:12px;transition:transform .2s}'
    $html += '.cb{padding:0 18px 14px;display:none}'
    $html += '.cb.open{display:block}'
    $html += ".g-list{list-style:none;padding:6px 0}.g-list li::before{content:'+ ';color:#22c55e}.g-list li{margin:2px 0}"
    $html += ".i-list{list-style:none;padding:6px 0}.i-list li::before{content:'! ';color:#f59e0b;font-weight:700}.i-list li{margin:2px 0;color:#fde68a}"
    $html += '.dt{width:100%;border-collapse:collapse;margin-top:8px;font-size:12px}'
    $html += '.dt td{padding:5px 10px;border-bottom:1px solid rgba(255,255,255,.06)}'
    $html += '.dt td.dk{color:var(--tx2);width:38%}'
    $html += '.lc{background:#0a0a0a;border-radius:8px;padding:14px;font-family:Consolas,monospace;font-size:11px;max-height:400px;overflow-y:auto}'
    $html += '.li{color:#666666}.lw{color:#f59e0b}.le{color:#ef4444}'
    $html += '.sh{font-size:12px;text-transform:uppercase;letter-spacing:1px;color:var(--tx2);margin:24px 0 10px;font-weight:600;border-bottom:1px solid var(--bd);padding-bottom:6px}'
    $html += '.ft{text-align:center;color:var(--tx2);font-size:12px;margin-top:32px;padding:16px;border-top:1px solid var(--bd)}'
    $html += '.dg{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:12px;margin-bottom:24px}'
    $html += '.dc{background:var(--bg2);border:1px solid var(--bd);border-radius:10px;padding:16px}'
    $html += '.dc-top{display:flex;justify-content:space-between;align-items:center;margin-bottom:3px}'
    $html += '.dc-name{font-weight:700;font-size:14px}'
    $html += '.dc-badge{border-radius:4px;padding:2px 9px;font-size:11px;font-weight:700;color:#000;letter-spacing:.3px}'
    $html += '.dc-desc{color:var(--tx2);font-size:11px;margin-bottom:10px}'
    $html += '.dc-bw{background:var(--bg3);border-radius:99px;height:6px;margin-bottom:5px}'
    $html += '.dc-bar{height:6px;border-radius:99px}'
    $html += '.dc-sc{font-size:22px;font-weight:800;margin-bottom:8px}'
    $html += '.dn{list-style:none;font-size:11px;padding:0}'
    $html += '.dn li{margin:3px 0;padding-left:14px;position:relative;color:var(--tx2)}'
    $html += '.dn li.ok::before{content:"+";position:absolute;left:0;color:#22c55e;font-weight:700}'
    $html += '.dn li.ng::before{content:"!";position:absolute;left:0;color:#f59e0b;font-weight:700}'
    $html += '.dn li.bad::before{content:"x";position:absolute;left:0;color:#ef4444;font-weight:700}'
    $html += '.dc-viab{border-color:var(--bd)}'
    $html += '.disc{background:#1a1a1a;border:1px solid #2b2b2b;border-left:4px solid #f59e0b;border-radius:8px;padding:14px 18px;margin-bottom:24px;font-size:12px;color:var(--tx2)}'
    $html += '.disc strong{color:#fde68a}'
    $html += '@media(max-width:640px){.mg{grid-template-columns:1fr 1fr}.vbox .sc{font-size:56px}.dg{grid-template-columns:1fr}}'
    $html += '@media print{.cb{display:block!important}}'
    $html += '</style></head><body><div class="wrap">'

    # Pull extra SysInfo fields
    $sysBIOS     = if ($si -and $si.Data.BIOS)         { $si.Data.BIOS }         else { 'N/A' }
    $sysBuild    = if ($si -and $si.Data.Build)         { $si.Data.Build }        else { 'N/A' }
    $sysArch     = if ($si -and $si.Data.Architecture) { $si.Data.Architecture } else { 'N/A' }
    $sysDomain   = if ($si -and $si.Data.Domain)       { $si.Data.Domain }       else { 'N/A' }
    $sysLastBoot = if ($si -and $si.Data.LastBoot)     { $si.Data.LastBoot }     else { 'N/A' }
    $sysUptime   = if ($si -and $si.Data.Uptime)       { $si.Data.Uptime }       else { 'N/A' }
    $sysUser     = if ($si -and $si.Data.User)         { $si.Data.User }         else { "$env:USERDOMAIN\$env:USERNAME" }

    # Header
    $html += '<div class="hdr"><h1>Hardware Evaluation Report</h1>'
    $html += "<div class='sub'>$($Script:CFG.Report.CompanyName) &middot; Session $Script:SESSID &middot; $genTime</div>"
    $html += '<div class="mg">'
    $html += "<div class='mi'><div class='lbl'>Manufacturer</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($(if ($si) { $si.Data.Manufacturer } else { 'N/A' })))</div></div>"
    $html += "<div class='mi'><div class='lbl'>Model</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($(if ($si) { $si.Data.Model } else { 'N/A' })))</div></div>"
    $html += "<div class='mi'><div class='lbl'>Serial</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($sysSerial))</div></div>"
    $html += "<div class='mi'><div class='lbl'>BIOS</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($sysBIOS))</div></div>"
    $html += "<div class='mi'><div class='lbl'>OS</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($(if ($si) { $si.Data.OS } else { 'N/A' })))</div></div>"
    $html += "<div class='mi'><div class='lbl'>Build</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($sysBuild))</div></div>"
    $html += "<div class='mi'><div class='lbl'>Architecture</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($sysArch))</div></div>"
    $html += "<div class='mi'><div class='lbl'>Hostname</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($env:COMPUTERNAME))</div></div>"
    $html += "<div class='mi'><div class='lbl'>Domain</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($sysDomain))</div></div>"
    $html += "<div class='mi'><div class='lbl'>Last Boot</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($sysLastBoot))</div></div>"
    $html += "<div class='mi'><div class='lbl'>Uptime</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($sysUptime))</div></div>"
    $html += "<div class='mi'><div class='lbl'>User</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($sysUser))</div></div>"
    $html += "<div class='mi'><div class='lbl'>Form Factor</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode("$formFact$isVMNote"))</div></div>"
    $techDisplay  = if ($Script:TechInfo) { $Script:TechInfo.Technician }   else { "$env:USERDOMAIN\$env:USERNAME" }
    $damageStatus = if ($Script:TechInfo) { $Script:TechInfo.DamageStatus } else { 'Not assessed' }
    $html += "<div class='mi'><div class='lbl'>Evaluated By</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($techDisplay))</div></div>"
    $html += "<div class='mi'><div class='lbl'>Physical Damage</div><div class='val'>$([System.Net.WebUtility]::HtmlEncode($damageStatus))</div></div>"
    $html += "<div class='mi'><div class='lbl'>Duration</div><div class='val'>$duration min</div></div>"
    $html += '</div></div>'

    # System Capabilities — quick visual summary of what hardware is present.
    # Helps explain why some scores are high/low (e.g. no Wi-Fi on a server).
    $caps = $Script:CAPS
    if ($caps) {
        $chip = {
            param($label, $present, $isInfo = $false)
            $bg = if ($isInfo) { '#1e293b' } elseif ($present) { '#14532d' } else { '#3f1d1d' }
            $fg = if ($isInfo) { '#cbd5e1' } elseif ($present) { '#86efac' } else { '#94a3b8' }
            $icon = if ($isInfo) { '' } elseif ($present) { '+ ' } else { 'x ' }
            "<span style='display:inline-block;background:$bg;color:$fg;border-radius:6px;padding:5px 12px;margin:3px;font-size:12px;font-weight:600'>$icon$label</span>"
        }
        $html += '<h2 class="sh">System Capabilities</h2>'
        $html += '<div style="margin-bottom:24px">'
        $formLabel = if ($caps.IsLaptop) { 'Form: Laptop / Notebook' } else { 'Form: Desktop / Server' }
        $html += (& $chip $formLabel $true $true)
        if ($caps.IsVM) { $html += (& $chip 'Virtual Machine' $true $true) }
        $html += (& $chip 'Battery'      $caps.HasBattery)
        $html += (& $chip 'Wi-Fi'        $caps.HasWiFi)
        $html += (& $chip 'Ethernet'     $caps.HasEthernet)
        $html += (& $chip 'Discrete GPU' $caps.HasDiscreteGPU)
        $html += (& $chip 'Bluetooth'    $caps.HasBluetooth)
        $html += '</div>'
        $html += "<div style='font-size:11px;color:var(--tx2);margin:-16px 0 24px'>Capability detection runs before tests and explains scoping decisions (e.g. battery test is skipped on machines without a battery, missing Wi-Fi is not penalized on servers).</div>"
    }

    # Damage details block - prominent if damage was reported
    if ($Script:TechInfo -and $Script:TechInfo.DamageDetails -and $Script:TechInfo.DamageDetails -notin @('None','N/A')) {
        $html += "<div class='disc'><strong>Reported physical damage:</strong> $([System.Net.WebUtility]::HtmlEncode($Script:TechInfo.DamageDetails))</div>"
    }

    # Deployment Blockers - issues the technician explicitly reported via Y/N prompts.
    # These are user-confirmed defects that block handover regardless of the auto score.
    $blockers = [System.Collections.Generic.List[string]]::new()

    if ($Script:TechInfo -and $Script:TechInfo.DamageStatus -eq 'Damage reported') {
        $dmg = if ($Script:TechInfo.DamageDetails -and $Script:TechInfo.DamageDetails -notin @('None','N/A')) { $Script:TechInfo.DamageDetails } else { 'unspecified' }
        $blockers.Add("Physical damage reported: $dmg")
    }
    $gpuData = if ($Script:RESULTS.GPU) { $Script:RESULTS.GPU.Data } else { $null }
    if ($gpuData) {
        if ("$($gpuData.Dead_Pixels)" -match 'Reported') {
            $blockers.Add("Dead/stuck pixels confirmed during color test: $($gpuData.Dead_Pixels)")
        }
        if ("$($gpuData.Screen_Quality)" -match '^Issue') {
            $blockers.Add("Screen color/brightness issue: $($gpuData.Screen_Quality)")
        }
    }
    $inpData = if ($Script:RESULTS.Input) { $Script:RESULTS.Input.Data } else { $null }
    if ($inpData) {
        if ($inpData.Keyboard_Broken_Keys) {
            $blockers.Add("Keyboard keys not working: $($inpData.Keyboard_Broken_Keys)")
        } elseif ("$($inpData.Keyboard_Check)" -match '^FAIL') {
            $blockers.Add("Keyboard test failed: $($inpData.Keyboard_Check)")
        }
        if ("$($inpData.Touch_Screen_Check)" -match '^FAIL') {
            $blockers.Add("Touch screen not responding: $($inpData.Touch_Screen_Check)")
        }
    }

    if ($blockers.Count -gt 0) {
        $html += "<div class='disc' style='border-left-color:#ef4444;background:#1f0a0a'>"
        $html += "<strong style='color:#fca5a5;font-size:14px'>DEPLOYMENT BLOCKERS &mdash; do not assign this laptop to a new employee until these are addressed</strong>"
        $confirmer = if ($Script:TechInfo -and $Script:TechInfo.Technician) { $Script:TechInfo.Technician } else { 'Evaluator' }
        $html += "<div style='margin-top:6px;color:var(--tx2);font-size:11px'>$([System.Net.WebUtility]::HtmlEncode($confirmer)) confirmed the following defects during this evaluation:</div>"
        $html += "<ul style='margin:10px 0 4px;padding-left:20px;color:#fde2e2'>"
        foreach ($b in $blockers) { $html += "<li style='margin:4px 0'>$([System.Net.WebUtility]::HtmlEncode($b))</li>" }
        $html += "</ul></div>"
    }

    # Verdict
    $html += '<div class="vbox">'
    $html += "<div class='sc'>$score</div><div class='slbl'>Overall Score (out of 100)</div>"
    $html += "<div class='bar-wrap'><div class='bar-fill'></div></div>"
    $html += "<div class='vt'>$vText</div>"
    $html += "<div class='vsub'>Errors: $Script:ERR_COUNT &nbsp;&middot;&nbsp; Warnings: $Script:WARN_COUNT &nbsp;&middot;&nbsp; Critical Fails: $($ov.Crits.Count)</div>"
    $html += '</div>'

    # Technician's expert verdict - shown right under the auto-verdict
    if ($Script:TechInfo -and $Script:TechInfo.FinalVerdict -and $Script:TechInfo.FinalVerdict -notmatch '^(N/A|No technician verdict)') {
        $tv      = $Script:TechInfo.FinalVerdict
        $tvNote  = $Script:TechInfo.FinalVerdictNote
        $tvColor = if ($tv -match 'NOT')        { '#ef4444' }
                   elseif ($tv -match 'Maybe')  { '#f59e0b' }
                   else                         { '#22c55e' }
        $html += "<div class='disc' style='border-left-color:$tvColor'>"
        $html += "<strong>$([System.Net.WebUtility]::HtmlEncode($Script:TechInfo.Technician))'s Verdict</strong>: "
        $html += "<span style='color:$tvColor;font-weight:700'>$([System.Net.WebUtility]::HtmlEncode($tv))</span>"
        if ($tvNote) { $html += "<br><span style='opacity:.85'>&ldquo;$([System.Net.WebUtility]::HtmlEncode($tvNote))&rdquo;</span>" }
        $html += '</div>'
    }

    # VM note (only shown when running inside a hypervisor)
    if ($Script:CAPS.IsVM) {
        $html += "<div class='disc'>This system is a <strong>virtual machine</strong> - hardware scores reflect the VM configuration, not the underlying physical host.</div>"
    }

    # Long-term Concerns - every issue from every module, grouped by severity.
    # Surfaces ALL issues prominently so they aren't hidden in collapsed module details,
    # regardless of overall verdict or technician's call.
    $critIssues = [System.Collections.Generic.List[object]]::new()
    $warnIssues = [System.Collections.Generic.List[object]]::new()
    foreach ($mk in $Script:RESULTS.Keys) {
        $r = $Script:RESULTS[$mk]
        if (-not $r -or -not $r.Issues -or $r.Issues.Count -eq 0) { continue }
        if ($mk -eq 'Overall') { continue }
        foreach ($iss in $r.Issues) {
            $entry = [PSCustomObject]@{ Module = $r.Module; Issue = $iss }
            if ($r.Status -eq 'FAIL' -or $r.Status -eq 'ERROR') { $critIssues.Add($entry) }
            else                                                 { $warnIssues.Add($entry) }
        }
    }

    if ($critIssues.Count -gt 0 -or $warnIssues.Count -gt 0) {
        $html += '<h2 class="sh">Long-term Concerns</h2>'
        $html += "<div class='disc'><strong>These issues will affect this device in the long run - review before deployment regardless of the overall verdict.</strong></div>"
        if ($critIssues.Count -gt 0) {
            $html += "<div class='disc' style='border-left-color:#ef4444'><strong style='color:#fca5a5'>Critical (must address):</strong><ul style='margin:8px 0 0;padding-left:20px'>"
            foreach ($e in $critIssues) {
                $html += "<li><strong>[$([System.Net.WebUtility]::HtmlEncode($e.Module))]</strong> $([System.Net.WebUtility]::HtmlEncode($e.Issue))</li>"
            }
            $html += '</ul></div>'
        }
        if ($warnIssues.Count -gt 0) {
            $html += "<div class='disc' style='border-left-color:#f59e0b'><strong style='color:#fde68a'>Warnings (monitor / plan to address):</strong><ul style='margin:8px 0 0;padding-left:20px'>"
            foreach ($e in $warnIssues) {
                $html += "<li><strong>[$([System.Net.WebUtility]::HtmlEncode($e.Module))]</strong> $([System.Net.WebUtility]::HtmlEncode($e.Issue))</li>"
            }
            $html += '</ul></div>'
        }
    }

    # Score table
    $html += '<h2 class="sh">Score Breakdown</h2>'
    $html += "<table class='st'><thead><tr><th>Module</th><th>Score</th><th>Weight</th><th>Contribution</th></tr></thead><tbody>$stRows</tbody></table>"

    # Vetos (catastrophic single-component issues that cap the overall verdict)
    if ($ov.Vetos -and $ov.Vetos.Count -gt 0) {
        $html += '<h2 class="sh">Critical Issues (veto)</h2>'
        $html += "<div class='disc'><strong>The overall verdict is capped because one or more components fail catastrophically:</strong><ul>"
        foreach ($vt in $ov.Vetos) { $html += "<li>$([System.Net.WebUtility]::HtmlEncode($vt))</li>" }
        $html += '</ul></div>'
    }

    # Fixable Issues - software/config problems the technician can resolve in-place,
    # then re-run the script to verify. Aggregated from every module's Fixables list.
    $allFixables = [System.Collections.Generic.List[object]]::new()
    foreach ($mk in $Script:RESULTS.Keys) {
        $r = $Script:RESULTS[$mk]
        if ($r -and $r.Fixables) {
            foreach ($f in $r.Fixables) {
                $entry = [ordered]@{ Module = $r.Module }
                foreach ($k in $f.Keys) { $entry[$k] = $f[$k] }
                $allFixables.Add($entry)
            }
        }
    }
    if ($allFixables.Count -gt 0) {
        $html += '<h2 class="sh">Fixable Issues - Run, Then Re-Test</h2>'
        $resolver = if ($Script:TechInfo -and $Script:TechInfo.Technician) { $Script:TechInfo.Technician } else { 'You' }
        $html += "<div class='disc'>These are software/config problems $([System.Net.WebUtility]::HtmlEncode($resolver)) can resolve before re-running this script. <strong>Fix - reboot if required - re-run the script</strong> to verify.</div>"
        $html += "<table class='st'><thead><tr><th>Module</th><th>What to fix</th><th>Why</th><th>Run this (admin)</th><th>Time</th><th>Reboot?</th></tr></thead><tbody>"
        foreach ($f in $allFixables) {
            $rb = if ($f.RequiresReboot) { 'Yes' } else { 'No' }
            $html += "<tr>"
            $html += "<td>$([System.Net.WebUtility]::HtmlEncode("$($f.Module)"))</td>"
            $html += "<td><strong>$([System.Net.WebUtility]::HtmlEncode("$($f.Title)"))</strong></td>"
            $html += "<td>$([System.Net.WebUtility]::HtmlEncode("$($f.Reason)"))</td>"
            $html += "<td><code style='background:#0a0a0a;padding:2px 6px;border-radius:4px;color:#a3e635;font-size:11px'>$([System.Net.WebUtility]::HtmlEncode("$($f.FixCommand)"))</code></td>"
            $html += "<td>$([System.Net.WebUtility]::HtmlEncode("$($f.EstimatedTime)"))</td>"
            $html += "<td>$rb</td>"
            $html += "</tr>"
        }
        $html += '</tbody></table>'
    }

    # BSOD details - show every crash with bugcheck code so the technician knows what failed
    $swRes = $Script:RESULTS.Software
    if ($swRes -and $swRes.Data -and $swRes.Data.BSOD_Last30Days -gt 0 -and $swRes.Data.BSOD_Details) {
        $html += '<h2 class="sh">Recent System Crashes (BSOD)</h2>'
        $html += "<div class='disc'>The system crashed <strong>$($swRes.Data.BSOD_Last30Days) times in the last 30 days</strong>. Driver-origin crashes can usually be fixed by reinstalling the offending driver. Hardware-origin crashes (WHEA, MCE, MEMORY_MANAGEMENT) point to a real component failure - run hardware diagnostics (RAM test, drive SMART, CPU stress).</div>"
        $html += "<table class='st'><thead><tr><th>When</th><th>Bugcheck</th><th>Name</th><th>Likely origin</th></tr></thead><tbody>"
        foreach ($entry in ($swRes.Data.BSOD_Details -split ' \|\| ')) {
            if ($entry -match '^(.+?)\s+(0x[0-9A-F]+|unparsed)\s+(\S+)\s+\[(.+)\]$') {
                $tm   = $Matches[1]; $cd = $Matches[2]; $nm = $Matches[3]; $or = $Matches[4]
                $rowColor = if ($or -match 'Hardware') { '#ef4444' } elseif ($or -match 'Driver|Software') { '#f59e0b' } else { 'inherit' }
                $html += "<tr><td>$([System.Net.WebUtility]::HtmlEncode($tm))</td><td><code>$([System.Net.WebUtility]::HtmlEncode($cd))</code></td><td>$([System.Net.WebUtility]::HtmlEncode($nm))</td><td style='color:$rowColor'>$([System.Net.WebUtility]::HtmlEncode($or))</td></tr>"
            }
        }
        $html += '</tbody></table>'
    }

    # Department / Use-Case Suitability
    $dept = Get-DepartmentScores
    $html += '<h2 class="sh">Use-Case Suitability</h2>'
    $html += '<div class="dg">'
    foreach ($dk in $dept.Keys) {
        $dp   = $dept[$dk]
        $bPct = "$($dp.Score)%"
        $nHtml = ''
        if ($dp.Notes.Count -gt 0) {
            $nHtml = '<ul class="dn">'
            foreach ($n in $dp.Notes) {
                $cls = if ($n -match 'comfortable|good|suitable|fast|smooth|NVMe|TPM 2\.0|BitLocker active|Dedicated GPU  -  GPU|Dedicated GPU  -  hardware|handles|confirmed|adequate|meets|sufficient') { 'ok' } else { 'ng' }
                $nHtml += "<li class='$cls'>$([System.Net.WebUtility]::HtmlEncode($n))</li>"
            }
            $nHtml += '</ul>'
        }
        $html += "<div class='dc'>"
        $html += "<div class='dc-top'><div class='dc-name'>$([System.Net.WebUtility]::HtmlEncode($dp.Name))</div>"
        $html += "<span class='dc-badge' style='background:$($dp.VColor)'>$($dp.Verdict)</span></div>"
        $html += "<div class='dc-bw'><div class='dc-bar' style='width:$bPct;background:$($dp.VColor)'></div></div>"
        $html += "<div class='dc-sc' style='color:$($dp.VColor)'>$($dp.Score)<span style='font-size:13px;font-weight:400;color:var(--tx2)'>/100</span></div>"
        $html += $nHtml
        $html += '</div>'
    }
    $html += '</div>'

    # Module cards
    $html += '<h2 class="sh">Detailed Test Results</h2>'
    $html += $cards

    # Log
    $html += $logHtml

    # Footer
    $html += "<div class='ft'>Hardware Evaluation &middot; $genTime &middot; Session $Script:SESSID</div>"
    $html += '</div>'
    $html += '<script>'
    $html += 'function tog(h){var b=h.nextElementSibling,c=h.querySelector(".cv");if(b.classList.contains("open")){b.classList.remove("open");c.style.transform=""}else{b.classList.add("open");c.style.transform="rotate(180deg)"}}'
    $html += 'document.querySelectorAll(".card").forEach(function(c){var d=c.querySelector(".dot");var bg=d&&d.style.background;if(bg==="rgb(239, 68, 68)"||bg==="#ef4444"||bg==="rgb(245, 158, 11)"||bg==="#f59e0b"){tog(c.querySelector(".ch"))}})'
    $html += '</script></body></html>'

    $html | Out-File -FilePath $File -Encoding UTF8
    Write-Log "Report saved: $File" 'OK'
    return $File
}

# ============================================================
#  INTERACTIVE MENU
# ============================================================

function Show-Menu {
    # Always run all modules - we want a full report every time, no menu.
    # The TestModules parameter still works for power users who explicitly pass a subset.
    $all = @('CPU','RAM','Storage','Battery','GPU','Network','Security','Ports','Software','Thermal','Audio','Input')
    $sel = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ($TestModules) {
        $TestModules.Split(',') | ForEach-Object { [void]$sel.Add($_.Trim()) }
    } else {
        $all | ForEach-Object { [void]$sel.Add($_) }
    }
    Write-Host ''
    Write-Host '  Running full evaluation (all modules) ...' -ForegroundColor Cyan
    Write-Host ''
    return $sel
}

# ============================================================
#  MAIN
# ============================================================

function Main {
    $Script:FullAuto = $FullAuto.IsPresent
    Write-Banner

    # Auto-elevate / offer elevation
    if (-not $Script:IS_ADMIN) {
        Write-Host '  [!] Running without Administrator privileges.' -ForegroundColor Yellow
        Write-Host '      SMART data, WinSAT, thermal sensors, battery report may be limited.' -ForegroundColor Yellow

        $doElevate = $false
        if ($FullAuto) {
            # Unattended mode: silently re-launch elevated without prompting
            $doElevate = $true
        } else {
            Write-Host '  Elevate now? [Y/N] ' -ForegroundColor Yellow -NoNewline
            $doElevate = ((Read-Host).Trim().ToUpper() -eq 'Y')
        }

        if ($doElevate) {
            $args2 = $MyInvocation.BoundParameters.GetEnumerator() | ForEach-Object {
                if ($_.Value -is [switch]) {
                    # Preserve explicit $false — "-Key:$false" disables the switch
                    if ($_.Value.IsPresent) { "-$($_.Key)" } else { "-$($_.Key):`$false" }
                } else {
                    # Escape embedded quotes in string/path values
                    "-$($_.Key) `"$($_.Value -replace '"', '\"')`""
                }
            }
            $argStr = $args2 -join ' '

            if ($PSCommandPath) {
                # Normal file execution: re-launch the same file elevated
                Start-Process powershell.exe -Verb RunAs `
                    -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $argStr"
            } else {
                # irm URL | iex (piped) mode: $PSCommandPath is empty.
                # Re-download from the canonical URL in the elevated session — avoids
                # the scriptblock capture problem where MyCommand.ScriptBlock is $null in iex context.
                $rawUrl  = 'https://raw.githubusercontent.com/maxdorx/hardware-eval/main/Hardware-Evaluation.ps1'
                $elevCmd = "& ([scriptblock]::Create((irm '$rawUrl'))) $argStr"
                try {
                    Start-Process powershell.exe -Verb RunAs `
                        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"$elevCmd`""
                } catch {
                    Write-Host "  [!] Could not relaunch elevated: $_" -ForegroundColor Red
                    Write-Host '      Please re-run PowerShell as Administrator manually.' -ForegroundColor Yellow
                }
            }
            exit 0
        }
        Write-Host ''
    }

    Initialize-Env
    Initialize-Tools
    Get-SysInfo

    $toRun = Show-Menu

    # Collect technician info and visual damage assessment before tests start
    Get-TechnicianInfo

    Write-Host ''
    Write-Host "  Modules: $($toRun -join ', ')" -ForegroundColor Cyan
    Write-Host ''

    $dispatch = [ordered]@{
        CPU      = { Test-CPU      }
        RAM      = { Test-RAM      }
        Storage  = { Test-Storage  }
        Battery  = { Test-Battery  }
        GPU      = { Test-GPU      }
        Network  = { Test-Network  }
        Security = { Test-Security }
        Ports    = { Test-Ports    }
        Software = { Test-Software }
        Thermal  = { Test-Thermal  }
        Audio    = { Test-Audio    }
        Input    = { Test-Input    }
    }

    foreach ($mod in $dispatch.Keys) {
        if (-not $toRun.Contains($mod)) { continue }
        Write-Log "--- $mod ---" 'HEAD'
        try { & $dispatch[$mod] } catch { Write-Log "Unhandled error in $mod : $_" 'ERROR' }
    }

    Invoke-Scoring

    # Now that the score is computed, ask the technician for their expert recommendation
    Get-TechnicianVerdict

    if (-not $NoReport) {
        $rd = $Script:CFG.Report.OutputDir
        if (-not (Test-Path $rd)) { New-Item -ItemType Directory $rd -Force | Out-Null }

        # Use chassis serial as the stable hardware identifier (survives OS reinstall / rename).
        # Falls back to primary NIC MAC, then hostname if serial is missing or generic.
        $hwId = ''
        $rawSerial = if ($Script:RESULTS.SysInfo) { "$($Script:RESULTS.SysInfo.Data.Serial)" } else { '' }
        $rawSerial = $rawSerial -replace '[^\w]', ''   # strip spaces, dashes, etc.
        if ($rawSerial -and $rawSerial.Length -ge 4 -and $rawSerial -notmatch '^(NA|ToBeFilledByOEM|Default|None|0+)$') {
            $hwId = $rawSerial
        } else {
            # MAC address of the first Up adapter (remove dashes)
            $mac = (Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
                    Where-Object { $_.Status -eq 'Up' } |
                    Sort-Object -Property Speed -Descending |
                    Select-Object -First 1).MacAddress -replace '-',''
            $hwId = if ($mac) { $mac } else { $env:COMPUTERNAME }
        }
        $fname = "HWEval_${hwId}_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        $fpath = Join-Path $rd $fname
        $rFile = New-Report -File $fpath

        Write-Host ''
        Write-Host "  Report: $rFile" -ForegroundColor Green

        if ($Script:CFG.Report.OpenAfter) { Start-Process $rFile }
    }

    # Remove only the DiskSpd test file; keep the tools dir so binaries are cached across runs
    try { Remove-Item (Join-Path $Script:CFG.Tools.WorkDir 'dstest.dat') -Force -ErrorAction SilentlyContinue } catch {}

    $mins = [Math]::Round(((Get-Date) - $Script:STARTED).TotalMinutes, 1)
    Write-Host ''
    Write-Host "  Done in $mins min  |  Score: $($Script:RESULTS.Overall.Score)/100  |  $($Script:RESULTS.Overall.Verdict)" -ForegroundColor Cyan
    Write-Host ''
}

Main
