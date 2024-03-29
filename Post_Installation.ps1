# START POST INSTALLATION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

#===============================================================================
# !0. Utilities
#===============================================================================
function Format-Color {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline, ParameterSetName = "Formatter")]
        [Alias("m")]
        $Message,
 
        [Parameter(ParameterSetName = "Formatter")]
        [Alias("f")]
        [ArgumentCompleter({ param ( 
                    $commandName,
                    $parameterName,
                    $wordToComplete,
                    $commandAst,
                    $fakeBoundParameters )
                @(
                    "Black"        
                    "DarkRed"     
                    "DarkGreen"   
                    "DarkYellow"  
                    "DarkBlue"    
                    "DarkMagenta" 
                    "DarkCyan"    
                    "Gray"        
                    "DarkGray"    
                    "Red"         
                    "Green"       
                    "Yello"       
                    "Blue"        
                    "Magenta"     
                    "Cyan"        
                    "White"
                ) | ? { $_ -like "$wordToComplete*" }  | % { $_ } })]
        $ForegroundColor,
 
        [Parameter(ParameterSetName = "Formatter")]
        [Alias("b")]
        [ArgumentCompleter({
                param ( 
                    $commandName,
                    $parameterName,
                    $wordToComplete,
                    $commandAst,
                    $fakeBoundParameters )
                @(
                    "Black"        
                    "DarkRed"     
                    "DarkGreen"   
                    "DarkYellow"  
                    "DarkBlue"    
                    "DarkMagenta" 
                    "DarkCyan"    
                    "Gray"        
                    "DarkGray"    
                    "Red"         
                    "Green"       
                    "Yello"       
                    "Blue"        
                    "Magenta"     
                    "Cyan"        
                    "White"       
                     
                ) | ? { $_ -like "$wordToComplete*" }  | % { $_ } 
            
            })]
        $BackgroundColor,
 
        [Parameter(ParameterSetName = "ShowColorTable")]
        [switch]
        $ShowColorTable
    )
 
    begin {
        #Refer: https://duffney.io/usingansiescapesequencespowershell/
        function Get-ColorTable {
            $esc = $([char]27)
            Write-Host "`n$esc[1;4m256-Color Foreground & Background Charts$esc[0m"
            foreach ($fgbg in 38, 48) {
                # foreground/background switch
                foreach ($color in 0..255) {
                    # color range
                    #Display the colors
                    $field = "$color".PadLeft(4)  # pad the chart boxes with spaces
                    Write-Host -NoNewLine "$esc[$fgbg;5;${color}m$field $esc[0m"
                    #Display 8 colors per line
                    if ( (($color + 1) % 8) -eq 0 ) { Write-Host }
                }
                Write-Host
            }
        }
        function ParseColor {
            param (
                $Value
            )
            if ($Value) {
                if ($Value -is [string]) {
                    if ($colorMapping.ContainsKey($Value)) {
                        Write-Output $colorMapping[$Value]
                    }
                }
                elseif ($Value -is [int]) {
                    # Write-Output $Valuess
                    if (($Value -le 255 -and $Value -ge 0)) {
                        Write-Output $Value
                    }
                    else {
                        throw "The color value should beteen 0 and 255, but the input value is $Value. Please check and retry again."
                    }
                }
            }
        }
 
        #Refer: https://ss64.com/nt/syntax-ansi.html#:~:text=How-to%3A%20Use%20ANSI%20colors%20in%20the%20terminal%20,%20%206%20%2018%20more%20rows%20
        $colorMapping = @{
            "Black"       = 30 
            "DarkRed"     = 31
            "DarkGreen"   = 32
            "DarkYellow"  = 33
            "DarkBlue"    = 34
            "DarkMagenta" = 35
            "DarkCyan"    = 36
            "Gray"        = 37
            "DarkGray"    = 90
            "Red"         = 91
            "Green"       = 92
            "Yello"       = 93
            "Blue"        = 94
            "Magenta"     = 95
            "Cyan"        = 96
            "White"       = 97
        }
 
        $esc = $([char]27)
        $backgroudSwitch = 48
        $foregroundSwitch = 38
 
        $ansiParam = @()
        if ($null -ne $PSBoundParameters["ForegroundColor"]) {
            if ($ForegroundColor -is [string]) {
                $ansiParam += "5;$(ParseColor $ForegroundColor)".Trim()
            }
            else {
                $ansiParam += "$foregroundSwitch;5;$(ParseColor $ForegroundColor)".Trim()
            }
        }
        if ($null -ne $PSBoundParameters["BackgroundColor"]) {
            if ($BackgroundColor -is [string]) {
                $ansiParam += "5;$((ParseColor $BackgroundColor)+10)".Trim()
            }
            else {
                $ansiParam += "$backgroudSwitch;5;$(ParseColor $BackgroundColor)".Trim()
            }
        }
    }
 
    process {
        $current = $_
        if ($PSCmdlet.ParameterSetName -eq "ShowColorTable") {
            Get-ColorTable
        }
        else {
            
            if ([string]::IsNullOrEmpty($current)) {
                $current = $Message
            }
            if ($ansiParam.Count -gt 0) {
                Write-Output "$esc[$($ansiParam -join ";")m$current$($esc)[0m"
            }
            else {
                Write-Output $current
            }
        }
    }
}

function Set-Env {
    param ([String]$Target, [String]$FullPath)
    [System.Environment]::SetEnvironmentVariable("Path", "$($Evn:PATH);$($FullPath)", $Target);
    $CurrentPath = [System.Environment]::GetEnvironmentVariable('Path', $Target);
    $SplittedPath = $CurrentPath -split ';';
    $CleanedPath = $SplittedPath | Sort-Object -Unique;
    $NewPath = $CleanedPath -join ';';
    [System.Environment]::SetEnvironmentVariable('Path', $NewPath, $Target);
    Format-Color -Message "$($FullPath) added to System Path." -ForegroundColor Green;
}

#===============================================================================
# !1. Activating Windows Operating System
#===============================================================================
function Initiate-WAS{

    $FilePath = "$env:TEMP\MAS.cmd"
    $ScriptArgs = "$args "

    $DownloadURL = 'https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/master/MAS/All-In-One-Version/MAS_AIO.cmd'
    $DownloadURL2 = 'https://gitlab.com/massgrave/microsoft-activation-scripts/-/raw/master/MAS/All-In-One-Version/MAS_AIO.cmd'
    $DownloadURLSpeedUP = 'https://ghdl.feizhuqwq.cf/https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/master/MAS/All-In-One-Version/MAS_AIO.cmd'

    try {
        Invoke-WebRequest -Uri $DownloadURLSpeedUP -UseBasicParsing -OutFile $FilePath -ErrorAction Stop
    }
    catch {
        try {
            Invoke-WebRequest -Uri $DownloadURL2 -UseBasicParsing -OutFile $FilePath -ErrorAction Stop
        }
        catch {
            try {
                Invoke-WebRequest -Uri $DownloadURL -UseBasicParsing -OutFile $FilePath -ErrorAction Stop
            }
            catch {
                Write-Error $_
                Return
            }
        }
    }

    if (Test-Path $FilePath) {
        Start-Process $FilePath $ScriptArgs -Wait
        $item = Get-Item -LiteralPath $FilePath
        $item.Delete()
    }
}

function Activate-Microsoft {
    param (
        [string]$ActivateFileURL,
        [string]$ActivateFilePath,
        [string]$ActivateFileDestination,
        [string]$SystemName
    )

    $config = Get-Content $configFile -Raw | ConvertFrom-Json
    $ActivateFileURL = $config.skus.download
    $ActivateFilePath = $config.skus.savePath
    $ActivateFileDestination = $config.skus.destination

    $SystemName = Get-WmiObject Win32_OperatingSystem | Select -Property Caption
    if ($SystemName.Caption.ToString().contains("LTSC")) {
        Invoke-WebRequest -Uri $ActivateFileURL -UseBasicParsing -OutFile $ActivateFilePath -ErrorAction Stop
        Compress-Archive -LiteralPath $ActivateFilePath -DestinationPath  $ActivateFileDestination
        Initiate-WAS
    }else { 
        Initiate-WAS
    }
}

#===============================================================================
# !2. Installing package and installer manager
#===============================================================================
function Install-Scoop {
# USE: Install-Scoop -AppPath $ScoopPath -App $ScoopApp -PS1File $ScoopPS1
    param (
        [string]$AppPath,
        [string]$App,
        [string]$PS1File
    )

    # If the Scoop file does not exist, go install it
    if (-not(Test-Path -Path $App -PathType Leaf)) {
        try {
            Format-Color -Message "Installing Scoop installer manager..." -ForegroundColor Blue;
            iex "& {$(irm get.scoop.sh)} -RunAsAdmin";
            Set-Env -Target 'Machine' -FullPath $AppPath;
            Unblock-File -Path $PS1File
            Unblock-File -Path $APP;

            # !important!
            Format-Color -Message "Important! Optimize Scoop Environment" -ForegroundColor Blue;
            scoop install sudo
            sudo Add-MpPreference -ExclusionPath 'C:\Users\Administrator\scoop';
            sudo Add-MpPreference -ExclusionPath 'C:\ProgramData\scoop';
            sudo Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' -Name 'LongPathsEnabled' -Value 1;

            Format-Color -Message "Installing and config aria2 downloader..." -ForegroundColor Blue;
            scoop install aria2;
            scoop config aria2-enabled true;
            scoop config aria2-retry-wait 4;
            scoop config aria2-split 6;
            scoop config aria2-max-connection-per-server 16;
            scoop config aria2-min-split-size 4M;

            Format-Color -Message "Installing git and add extra source to scoop..." -ForegroundColor Blue;
            scoop bucket add extras;
            scoop update;
            # 7Zip will be installed while installing git 
            scoop install git;
            scoop bucket add scoopet https://github.com/ivaquero/scoopet;
            scoop update;
        }
        catch {
            throw $_.Exception.Message
        }
    }
    # If the file already exists, show the message and do nothing.
    else {
        Format-Color -Message "An existing Scoop installation was detected!" -ForegroundColor Orange;
    }
}

function Install-Chocolatey {
# USE: Install-Chocolatey -AppPath $ChocoPath -DownloadURL $ChocoURL
    param (
        [string]$AppPath,
        [string]$DownloadURL
    )
    
    # If the choco.exe file does not exist, go install it
    if (-not(Test-Path -Path $AppPath -PathType Leaf)) {
        try {
            Format-Color -Message "Installing Chocolatey package manager..." -ForegroundColor Blue;
            iex ((New-Object System.Net.WebClient).DownloadString($DownloadURL));
            
            # Chocolatey steup
            # ![o] Setup environment variable
            [environment]::SetEnvironmentVariable('ChocolateyInstall', 'C:\ProgramData\chocolatey', 'Machine');
            [environment]::SetEnvironmentVariable('ChocolateyToolsLocation', 'C\ProgramData\chocotools', 'Machine');

            # ![o] Setup app installation behavior
            choco feature enable -n allowGlobalConfirmation;
            choco feature enable -name=exitOnRebootDetected;

        }
        catch {
            throw $_.Exception.Message;
        }
    }
    # If the choco already installed, show the message and do nothing.
    else {
        Format-Color -Message "An existing Chocolatey installation was detected!" -ForegroundColor Orange;
    }
}

#===============================================================================
# !3. Installing Applicaitons
#===============================================================================

#-------------------------------------------------------------------------------
# !3-1. Install and setup development appplications
#-------------------------------------------------------------------------------

function Setup-Python {
    param (
        [string]$Index_URL,
        [string]$Trusted_Host
    )
    $config = Get-Content .\config.json -Raw | ConvertFrom-Json
    #$Index_URL = $config.pipSource
    
    # if Python was not installed, try to install it using Choco first
    $CheckResult = python --version | Select-String -pattern "\d"
    if (-not($CheckResult.ToString().Contains("3"))){
        choco install python --refreshenv;
    }
    # upgrade pip
    python -m pip install --upgrade pip;
    
    pip config set global.index-url $Index_URL --user;
    pip config set install.trusted-host $Trusted_Host --user;

    pip config list -v
}


#-------------------------------------------------------------------------------
# !3-2. Install common applications
#-------------------------------------------------------------------------------

<# TODO: Applicaitons list

Utilities
    - Windows power toys
    - Quicklook

Office Tools:
    - Office 2019
    - WPS
    - Bukeng Office Box
    - Sumartra
    - Adobe Acrobat

Media Tools:
    - Immage Related:
        - GIMP
        - Imageglass
        - ScreenX
    - Video Related:
        - Potplayer
        - FFmpeg-GUI
    - Audio Related:
        - Audicity

Security:
    - HuoRong
#>

# TODO: 1. reconfigure the downloading urls into config json file
# TODO: 2. adding windows activation method for Windows operating system under win10

# Setup command line execution policy
Set-ExecutionPolicy Bypass -Scope Process -Force;
# Enable TLSv1.2 for compatibility with older clients
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072;

$ScoopPath = "C:\Users\$($Env:username)\scoop\shims"
$ScoopApp = "C:\Users\$($Env:username)\scoop\shims\scoop"
$ScoopPS1 = "C:\Users\$($Env:username)\scoop\shims\scoop.ps1"
$ChocoPath = "C:\ProgramData\Chocolatey\Choco.exe"
$ChocoURL = "https://community.chocolatey.org/install.ps1"

$configDownloadURL = 'https://raw.githubusercontent.com/fihyer/Windows-Post-Installation-Script/main/configs/config.json'
$configFile = "$env:TEMP\config.json"
Invoke-WebRequest -Uri $configDownloadURL -UseBasicParsing -OutFile $configFile -ErrorAction Stop

do {
    Format-Color -Message @"
       ______________________________________________________________

                Windows Post Installation ToolBox

             [1] Activate Microsoft Products
             [2] Install Chocolatey Package Management Tool
             [3] Install Scoop installer Management Tool
             [4] Dev Environment Setup
             [Q] Exit
             __________________________________________________      

             Author     :   Runze Sun
             Version    :   V1.3
             License    :   MIT
       ______________________________________________________________
"@
    Write-Host @"

            Enter a menu option in the keyboard [1,2,3,Q] to continue
"@ -NoNewline -ForegroundColor Green

    $Choice = Read-Host ' '
    
    switch ($Choice) {
        1 { Activate-Microsoft }
        2 { Install-Chocolatey -AppPath $ChocoPath -DownloadURL $ChocoURL }
        3 { Install-Scoop -AppPath $ScoopPath -App $ScoopApp -PS1File $ScoopPS1 }
    }
} until ($Choice -eq 'q')
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<END POST INSTALLATION