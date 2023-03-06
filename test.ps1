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


function Activate-Microsoft{

    Format-Color -Message "Start to activiting Windows products" -ForegroundColor Blue

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

do {
    Format-Color -Message @"
       ______________________________________________________________

                Windows Post Installation ToolBox

             [1] Activate Microsoft Products
             [2] Install Package Management Tools
             [3] Install Applications
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

    $Choice=Read-Host ' '
    
    switch ($Choice){
        1 {Activate-Microsoft}
        2 {Format-Color -Message "Install Package Mangement Tools"}
        3 {Format-Color -Message "Install Applications"}
    }
} until ($Choice -eq 'q')