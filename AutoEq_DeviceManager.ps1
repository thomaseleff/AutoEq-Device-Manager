# Set-PsDebug -trace 1

# Import Modules
try {
    if (
        Get-Module -ListAvailable -Name AudioDeviceCmdlets
    ) {
        Import-Module -Name AudioDeviceCmdlets
    } else {
        Install-Module -Name AudioDeviceCmdlets
        Import-Module -Name AudioDeviceCmdlets
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show('ERROR: Run with Admin priviledges to install the AudioDeviceCmdlets module.')
    Write-Error -Message ('ERROR: Run with Admin priviledges to install the AudioDeviceCmdlets module.') -ErrorAction Stop
}

# Import Assemblies
Add-Type -AssemblyName System.Windows.Forms 
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Speech

# Assign Global Variables
Set-Variable -Name 'fontReg' -Value ([System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Regular))
Set-Variable -Name 'fontBold' -Value ([System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold))
Set-Variable -Name 'fontItalic' -Value ([System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Italic))
Set-Variable -Name 'outputHeader' -Value ([System.Drawing.Font]::new('Courier New', 9, [System.Drawing.FontStyle]::Bold))
Set-Variable -Name 'outputBody' -Value ([System.Drawing.Font]::new('Courier New', 9, [System.Drawing.FontStyle]::Regular))

# Initialize System Tray Dictionary Object
$appDict = @{
    'dir' = (split-path $MyInvocation.MyCommand.Path -Parent)
    'sysTrayAppName' = 'Audio Device Manager'
    'version' = (New-Object System.Version('2.0.0'))
    'reg' = 'HKLM:HKEY_LOCAL_MACHINE\SOFTWARE'
    'regUser' = 'HKCU:HKEY_CURRENT_USER\Software'
    'windowsTheme' = ((Get-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\ -Name "SystemUsesLightTheme").SystemUsesLightTheme)
    'appTheme' = ((Get-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\ -Name "AppsUseLightTheme").AppsUseLightTheme)
    'debug' = $false
}

# Create Text to Speech Synth Object
Set-Variable -Name 'synth' -Value (New-Object System.Speech.Synthesis.SpeechSynthesizer)
$synth.Rate = -2
$global:narrator = $false

# Enable Notifications
$global:notifications = $true

# Initialize Config
if (
    -Not [System.IO.File]::Exists($appDict['dir']+'\config\config.txt')
) {
    Out-File -FilePath ($appDict['dir']+'\config\config.txt')
}

# Define Functions

function create_icon {
    param (
        $unicodeChar,
        $size,
        $theme,
        $x,
        $y,
        $unicodeChar1 = $false,
        $size1 = $false,
        $color1 = $false,
        $x1 = $false,
        $y1 = $false,
        $unicodeChar2 = $false,
        $size2 = $false,
        $color2 = $false,
        $x2 = $false,
        $y2 = $false
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $unicodeChar            = [Str] Segoe MDL Icon Asset as Unicode
    $size                   = [Int] Font size of the $unicodeChar
    $theme                  = [Bool] Windows theme (0 - "Light Theme", 1 - "Dark Theme")
    $x                      = [Int] Horizontal position of the $unicodeChar
    $y                      = [Int] Vertical position of the $unicodeChar
    $unicodeChar1           = [Str] Second Segoe MDL Icon Asset as Unicode
    $size                   = [Int] Font size of the $unicodeChar
    $theme                  = [Bool] Windows theme (0 - "Light Theme", 1 - "Dark Theme")
    $x1                     = [Int] Horizontal position of the $unicodeChar1
    $y1                     = [Int] Vertical position of the $unicodeChar1
    $unicodeChar2           = [Str] Third Segoe MDL Icon Asset as Unicode
    $size                   = [Int] Font size of the $unicodeChar
    $theme                  = [Bool] Windows theme (0 - "Light Theme", 1 - "Dark Theme")
    $x2                     = [Int] Horizontal position of the $unicodeChar2
    $y2                     = [Int] Vertical position of the $unicodeChar2

    Description
    ----------------------------------------------------------------------------------------------------
    Layers up to three Segoe MDL Icon Assets into a single bitmap and returns an icon.
    #>

    # Set Icon Color Based on Theme
    if (
        $theme -eq 1
    ) {
        Set-Variable -Name 'color' -Value 'Black'
    } else {
        Set-Variable -Name 'color' -Value 'White'
    }

    # Create Icon Bitmap from Segoe MDL2 Assets
    Set-Variable -Name 'fontIcon' -Value ([System.Drawing.Font]::new('Segoe MDL2 Assets', $size, [System.Drawing.FontStyle]::Regular))
    $brush = [System.Drawing.Brushes]::$color
    $bitmap = New-Object System.Drawing.Bitmap 256,256
    $bitmapGraphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $bitmapGraphics.DrawString([char]($unicodeChar), $fontIcon, $brush, $x, $y)
    if (
        $unicodeChar1 -ne $false
    ) {
        $brush1 = [System.Drawing.Brushes]::$color1
        $bitmapGraphics.DrawString([char]($unicodeChar1), $fontIcon, $brush1, $x1, $y1)
    }
    if (
        $unicodeChar2 -ne $false
    ) {
        $brush2 = [System.Drawing.Brushes]::$color2
        $bitmapGraphics.DrawString([char]($unicodeChar2), $fontIcon, $brush2, $x2, $y2)
    }
    $bitmapGraphics.SmoothingMode = 'AntiAlias'
    $bitmapGraphics.TextRenderingHint = 'AntiAliasGridFit'
    $bitmapGraphics.InterpolationMode = 'High'

    # Cleanup Graphics Object
    $bitmapGraphics.Flush()
    $bitmapGraphics.Dispose()

    return [System.Drawing.Icon]::FromHandle($bitmap.GetHicon())
}

function write_output {
    param (
        $appDict,
        $returnAfter,
        $str
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $appDict                = [Dict] Dictionary containing the parameters essential to the application
    $returnAfter            = [Bool] Indicates to output a blank line following $str
    $str                    = [Str] Text to write to ~\output.txt

    Description
    ----------------------------------------------------------------------------------------------------
    Writes text to ~\output.txt.
    #>

    # Write Str to Console
    # Write-Host $str

    # Write Str to output.txt
    if (
        [System.IO.File]::Exists($appDict['dir']+'\config\output.txt')
    ) {
        if (
            $returnAfter
        ) {
            Out-File -FilePath ($appDict['dir']+'\config\output.txt') -InputObject $str -Encoding utf8 -Append
            Out-File -FilePath ($appDict['dir']+'\config\output.txt') -InputObject '' -Encoding utf8 -Append
        } else {
            Out-File -FilePath ($appDict['dir']+'\config\output.txt') -InputObject $str -Encoding utf8 -Append
        }
    } else {
        if (
            $returnAfter
        ) {
            Out-File -FilePath ($appDict['dir']+'\config\output.txt') -InputObject $str -Encoding utf8
            Out-File -FilePath ($appDict['dir']+'\config\output.txt') -InputObject '' -Encoding utf8 -Append
        } else {
            Out-File -FilePath ($appDict['dir']+'\config\output.txt') -InputObject $str -Encoding utf8
        }
    }
}

function create_reg {
    param (
        $appDict
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $appDict                = [Dict] Dictionary containing the parameters essential to the application

    Description
    ----------------------------------------------------------------------------------------------------
    Creates the registry directory within $appDict['regUser'] to store application settings.
    #>

    # Create Registry Directory
    if (
        -Not (Test-Path ($appDict['regUser']+'\EqualizerAPO'))
    ) {
        $apoReg = New-Item -Path ($appDict['regUser']) -Name 'EqualizerAPO'
        $appReg = New-Item -Path ($appDict['regUser']+'\EqualizerAPO') -Name 'Device Manager'
        $deviceReg = New-Item -Path ($appDict['regUser']+'\EqualizerAPO\Device Manager') -Name 'Playback Devices'
        $profileReg = New-Item -Path ($appDict['regUser']+'\EqualizerAPO\Device Manager') -Name 'Equalizer Profiles'
    } else {
        if (
            -Not (Test-Path ($appDict['regUser']+'\EqualizerAPO\Device Manager'))
        ) {
            $appReg = New-Item -Path ($appDict['regUser']+'\EqualizerAPO') -Name 'Device Manager'
            $deviceReg = New-Item -Path ($appDict['regUser']+'\EqualizerAPO\Device Manager') -Name 'Playback Devices'
            $profileReg = New-Item -Path ($appDict['regUser']+'\EqualizerAPO\Device Manager') -Name 'Equalizer Profiles'
        }
    }
}

function convert_reg_to_dict {
    param (
        $appDict,
        $subKey
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $appDict                = [Dict] Dictionary containing the parameters essential to the application
    $subKey                 = [Str] Sub-key within the ~\EqualizerAPO\Device Manager\ registry

    Description
    ----------------------------------------------------------------------------------------------------
    Converts the sub-key registry and its parameters and values into a dictionary object and
        returns the dictionary.
    #>

    Set-Variable -Name 'dict' -Value @{}

    # Convert Registry Subkey to Dictionary of Dictionaries
    if (
        (Test-Path ($appDict['regUser']+'\EqualizerAPO\Device Manager\'+$subKey))
    ) {
        $keyLst = @(Get-ChildItem -Path ($appDict['regUser']+'\EqualizerAPO\Device Manager\'+$subKey) | Select-Object Name)

        if (
            $keyLst
        ) {
            $keyLst | ForEach-Object {
                $subDict = @{}
                $key = ($_.Name -split '\\')[-1]
                $subKeyProperties = Get-Item -Path ($appDict['regUser']+'\EqualizerAPO\Device Manager\'+$subKey+'\'+$key) | Select-Object -ExpandProperty Property
                
                $subKeyProperties | ForEach-Object {
                    $subDict[$_] = (Get-ItemPropertyValue -Path ($appDict['regUser']+'\EqualizerAPO\Device Manager\'+$subKey+'\'+$key) -Name $_)
                }
    
                # Add Reg Properties to List
                $dict[$key] = $subDict
            }
        } else {
            $dict = $false
        }
    }

    return $dict
}

function push_dict_to_reg {
    param (
        $appDict,
        $subKey,
        $dict,
        $properties
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $appDict                = [Dict] Dictionary containing the parameters essential to the application
    $subKey                 = [Str] Sub-key within the ~\EqualizerAPO\Device Manager\ registry
    $dict                   = [Dict] Dictionary to convert to a registry directory.
    $properties             = [List] Properties within the dictionary to convert to registry keys.

    Description
    ----------------------------------------------------------------------------------------------------
    Converts the dictionary and its parameters and values into a registry directory.
    #>

    # Convert Dictionary of Dictionaries to Registry Subkeys
    if (
        (Test-Path ($appDict['regUser']+'\EqualizerAPO\Device Manager\'+$subKey))
    ) {
        if (
            $dict
        ) {
            ForEach ($subDict in $dict.Keys) {
                if (
                    -Not (Test-Path ($appDict['regUser']+'\EqualizerAPO\Device Manager\'+$subKey+'\'+$subDict))
                ) {
                    $regKey = New-Item -Path ($appDict['regUser']+'\EqualizerAPO\Device Manager\'+$subKey) -Name $subDict
                    ForEach ($key in $dict[$subDict].Keys) {
                        if (
                            $key -in $properties
                        ) {
                            $regKey.SetValue($key, $dict[$subDict][$key])
                        }
                    }
                } else {
                    ForEach ($key in $dict[$subDict].Keys) {
                        if (
                            $key -in $properties
                        ) {
                            Set-ItemProperty -Path ($appDict['regUser']+'\EqualizerAPO\Device Manager\'+$subKey+'\'+$subDict) -Name $key -Value $dict[$subDict][$key]
                        }
                    }
                }
            }
        }
    }
}

function raise_notification {
    param (
        $sysTrayApp,
        $appDict,
        $type,
        $str
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $sysTrayApp             = [System.Windows.Forms.NotifyIcon] System tray application object
    $appDict                = [Dict] Dictionary containing the parameters essential to the application
    $type                   = [Str] Indicates the type of notification to raise (e.g. warning, info, etc.)
    $str                    = [Str] Text to display in the balloon tip notification

    Description
    ----------------------------------------------------------------------------------------------------
    Raises a balloon tip notification.
    #>

    # Raise Notification
    $sysTrayApp.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::$type
    $sysTrayApp.BalloonTipTitle = $appDict['sysTrayAppName']
    $sysTrayApp.BalloonTipText = $str
    $sysTrayApp.ShowBalloonTip(500)
}

function eval_connection {
    param (
        $appDict,
        $domain
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $appDict                = [Dict] Dictionary containing the parameters essential to the application
    $domain                 = [Str] URL

    Description
    ----------------------------------------------------------------------------------------------------
    Validates the user's internet connection to $domain, updates $appDict and returns $appDict.
    #>

    # Validate Internet Connection
    if (
        -Not (Test-Connection -ComputerName $domain -Quiet -Count 1)
    ) {

        # Capture Connection Error for Notification
        $appDict['connectionErr'] = $true
        write_output -returnAfter $false -appDict $appDict -str ('ERROR: No Internet connection to ['+$domain+']. Connect to the Internet or resolve connection issues.')
    }

    # Debug
    if (
        $appDict['debug']
    ) {
        Write-Host 'connectionErr               : '$appDict['connectionErr']
    }

    return $appDict
}

function eval_version {
    param (
        $appDict
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $appDict                = [Dict] Dictionary containing the parameters essential to the application

    Description
    ----------------------------------------------------------------------------------------------------
    Evaluates whether a newer version of the system tray application is available, updates
        $appDict and returns $appDict.
    #>

    # Retrieve Latest Production Release Version
    if (
        -Not $appDict['connectionErr']
    ) {

        # Request HTML
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $request = Invoke-WebRequest -Uri 'https://github.com/thomaseleff/AutoEq-Device-Manager/releases/latest'

        # Parse Release Version
        $versionLatest = New-Object System.Version((($request.BaseResponse.ResponseUri.AbsoluteUri -split '\/')[-1] -replace "[^0-9.]"))

        # Validate System Tray Application Version
        if (
            (($appDict['version']).CompareTo($versionLatest)) -lt 0
        ) {

            # Capture Version Mismatch for Notification
            $appDict['versionMismatch'] = $true
            write_output -returnAfter $true -appDict $appDict -str ('NOTE: A new version, v'+$versionLatest+', is available on [www.github.com].')
        }
    }

    # Debug
    if (
        $appDict['debug']
    ) {
        Write-Host 'versionMismatch             : '$appDict['versionMismatch']
    }

    return $appDict
}

function eval_auto_eq_conf {
    param (
        $appDict
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $appDict                = [Dict] Dictionary containing the parameters essential to the application

    Description
    ----------------------------------------------------------------------------------------------------
    Validates the user-provided autoeq_config.json, updates $appDict and returns a json object
        containing the user-provided selections and returns $appDict.
    #>

    # Manage Depreciated EQ Profiles Config
    if (
        [System.IO.File]::Exists($appDict['dir']+'\config\eq_profiles.json')
    ) {

        # Backup Depreciated EQ Profiles Config
        Copy-Item -Path ($appDict['dir']+'\config\eq_profiles.json') -Destination ($appDict['dir']+'\config\eq_profiles_depreciated_'+(Get-Date -Format 'MM_dd_yyyy_HH_mm_ss')+'.json')

        # Copy Depreciated EQ Profiles Config into AutoEq Config
        Copy-Item -Path ($appDict['dir']+'\config\eq_profiles.json') -Destination ($appDict['dir']+'\config\autoeq_config.json')

        # Delete Depreciated EQ profiles Config
        Remove-Item -Path ($appDict['dir']+'\config\eq_profiles.json')
    }

    # Validate User-Provided AutoEq Config
    if (
        [System.IO.File]::Exists($appDict['dir']+'\config\autoeq_config.json')
    ) {

        # Logging
        write_output -returnAfter $true -appDict $appDict -str ('NOTE: autoeq_config.json exists within the \config folder as expected.')

        try {
            $autoEqConf = Get-Content -Raw -Path ($appDict['dir']+'\config\autoeq_config.json') | ConvertFrom-Json
            $autoEqConf = $autoEqConf.GetEnumerator() | Sort-Object device
            # $autoEqConf | ConvertTo-Json | Out-File ($appDict['dir']+'\config\autoeq_config.json')
        } catch {

            # Capture Json Error for Notification
            $appDict['invalidConfigErr'] = $true
            write_output -returnAfter $false -appDict $appDict -str ("ERROR: autoeq_config.json is not a valid '.json' file.")

            # Replace autoeq_config.json
            if (
                -Not $appDict['connectionErr']
            ) {

                # Backup Existing autoeq_config.json
                Copy-Item -Path ($appDict['dir']+'\config\autoeq_config.json') -Destination ($appDict['dir']+'\config\autoeq_config_backup_'+(Get-Date -Format 'MM_dd_yyyy_HH_mm_ss')+'.json')

                # Download Latest Valid autoeq_config.json
                try {
                    Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/thomaseleff/AutoEq-Device-Manager/main/config/autoeq_config.json' -OutFile ($appDict['dir']+'\config\autoeq_config.json')
                } catch {
                    Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/thomaseleff/AutoEq-Device-Manager/main/config/eq_profiles.json' -OutFile ($appDict['dir']+'\config\autoeq_config.json')
                }

                # Logging
                write_output -returnAfter $false -appDict $appDict -str ('     ~ The current autoeq_config.json has been backed up as autoeq_config_backup_'+(Get-Date -Format 'MM_dd_yyyy_HH_MM_SS')+'.json within the \config folder.')
                write_output -returnAfter $false -appDict $appDict -str ('     ~ The latest autoeq_config.json has been downloaded from [www.github.com].')
                write_output -returnAfter $false -appDict $appDict -str ("     ~ Modify the new template autoeq_config.json and then click 'Refresh' from the tool menu.")
            }
        }
    } else {

        # Capture Profile Warning for Notification
        $appDict['missingConfigErr'] = $true
        write_output -returnAfter $false -appDict $appDict -str ('ERROR: autoeq_config.json not found within the \config folder.')

        # Download autoeq_config.json
        if (
            -Not $appDict['connectionErr']
        ) {

            # Download Latest Valid autoeq_config.json
            try {
                Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/thomaseleff/AutoEq-Device-Manager/main/config/autoeq_config.json' -OutFile ($appDict['dir']+'\config\autoeq_config.json')
            } catch {
                Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/thomaseleff/AutoEq-Device-Manager/main/config/eq_profiles.json' -OutFile ($appDict['dir']+'\config\autoeq_config.json')
            }

            # Logging
            write_output -returnAfter $false -appDict $appDict -str ('     ~ The latest autoeq_config.json has been downloaded from [www.github.com].')
            write_output -returnAfter $false -appDict $appDict -str ("     ~ Modify the new template autoeq_config.json and then click 'Refresh' from the tool menu.")
        }
    }
    
    # Logging
    write_output -returnAfter $false -appDict $appDict -str ('')

    # Check Empty Json
    if (
        -Not $autoEqConf
    ) {
        $autoEqConf = $false
    }

    # Debug
    if (
        $appDict['debug']
    ) {
        Write-Host 'invalidConfigErr            : '$appDict['invalidConfigErr']
        Write-Host 'missingConfigErr            : '$appDict['missingConfigErr']
    }

    return $autoEqConf, $appDict
}

function eval_dict_against_settings {
    param (
        $dict,
        $settings
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $dict                   = [Dict] Dictionary object to compare against the registry settings
    $settings               = [Dict] Dictionary object containing the registry settings

    Description
    ----------------------------------------------------------------------------------------------------
    Validates a dictionary object against the registry settings and returns a parameter indicating
        whether new keys exist in $dict that are not found in $settings.
    #>

    Set-Variable -Name 'newSetting' -Value $false

    # Evaluate If New Items Exist in the Dictionary Object that are Not in the Settings
    if (
        $settings
    ) {
        if (
            $dict
        ) {
            $dict.Keys | ForEach-Object {
                if (
                    $_ -NotIn $settings.Keys
                ) {
                    $newSetting = $true
                }
            }
        }
    } else {
        $newSetting = $true
    }

    return $newSetting
}

function retrieve_apo_devices {
    param (
        $appDict
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $appDict                = [Dict] Dictionary containing the parameters essential to the application

    Description
    ----------------------------------------------------------------------------------------------------
    Returns a list of Playback Audio Device IDs that have Equalizer APO installed.
    #>

    # Initialize Local Variables
    Set-Variable -Name 'apoChildLst' -Value @(Get-ChildItem -Path ($appDict['reg']+'\EqualizerAPO\Child APOs') | Select-Object Name)
    Set-Variable -Name 'apoDeviceIdLst' -Value @()
    
    # Retrieve List of Audio Devices with APO Installed
    if (
        $apoChildLst
    ) {
        $apoChildLst | ForEach-Object {
            $apoDeviceIdLst += ($_.Name -split '\\')[-1]
        }
    }

    return $apoDeviceIdLst
}

function retrieve_eq_profiles {
    param (
        $appDict,
        $autoEqConf
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $appDict                = [Dict] Dictionary containing the parameters essential to the application
    $autoEqConf             = [Json] Json object containing the user-provided selections for autoEq
                                profiles

    Description
    ----------------------------------------------------------------------------------------------------
    Validates the user-provided selections for autoEq profiles, retrieves them from the Github
        repository, returns a dictionary containing the profiles, their parameters and values,
        and returns $appDict.
    #>

    # Initialize Equalizer Profile Dictionary Object
    $global:equalizerProfileLst = @{}

    if (
        ((-Not $appDict['connectionErr']) -And (-Not $appDict['invalidConfigErr']) -And (-Not $appDict['missingConfigErr']) -And ($autoEqConf))
    ) {

        # Retrieve
        $autoEqConf | ForEach-Object {
            if (
                ('device' -notin $_.PSobject.Properties.Name -And 'parametricConfig' -notin $_.PSobject.Properties.Name)
            ) {

                # Capture Profile Error for Notification
                if (
                    -Not $appDict['profileErr']
                ) {
                    $appDict['profileErr'] = $true
                }
                write_output -returnAfter $false -appDict $appDict -str ('Invalid EQ Profile Entry')
                write_output -returnAfter $false -appDict $appDict -str ('------------------------')
                write_output -returnAfter $false -appDict $appDict -str ('    ERROR: Invalid EQ Profile entry in autoeq_config.json.')
                write_output -returnAfter $true -appDict $appDict -str ('         ~ [device] and [parametricConfig] parameters not found.')
            } elseif (
                'device' -notin $_.PSobject.Properties.Name
            ) {

                # Capture Profile Error for Notification
                if (
                    -Not $appDict['profileErr']
                ) {
                    $appDict['profileErr'] = $true
                }
                write_output -returnAfter $false -appDict $appDict -str ('Invalid EQ Profile Entry')
                write_output -returnAfter $false -appDict $appDict -str ('------------------------')
                write_output -returnAfter $false -appDict $appDict -str ('    ERROR: Invalid EQ Profile entry in autoeq_config.json.')
                write_output -returnAfter $true -appDict $appDict -str ('         ~ [device] parameter not found.')
            } elseif (
                'parametricConfig' -notin $_.PSobject.Properties.Name
            ) {

                # Capture Profile Error for Notification
                if (
                    -Not $appDict['profileErr']
                ) {
                    $appDict['profileErr'] = $true
                }
                if (
                    ([string]::IsNullOrWhiteSpace($_.device)) 
                ) {
                    write_output -returnAfter $false -appDict $appDict -str ('Invalid EQ Profile Entry')
                    write_output -returnAfter $false -appDict $appDict -str ('------------------------')
                    write_output -returnAfter $false -appDict $appDict -str ('    ERROR: Invalid EQ Profile entry in autoeq_config.json.')
                    write_output -returnAfter $false -appDict $appDict -str ('         ~ [device] parameter exists but is null, empty or contains only white space.')
                    write_output -returnAfter $true -appDict $appDict -str ('         ~ [parametricConfig] parameter not found.')
                } else {
                    write_output -returnAfter $false -appDict $appDict -str ($_.device)
                    write_output -returnAfter $false -appDict $appDict -str ('+'+('-' * ($_.device.Length-2)+'+'))
                    write_output -returnAfter $false -appDict $appDict -str ('    ERROR: Invalid EQ Profile entry in autoeq_config.json.')
                    write_output -returnAfter $true -appDict $appDict -str ('         ~ [parametricConfig] parameter not found.')
                }

            } elseif (
                ([string]::IsNullOrWhiteSpace($_.device))
            ) {
                
                # Capture Profile Error for Notification
                if (
                    -Not $appDict['profileErr']
                ) {
                    $appDict['profileErr'] = $true
                }
                write_output -returnAfter $false -appDict $appDict -str ('Invalid EQ Profile Entry')
                write_output -returnAfter $false -appDict $appDict -str ('------------------------')
                write_output -returnAfter $false -appDict $appDict -str ('    ERROR: Invalid EQ Profile entry in autoeq_config.json.')
                write_output -returnAfter $true -appDict $appDict -str ('         ~ [device] parameter exists but is null, empty or contains only white space.')
            } else {

                # Logging
                write_output -returnAfter $false -appDict $appDict -str ($_.device)
                write_output -returnAfter $false -appDict $appDict -str ('-' * ($_.device.Length))
                if (
                    ($_.parametricConfig.ToLower().Contains('https://raw.githubusercontent.com/jaakkopasanen/AutoEq/master/results/'.ToLower()) -And $_.parametricConfig.ToLower().Contains('ParametricEQ.txt'.ToLower()))
                ) {
                    if (
                        -Not [System.IO.File]::Exists($appDict['dir']+'\config\Parametric_EQ_'+$_.device+'.txt')
                    ) {
                        try {
                            Invoke-WebRequest -Uri $_.parametricConfig -OutFile ($appDict['dir']+'\config\Parametric_EQ_'+$_.device+'.txt')
    
                            # Logging
                            write_output -returnAfter $false -appDict $appDict -str ('    NOTE: URL to [parametricConfig] is valid.')
                            write_output -returnAfter $true -appDict $appDict -str ('    NOTE: Parametric EQ Config retrieved successfully.')
    
                            # Add Profile to Valid List
                            $global:equalizerProfileLst[$_.device] = @{
                                'Name' = $_.device
                                'DeviceType' = 'Headphone'
                                'Source' = 'AutoEq'
                            }
                        } catch {
    
                            # Capture Profile Error for Notification
                            if (
                                -Not $appDict['profileErr']
                            ) {
                                $appDict['profileErr'] = $true
                            }
                            write_output -returnAfter $true -appDict $appDict -str ('    ERROR: URL to [parametricConfig] not found.')
                        }     
                    } else {
                        write_output -returnAfter $true -appDict $appDict -str ('    NOTE: Passing, Parametric EQ Config already exists within the \config folder.')
    
                        # Add Profile to Valid List
                        $global:equalizerProfileLst[$_.device] = @{
                            'Name' = $_.device
                            'DeviceType' = 'Headphone'
                            'Source' = 'AutoEq'
                        }
                    }
                } else {
    
                    # Capture Profile Error for Notification
                    if (
                        -Not $appDict['profileErr']
                    ) {
                        $appDict['profileErr'] = $true
                    }
                    write_output -returnAfter $false -appDict $appDict -str ('    ERROR: URL to [parametricConfig] is invalid or is null, empty or contains only white space.')
                    write_output -returnAfter $false -appDict $appDict -str ('         ~ The URL must begin with [https://raw.githubusercontent.com/] and link to a')
                    write_output -returnAfter $true -appDict $appDict -str ('         ~ Parametric EQ profile within the [jaakkopasanen/AutoEQ] Github project.')
                }
            }
        }
    }

    # Add User-Created Profiles
    $userProfileLst = @(Get-ChildItem -Path ($appDict['dir']+'\config') | Select-Object Name | Where-Object {$_.Name -like 'Parametric_EQ_*.txt'})

    if (
        $userProfileLst
    ) {

        # Add Profile to Valid List
        $userProfileLst | ForEach-Object {
            if (
                (($_.Name -split '_')[2] -split '\.')[0] -NotIn $global:equalizerProfileLst.Keys
            ) {
                $global:equalizerProfileLst[(($_.Name -split '_')[2] -split '\.')[0]] = @{
                    'Name' = (($_.Name -split '_')[2] -split '\.')[0]
                    'DeviceType' = 'Headphone'
                    'Source' = 'User'
                }
            }
        }
    }

    # Check Empty Valid Profiles
    if (
        -Not $global:equalizerProfileLst
    ) {
        $global:equalizerProfileLst = $false
    }

    # Debug
    if (
        $appDict['debug']
    ) {
        Write-Host 'profileErr                  : '$appDict['profileErr']
        Write-Host 'equalizerProfileLst         : '$global:equalizerProfileLst.Keys
    }

    return $global:equalizerProfileLst, $appDict
}

function retrieve_playback_devices {
    param (
        $appDict
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $appDict                = [Dict] Dictionary containing the parameters essential to the application

    Description
    ----------------------------------------------------------------------------------------------------
    Validates the active Playback Devices, returns a dictionary containing the Playback Devices, their
        parameters and values, and returns $appDict.
    #>

    # Initialize Local Variables and Playback Device Dictionary Object
    Set-Variable -Name 'index' -Value 0
    Set-Variable -Name 'done' -Value $false
    $global:playbackDeviceLst = @{}

    while (
        -Not $done
    ) {
        $index += 1

        # Check If an Audio Device is Found at Each Index
        try {
            Set-Variable -Name 'audioDevice' -Value (Get-AudioDevice -Index $index)
        } catch {
            Set-Variable -Name 'returnedVal' -Value $_
            Set-Variable -Name 'returnedVal' -Value ($returnedVal -Replace '\s','')
            if (
                $returnedVal -eq 'NoAudioDevicewiththatIndex'
            ) {
                $done = $true
            } else {
                continue
            }
        }

        # Record Audio Device Parameters for Only Playback Devices
        if (
            $audioDevice.Type -eq 'Playback'
        ) {

            # Assign a Default Device Type
            if (
                $audioDevice.Name.ToLower().Contains('speaker')
            ) {
                $deviceType = 'Speaker'
            } elseif (
                ($audioDevice.Name.ToLower().Contains('headphone')) -Or
                ($audioDevice.Name.ToLower().Contains('earphone')) -Or
                ($audioDevice.Name.ToLower().Contains('headset')) -Or
                ($audioDevice.Name.ToLower().Contains('head phone')) -Or
                ($audioDevice.Name.ToLower().Contains('ear phone')) -Or
                ($audioDevice.Name.ToLower().Contains('head set'))
            ) {
                $deviceType = 'Headphone'
            } else {
                $deviceType = 'Line-Out'
            }

            $global:playbackDeviceLst[$audioDevice.Name] = @{
                'Index' = $index
                'Name' = ($index, $audioDevice.Name -join ': ')
                'Default' = $audioDevice.Default
                'DeviceType' = $deviceType
                'DefaultProfile' = 'No Profile'
                'Type' = $audioDevice.Type
                'ID' = $audioDevice.ID
            }
        }
    }

    # Capture APO Device Warning for Notification
    ForEach ($device in $global:playbackDeviceLst.Keys) {
        if (
            ($global:playbackDeviceLst[$device]['ID'] -split '\.')[-1] -in (retrieve_apo_devices -appDict $appDict)
        ) {
            $appDict['apoErr'] = $false
        }
    }

    # Logging
    if (
        $appDict['apoErr']
    ) {
        write_output -returnAfter $false -appDict $appDict -str ('ERROR: EqualizerAPO is not installed for any connected Audio Devices.')
        write_output -returnAfter $false -appDict $appDict -str ("     ~ Click 'Open Configurator' from the tool menu to enable EqualizerAPO for any connected Audio Device.")
    }

    # Capture No Audio Devices Error
    if (
        -Not $global:playbackDeviceLst
    ) {
        $global:playbackDeviceLst = $false
        $appDict['missingPlaybackDeviceErr'] = $true

        # Logging
        write_output -returnAfter $false -appDict $appDict -str ('ERROR: No connected Audio Device(s) found. Ensure at least one Audio Device is enabled within the system device settings.')
    }

    # Debug
    if (
        $appDict['debug']
    ) {
        Write-Host 'apoErr                      : '$appDict['apoErr']
        Write-Host 'missingPlaybackDeviceErr    : '$appDict['missingPlaybackDeviceErr']
        Write-Host 'playbackDeviceLst           : '$global:playbackDeviceLst.Keys
    }

    return $global:playbackDeviceLst, $appDict
}

function apply_settings_to_dict {
    param (
        $dict,
        $settings
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $dict                   = [Dict] Dictionary object
    $settings               = [Dict] Dictionary object containing the registry settings

    Description
    ----------------------------------------------------------------------------------------------------
    Updates a dictionary object with the registry settings and returns the $dict.
    #>

    # Apply Settings to Dictionary Object
    if (
        $settings
    ) {
        if (
            $dict
        ) {
            ForEach ($subDict in $dict.Keys) {
                if (
                    $subDict -In $settings.Keys
                ) {
                    ForEach ($key in $settings[$subDict].Keys) {
                        if (
                            $key -In $dict[$subDict].Keys
                        ) {
                            $dict[$subDict][$key] = $settings[$subDict][$key]
                        }
                    }
                }
            }
        }
    }

    return $dict
}

function update_settings_dict {
    param (
        $dict,
        $subDict,
        $key,
        $value
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $dict                   = [Dict] Dictionary object
    $subDict                = [Str] Key within $dict that contains a dictionary object
    $key                    = [Str] Existing setting parameter
    $value                  = [Str] New setting value

    Description
    ----------------------------------------------------------------------------------------------------
    Updates a dictionary object with the registry settings and returns the $dict.
    #>

    $dict[$subDict][$key] = $value

    return $dict
}

function filter_dict_keys {
    param (
        $dict,
        $filterOn,
        $value
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $dict                   = [Dict] Dictionary object
    $filterOn               = [Str] Parameter within each sub-dictionary to filter on
    $value                  = [Str] Parameter value to filter

    Description
    ----------------------------------------------------------------------------------------------------
    Returns a sorted list of $dict keys where $filterOn = $value.
    #>

    $lst = @()

    # Return Filtered List of Dictionary Keys Based on Key & Value
    if (
        $dict
    ) {
        ForEach($subDict in $Dict.Keys) {
            if (
                $dict[$subDict][$filterOn] -eq $value
            ) {
                $lst += $subDict
            }
        }
    }

    return ($lst | Sort-Object)
}

function add_audio_device_click_action {
    param (
        $sysTrayApp,
        $appDict,
        $selectedItem,
        $parentItem,
        $reset
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $sysTrayApp             = [System.Windows.Forms.NotifyIcon] System tray application object
    $appDict                = [Dict] Dictionary containing the parameters essential to the application
    $selectedItem           = [Str] Indicates the selected text of the system tray menu item
    $parentItem             = [Str] Indicates the text of the parent system tray menu item
    $reset                  = [Bool] Indicates whether to reset ~\config.txt

    Description
    ----------------------------------------------------------------------------------------------------
    Switches playback audio devices, applies the selected equalizer profile and sets the tool-tip title
        of the system tray application.
    #>

    # Assign Local Variables
    if (
        -Not $parentItem
    ) {
        Set-Variable -Name 'selectedDeviceName' -Value ($selectedItem)
        Set-Variable -Name 'selectedProfileName' -Value 'No Profile'
    } else { 
        Set-Variable -Name 'selectedDeviceName' -Value ($parentItem)
        Set-Variable -Name 'selectedProfileName' -Value ($selectedItem)
    }
    Set-Variable -Name 'activeDevice' -Value (Get-AudioDevice -Playback)
    Set-Variable -Name 'activeDeviceName' -Value $activeDevice.Name
    Set-Variable -Name 'deviceIndex' -Value (($selectedDeviceName -split ': ')[0])
    Set-Variable -Name 'deviceUserName' -Value (($selectedDeviceName -split ': ')[1] -split ' \(')[0]
    if (
        $reset
    ) {
        Set-Variable -Name 'configName' -Value 'config.txt'
    } else {
        Set-Variable -Name 'configName' -Value ('Parametric_EQ_'+$selectedProfileName+'.txt')
    }

    # Debug
    if (
        $appDict['debug']
    ) {
        Write-Host 'Selected Item Name      :'$selectedItem
        Write-Host 'Parent Item Name        :'$parentItem
        Write-Host 'Selected Device Index   :'$deviceIndex
        Write-Host 'Selected Device Name    :'($selectedDeviceName -split ': ')[1]
        Write-Host 'Selected Profile Name   :'$selectedProfileName
        Write-Host 'Active Device Name      :'$activeDeviceName
        Write-Host 'Device User Name        :'$deviceUserName
    }

    if (
        [System.IO.File]::Exists($appDict['dir']+'\config\'+$configName)
    ) {
        if (
            $activeDeviceName -eq ($selectedDeviceName -split ': ')[1]
        ) {

            # Maintain Audio Device
            if (
                $reset
            ) {

                # Reset Parametric EQ Profile
                Out-File -FilePath ($appDict['dir']+'\config\'+$configName)
                write_output -returnAfter $false -appDict $appDict -str ('NOTE: Parametric EQ profile successfully unassigned.')
                
            } else {

                # Switch Parametric EQ Profile
                Copy-Item -Path ($appDict['dir']+'\config\'+$configName) -Destination ($appDict['dir']+'\config\config.txt')
                write_output -returnAfter $false -appDict $appDict -str ('NOTE: '+$selectedProfileName+' parametric EQ profile successfully assigned.')
            }

            # Narrate
            if (
                $global:narrator
            ) {
                $synth.Speak($selectedProfileName)
            }
        } else {

            # Switch Audio Device
            Set-AudioDevice -Index $deviceIndex
            if (
                $reset
            ) {

                # Reset Parametric EQ Profile
                Out-File -FilePath ($appDict['dir']+'\config\'+$configName)
                write_output -returnAfter $false -appDict $appDict -str ('NOTE: '+($selectedDeviceName -split ': ')[1]+' successfully assigned with no parametric EQ profile.')
            } else {

                # Switch Parametric EQ Profile
                Copy-Item -Path ($appDict['dir']+'\config\'+$configName) -Destination ($appDict['dir']+'\config\config.txt')
                write_output -returnAfter $false -appDict $appDict -str ('NOTE: '+($selectedDeviceName -split ': ')[1]+' successfully assigned with '+$selectedProfileName+' parametric EQ profile.')  
            }
            
            # Narrate
            if (
                $global:narrator
            ) {
                $synth.Speak($deviceUserName + ' with ' + $selectedProfileName)
            }
        }

        # Update App Title
        $appTitle = $appDict['sysTrayAppName']+' - '+$deviceUserName+' ('+$selectedProfileName+')'
    } else {
        if (
            $activeDeviceName -eq ($selectedDeviceName -split ': ')[1]
        ) {

            # Maintain Audio Device and Reset Parametric EQ Profile
            Out-File -FilePath ($appDict['dir']+'\config\config.txt')
            write_output -returnAfter $false -appDict $appDict -str ('ERROR: \config\'+($configName)+' does not exist. Parametric EQ profile unassigned.')
            write_output -returnAfter $false -appDict $appDict -str ("     ~ Click 'Refresh' from the tool menu.")

            # Narrate
            if (
                $global:narrator
            ) {
                $synth.Speak('No profile')
            }
        } else {

            # Switch Audio Device and Reset Parametric EQ Profile
            Set-AudioDevice -Index $deviceIndex
            Out-File -FilePath ($appDict['dir']+'\config\config.txt')
            write_output -returnAfter $false -appDict $appDict -str ('ERROR: \config\'+($configName)+' does not exist. '+($selectedDeviceName -split ': ')[1]+' assigned with no parametric EQ profile.')
            write_output -returnAfter $false -appDict $appDict -str ("     ~ Click 'Refresh' from the tool menu.")

            # Narrate
            if (
                $global:narrator
            ) {
                $synth.Speak($deviceUserName + 'with no profile.')
            }
        }

        # Update App Title
        $appTitle = $appDict['sysTrayAppName']+' - '+$deviceUserName+' (No Profile)'
    }

    # Update App Title
    if (
        ($appTitle).Length -gt 63
    ) {
        $appTitle = $appTitle.Substring(0, 58) + '...))'
    }
    $sysTrayApp.Text = $appTitle
}

function add_tool_strip_menu_item {
    param (
        $text,
        $font,
        $menuItem
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $text                   = [Str] Text of the new item to add to the system tray application
    $font                   = [System.Drawing.Font] Font of the new item to add to the system tray
                                application
    $menuItem               = [System.Windows.Forms.ToolStripMenutItem] System tray application menu
                                object to add the new menu item

    Description
    ----------------------------------------------------------------------------------------------------
    Adds a new system tray application menu item to $menuItem.
    #>

    # Add ToolStrip Sub-Menu Item
    $toolStripSubMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $toolStripSubMenuItem.Text = $text
    $toolStripSubMenuItem.Font = $font
    $packItem = $menuItem.DropDownItems.Add($toolStripSubMenuItem)

    return $toolStripSubMenuItem
}

function add_tool_strip_drop_down_sep {
    param (
        $subMenu
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $subMenu                = [System.Windows.Forms.ToolStripMenutItem] System tray application sub-menu
                                object to add the separator

    Description
    ----------------------------------------------------------------------------------------------------
    Adds a new system tray application separator to $subMenu.
    #>

    # Add ToolStrip Separator
    $toolStripSep = New-Object System.Windows.Forms.ToolStripSeparator
    $packSep = $subMenu.DropDownItems.Add($toolStripSep)
}

function add_tool_strip_sep {
    param (
        $menu
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $menu                   = [System.Windows.Forms.ToolStripMenutItem] System tray application menu
                                object to add the separator

    Description
    ----------------------------------------------------------------------------------------------------
    Adds a new system tray application separator to $subMenu.
    #>

    # Add ToolStrip Separator
    $toolStripSep = New-Object System.Windows.Forms.ToolStripSeparator
    $packSep = $menu.Items.Add($toolStripSep)
}

function open_web_address {
    param (
        $webAddress,
        $narration
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $webAddress             = [Str] URL
    $narration              = [Str] Text to narrate when opening $webAddress

    Description
    ----------------------------------------------------------------------------------------------------
    Opens the $webAddress in the system default browser.
    #>

    # Open Web Address
    Start-Process $webAddress

    # Narrate
    if (
        $global:narrator
    ) {
        $synth.Speak($narration)
    }
}

function build_settings {
    param (
        $sysTrayApp,
        $contextMenu,
        $appDict,
        $playbackDeviceLst,
        $equalizerProfileLst,
        $tab
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $sysTrayApp             = [System.Windows.Forms.NotifyIcon] System tray application object
    $contextMenu            = [System.Windows.Forms.ContextMenuStrip] System tray application menu object
    $appDict                = [Dict] Dictionary containing the parameters essential to the application
    $playbackDeviceLst      = [Dict] Dictionary containing the active Playback Devices
    $equalizerProfileLst    = [Dict] Dictionary containing the available equalizer profiles
    $tab                    = [Str] Indicates which settings tab to generate

    Description
    ----------------------------------------------------------------------------------------------------
    Builds the settings window form.
    #>

    # Define Window and Control Dimensions
    $measure = @($global:playbackDeviceLst.Count, $global:equalizerProfileLst.Count, 7) | Measure-Object -Maximum
    $maxItems = $measure.Maximum
    $windowLength = [Int](145 + ($maxItems * 32))
    $tabControlLength = $windowLength - 94
    $tableLength = [Int]($maxItems * 9.1667)

    # Build Settings Form
    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = 'Settings'
    $settingsForm.Size = New-Object System.Drawing.Size @(620, $windowLength)
    $settingsForm.StartPosition = 'CenterScreen'
    $settingsForm.AutoScroll = $true
    $settingsForm.ShowIcon = $false
    $settingsForm.ControlBox = $false

    # Build Tab Control
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point 10, 10
    $tabControl.Size = New-Object System.Drawing.Point 585, $tabControlLength
    $tabControl.Anchor = ('Left', 'Right', 'Top', 'Bottom')
    $settingsForm.Controls.Add($tabControl)

    if (
        ($tab -eq 'Playback Devices' -Or $tab -eq 'All')
    ) {
        # Build Playback Devices Tab
        $deviceTab = New-Object System.Windows.Forms.Tabpage
        $deviceTab.Text = 'Playback devices'
        $deviceTab.Margin = 4
        $tabControl.Controls.Add($deviceTab)

        # Pack Playback Device Instructions
        $deviceInstruction = New-Object System.Windows.Forms.Label
        $deviceInstruction.Text = 'Select a device type and default equalizer profile for each playback device.'
        $deviceInstruction.Location = New-Object System.Drawing.Point 10, 18
        $deviceInstruction.Size = New-Object System.Drawing.Point 565, 16
        $deviceTab.Controls.Add($deviceInstruction)

        # Pack Playback Device Table
        if (
            $global:playbackDeviceLst
        ) {

            # Pack Table
            $deviceTable = New-Object System.Windows.Forms.TableLayoutPanel
            $deviceTable.Location = New-Object System.Drawing.Point 0, 52
            $deviceTable.Size = New-Object System.Drawing.Point 200, $tableLength
            $deviceTable.Anchor = ('Left', 'Right', 'Top', 'Bottom')
            $deviceTable.ColumnCount = 3
            $deviceTable.RowCount = $maxItems
            $deviceTable.BackColor = 'White'
            $deviceTab.Controls.Add($deviceTable)

            $deviceHeader = New-Object System.Windows.Forms.Label
            $deviceHeader.Anchor = ('Left', 'Right', 'Top', 'Bottom')
            $deviceHeader.Margin = 3
            $deviceHeader.Padding = 2
            $deviceHeader.Text = 'Device'
            $deviceHeader.BorderStyle = '2'
            $deviceTable.Controls.Add($deviceHeader, 0, 0)

            $adapterHeader = New-Object System.Windows.Forms.Label
            $adapterHeader.Anchor = ('Left', 'Right', 'Top', 'Bottom')
            $adapterHeader.Margin = 3
            $adapterHeader.Padding = 2
            $adapterHeader.Text = 'Type'
            $adapterHeader.BorderStyle = '2'
            $deviceTable.Controls.Add($adapterHeader, 1, 0)

            $defaultHeader = New-Object System.Windows.Forms.Label
            $defaultHeader.Anchor = ('Left', 'Right', 'Top', 'Bottom')
            $defaultHeader.Margin = 3
            $defaultHeader.Padding = 2
            $defaultHeader.Text = 'Default'
            $defaultHeader.BorderStyle = '2'
            $deviceTable.Controls.Add($defaultHeader, 2, 0)

            # # Initialize Vertical Position
            $row = 1

            # Pack Audio Playback Device Table
            ForEach ($device in ($global:playbackDeviceLst.Keys | Sort-Object)) {

                # Pack Device Label
                $deviceLabel = New-Object System.Windows.Forms.Label
                $deviceLabel.Size = New-Object System.Drawing.Point 218, 20
                $deviceLabel.Margin = 3
                $deviceLabel.Padding = 2
                $deviceLabel.Text = $device
                $deviceLabel.Tag = '0'+$row
                $deviceTable.Controls.Add($deviceLabel, 0, $row)

                # Pack Settings for Valid Equalizer APO Playback Devices
                if (
                    ($global:playbackDeviceLst[$device]['ID'] -split '\.')[-1] -in (retrieve_apo_devices -appDict $appDict)
                ) {
                    # Pack Adapter Type Combobox
                    $adapterMenu = New-Object System.Windows.Forms.ComboBox
                    $adapterMenu.DropDownStyle = 'DropDownList'
                    $adapterMenu.Size = New-Object System.Drawing.Point 107, 20
                    $adapterMenu.Margin = 3
                    $adapterMenu.Padding = 2
                    $adapterMenu.DropDownWidth = 107
                    $adapterMenu.Items.AddRange(@('Headphone', 'Speaker', 'Line-Out'))
                    $adapterMenu.Text = $global:playbackDeviceLst[$device]['DeviceType']
                    $adapterMenu.Name = $device
                    $adapterMenu.Tag = '1'+$row
                    $deviceTable.Controls.Add($adapterMenu, 1, $row)
    
                    $adapterMenu.Add_SelectedIndexChanged(
                        {
    
                            # Update Equalizer Profile Drop-Down Selections
                            if (
                                $global:playbackDeviceLst[$this.Name]['DeviceType'] -ne $this.SelectedItem
                            ) {
                                ForEach($control in $deviceTable.Controls) {
                                    $col = [Int]($this.Tag.Substring(0, 1)) + 1
                                    $row = [Int]($this.Tag.Substring(1, 1))
                                    if (
                                        $control.Tag -eq ($col, $row -join '')
                                    ) {
                                        $control.Items.Clear()
                                        if (
                                            $this.SelectedItem -eq 'Line-Out'
                                        ) {
                                            $control.Items.AddRange(@($global:equalizerProfileLst.Keys | Sort-Object) + @('No Profile'))
                                        } else {
                                            $control.Items.AddRange(@(filter_dict_keys -dict $global:equalizerProfileLst -filterOn 'DeviceType' -value $this.SelectedItem) + @('No Profile'))
                                        }
                                        $control.Text = 'No Profile'
                                    }
                                }
                            }
    
                            # Update Settings Dictionary Objects
                            $global:playbackDeviceLst = update_settings_dict -dict $global:playbackDeviceLst -subDict $this.Name -key 'DeviceType' -value $this.SelectedItem
                        }
                    )
    
                    # Pack Default Profile ComboBox
                    if (
                        $global:equalizerProfileLst
                    ) {
                        
                        # Set Acceptable Equalizer Profile List
                        if (
                            $global:playbackDeviceLst[$device]['DeviceType'] -eq 'Headphone'
                        ) {
                            $comboBoxProfileLst = filter_dict_keys -dict $global:equalizerProfileLst -filterOn 'DeviceType' -value 'Headphone'
                        } elseif (
                            $global:playbackDeviceLst[$device]['DeviceType'] -eq 'Speaker' 
                        ) {
                            $comboBoxProfileLst = filter_dict_keys -dict $global:equalizerProfileLst -filterOn 'DeviceType' -value 'Speaker'
                        } else {
                            $comboBoxProfileLst = ($global:equalizerProfileLst.Keys | Sort-Object)
                        }
    
                        # Pack Equalizer Profile ComboBox
                        $defaultMenu = New-Object System.Windows.Forms.ComboBox
                        $defaultMenu.DropDownStyle = 'DropDownList'
                        $defaultMenu.Size = New-Object System.Drawing.Point 232, 20
                        $defaultMenu.Margin = 3
                        $defaultMenu.Padding = 2
                        $defaultMenu.DropDownWidth = 232
                        $defaultMenu.Items.AddRange(@($comboBoxProfileLst) + @('No Profile'))
                        $defaultMenu.Text = $global:playbackDeviceLst[$device]['DefaultProfile']
                        $defaultMenu.Name = $device
                        $defaultMenu.Tag = '2'+$row
                        $deviceTable.Controls.Add($defaultMenu, 2, $row)
    
                        $defaultMenu.Add_SelectedIndexChanged(
                            {
                                # Update Settings Dictionary Objects
                                $global:playbackDeviceLst = update_settings_dict -dict $global:playbackDeviceLst -subDict $this.Name -key 'DefaultProfile' -value $this.SelectedItem
                            }
                        )
                    } else{
    
                        # Pack No Equalizer Profiles Available
                        $noProfileLabel = New-Object System.Windows.Forms.Label
                        $noProfileLabel.Size = New-Object System.Drawing.Point 218, 20
                        $noProfileLabel.Margin = 3
                        $noProfileLabel.Padding = 2
                        $noProfileLabel.Text = 'No equalizer profiles available...'
                        $noProfileLabel.Tag = '2'+$row
                        $deviceTable.Controls.Add($noProfileLabel, 2, $row)
                    }
                } else {
                    # Pack Adapter Type Combobox
                    $adapterMenu = New-Object System.Windows.Forms.ComboBox
                    $adapterMenu.DropDownStyle = 'DropDownList'
                    $adapterMenu.Size = New-Object System.Drawing.Point 107, 20
                    $adapterMenu.Margin = 3
                    $adapterMenu.Padding = 2
                    $adapterMenu.DropDownWidth = 107
                    $adapterMenu.Items.AddRange(@('Headphone', 'Speaker', 'Line-Out'))
                    $adapterMenu.Text = $global:playbackDeviceLst[$device]['DeviceType']
                    $adapterMenu.Name = $device
                    $adapterMenu.Enabled = $false
                    $adapterMenu.Tag = '1'+$row
                    $deviceTable.Controls.Add($adapterMenu, 1, $row)

                    # Pack No Equalizer Profiles Available
                    $defaultMenu = New-Object System.Windows.Forms.ComboBox
                    $defaultMenu.DropDownStyle = 'DropDownList'
                    $defaultMenu.Size = New-Object System.Drawing.Point 232, 20
                    $defaultMenu.Margin = 3
                    $defaultMenu.Padding = 2
                    $defaultMenu.DropDownWidth = 232
                    $defaultMenu.Items.AddRange(@('EqualizerAPO not installed...'))
                    $defaultMenu.Text = 'EqualizerAPO not installed...'
                    $defaultMenu.Name = $device
                    $defaultMenu.Enabled = $false
                    $defaultMenu.Tag = '2'+$row
                    $deviceTable.Controls.Add($defaultMenu, 2, $row)  
                }

                # Increment Row
                $row = [Int]($row + 1)
            }
        } else {

            # Pack No Audio Playback Devices Available
            $noDeviceHeader = New-Object System.Windows.Forms.Label
            $noDeviceHeader.Location = New-Object System.Drawing.Point 20, 50
            $noDeviceHeader.Size = New-Object System.Drawing.Point 555, 16
            $noDeviceHeader.Text = 'No audio playback devices available...'
            $deviceTab.Controls.Add($noDeviceHeader)
        }
    }

    if (
        ($tab -eq 'Equalizer Profiles' -Or $tab -eq 'All')
    ) {
        # Build Equalizer Profiles Tab
        $profileTab = New-object System.Windows.Forms.Tabpage
        $profileTab.Text = 'Equalizer profiles'
        $profileTab.Margin = 4
        $tabControl.Controls.Add($profileTab)

        # Pack Profile Header
        $profileHeader = New-Object System.Windows.Forms.Label
        $profileHeader.Text = 'Select a device type for each equalizer profile.'
        $profileHeader.Location = New-Object System.Drawing.Point 10, 18
        $profileHeader.Size = New-Object System.Drawing.Point 565, 16
        $profileTab.Controls.Add($profileHeader)

        # Pack Equalizer Profile Table
        if (
            $global:equalizerProfileLst
        ) {
            # Pack Table
            $profileTable = New-Object System.Windows.Forms.TableLayoutPanel
            $profileTable.Location = New-Object System.Drawing.Point 0, 52
            $profileTable.Size = New-Object System.Drawing.Point 200, $tableLength
            $profileTable.Anchor = ('Left', 'Right', 'Top', 'Bottom')
            $profileTable.ColumnCount = 3
            $profileTable.RowCount = $maxItems
            $profileTable.BackColor = 'White'
            $profileTab.Controls.Add($profileTable)

            $profileHeader = New-Object System.Windows.Forms.Label
            $profileHeader.Anchor = ('Left', 'Right', 'Top', 'Bottom')
            $profileHeader.Margin = 3
            $profileHeader.Padding = 2
            $profileHeader.Text = 'Profile'
            $profileHeader.BorderStyle = '2'
            $profileTable.Controls.Add($profileHeader, 0, 0)

            $sourcetHeader = New-Object System.Windows.Forms.Label
            $sourcetHeader.Anchor = ('Left', 'Right', 'Top', 'Bottom')
            $sourcetHeader.Margin = 3
            $sourcetHeader.Padding = 2
            $sourcetHeader.Text = 'Source'
            $sourcetHeader.BorderStyle = '2'
            $profileTable.Controls.Add($sourcetHeader, 1, 0)

            $adapterHeader = New-Object System.Windows.Forms.Label
            $adapterHeader.Anchor = ('Left', 'Right', 'Top', 'Bottom')
            $adapterHeader.Margin = 3
            $adapterHeader.Padding = 2
            $adapterHeader.Text = 'Type'
            $adapterHeader.BorderStyle = '2'
            $profileTable.Controls.Add($adapterHeader, 2, 0)

            # Initialize Vertical Position
            $row = 1

            # Pack Equalizer Profile Table
            ForEach ($profile in ($global:equalizerProfileLst.Keys | Sort-Object)) {

                # Pack Profile Label
                $profileLabel = New-Object System.Windows.Forms.Label
                $profileLabel.Size = New-Object System.Drawing.Point 218, 20
                $profileLabel.Margin = 3
                $profileLabel.Padding = 2
                $profileLabel.Text = $profile
                $profileLabel.Tag = '0'+$row
                $profileTable.Controls.Add($profileLabel, 0, $row)

                # Pack Source Label
                $sourceLabel = New-Object System.Windows.Forms.Label
                $sourceLabel.Size = New-Object System.Drawing.Point 107, 20
                $sourceLabel.Margin = 3
                $sourceLabel.Padding = 2
                $sourceLabel.Text = $global:equalizerProfileLst[$profile]['source']
                $sourceLabel.Tag = '1'+$row
                $profileTable.Controls.Add($sourceLabel, 1, $row)

                # Pack Adapter Type Combobox
                if (
                    $global:equalizerProfileLst[$profile]['source'] -eq 'AutoEq'
                ) {
                    $adapterMenu = New-Object System.Windows.Forms.ComboBox
                    $adapterMenu.DropDownStyle = 'DropDownList'
                    $adapterMenu.Size = New-Object System.Drawing.Point 107, 20
                    $adapterMenu.Margin = 3
                    $adapterMenu.Padding = 2
                    $adapterMenu.DropDownWidth = 107
                    $adaptermenu.Items.AddRange(@($global:equalizerProfileLst[$profile]['DeviceType']))
                    $adapterMenu.Text = $global:equalizerProfileLst[$profile]['DeviceType']
                    $adapterMenu.Name = $profile
                    $adapterMenu.Enabled = $false
                    $adapterMenu.Tag = '2'+$row
                    $profileTable.Controls.Add($adaptermenu, 2, $row)
                } else {
                    $adapterMenu = New-Object System.Windows.Forms.ComboBox
                    $adapterMenu.DropDownStyle = 'DropDownList'
                    $adapterMenu.Size = New-Object System.Drawing.Point 107, 20
                    $adapterMenu.Margin = 3
                    $adapterMenu.Padding = 2
                    $adapterMenu.DropDownWidth = 107
                    $adaptermenu.Items.AddRange(@('Headphone', 'Speaker'))
                    $adapterMenu.Text = $global:equalizerProfileLst[$profile]['DeviceType']
                    $adaptermenu.Name = $profile
                    $adapterMenu.Tag = '2'+$row
                    $profileTable.Controls.Add($adaptermenu, 2, $row)

                    $adapterMenu.Add_SelectedIndexChanged(
                        {

                            # Update Settings Dictionary Objects
                            $global:equalizerProfileLst = update_settings_dict -dict $global:equalizerProfileLst -subDict $this.Name -key 'DeviceType' -value $this.SelectedItem

                            # Update Equalizer Profile Drop-Down Selections
                            ForEach($control in $deviceTable.Controls) {
                                if (
                                    $control.Tag -match '1[0-9]'
                                ) {
                                    if (
                                        ($control.SelectedItem -eq 'Headphone' -and $control.Enabled)
                                    ) {
                                        $col = [Int]($control.Tag.Substring(0, 1)) + 1
                                        $row = [Int]($control.Tag.Substring(1, 1))
                                        ForEach($defaultControl in $deviceTable.Controls) {
                                            if (
                                                $defaultControl.Tag -eq ($col, $row -join '')
                                            ) {
                                                $selectedItem = $defaultControl.SelectedItem
                                                $defaultControl.Items.Clear()
                                                $defaultControl.Items.AddRange(@(filter_dict_keys -dict $global:equalizerProfileLst -filterOn 'DeviceType' -value 'Headphone') + @('No Profile'))
                                                if (
                                                    $selectedItem -in (@(filter_dict_keys -dict $global:equalizerProfileLst -filterOn 'DeviceType' -value 'Headphone') + @('No Profile'))
                                                ) {
                                                    $defaultControl.Text = $selectedItem
                                                } else {
                                                    $defaultControl.Text = 'No Profile'
                                                }
                                            }
                                        }
                                    } elseif (
                                        ($control.SelectedItem -eq 'Speaker' -and $control.Enabled)
                                    ) {
                                        $col = [Int]($control.Tag.Substring(0, 1)) + 1
                                        $row = [Int]($control.Tag.Substring(1, 1))
                                        ForEach($defaultControl in $deviceTable.Controls) {
                                            if (
                                                $defaultControl.Tag -eq ($col, $row -join '')
                                            ) {
                                                $selectedItem = $defaultControl.SelectedItem
                                                $defaultControl.Items.Clear()
                                                $defaultControl.Items.AddRange(@(filter_dict_keys -dict $global:equalizerProfileLst -filterOn 'DeviceType' -value 'Speaker') + @('No Profile'))
                                                if (
                                                    $selectedItem -in (@(filter_dict_keys -dict $global:equalizerProfileLst -filterOn 'DeviceType' -value 'Speaker') + @('No Profile'))
                                                ) {
                                                    $defaultControl.Text = $selectedItem
                                                } else {
                                                    $defaultControl.Text = 'No Profile'
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    )
                }

                # Increment Row
                $row = [Int]($row + 1)
            }
        } else {

            # Pack No Equalizer Profiles Available
            $noProfileHeader = New-Object System.Windows.Forms.Label
            $noProfileHeader.Location = New-Object System.Drawing.Point 20, 50
            $noProfileHeader.Size = New-Object System.Drawing.Point 555, 16
            $noProfileHeader.Text = 'No equalizer profiles available...'
            $profileTab.Controls.Add($noProfileHeader)
        }
    }

    # Pack Save Button
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Location = New-Object System.Drawing.Point (515 - 90), ($windowLength - 74)
    $saveButton.Size = New-Object System.Drawing.Point 80, 24
    $saveButton.Margin = 3
    $saveButton.Padding = 2
    $saveButton.Text = 'Save'
    $settingsForm.Controls.Add($saveButton)

    $saveButton.add_Click(
        {
            # Close Settings Form
            $settingsForm.Close()

            # Push Equalizer Profile User-Inputs to Settings
            push_dict_to_reg -appDict $appDict -subKey 'Equalizer Profiles' -dict $global:equalizerProfileLst -properties @('DeviceType')

            # Push Playback Device User-Inputs to Settings
            push_dict_to_reg -appDict $appDict -subKey 'Playback Devices' -dict $global:playbackDeviceLst -properties @('DeviceType', 'DefaultProfile')

            # Re-Initialize
            init_system_tray_menu -sysTrayApp $sysTrayApp -contextMenu $contextMenu -appDict $appDict -restart $true
        }
    )

    # Pack Cancel Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point 515, ($windowLength - 74)
    $cancelButton.Size = New-Object System.Drawing.Point 80, 24
    $cancelButton.Margin = 3
    $cancelButton.Padding = 2
    $cancelButton.Text = 'Cancel'
    $settingsForm.Controls.Add($cancelButton)

    $cancelButton.add_Click(
        {
            # Close Settings Form
            $settingsForm.Close()
    
            # Reset Equalizer Profile Dictionary from Settings
            $global:equalizerProfileLst = apply_settings_to_dict -dict $global:equalizerProfileLst -settings (convert_reg_to_dict -appDict $appDict -subKey 'Equalizer Profiles')
            
            # Reset Playback Device Dictionary from Settings
            $global:playbackDeviceLst = apply_settings_to_dict -dict $global:playbackDeviceLst -settings (convert_reg_to_dict -appDict $appDict -subKey 'Playback Devices')
        }
    )

    # Display Output Form
    # $settingsForm.Topmost = $true
    $settingsForm.Control.ControlCollection
    $settingsForm.Add_Shown({$settingsForm.Activate()})
    [void] $settingsForm.ShowDialog()
}

function build_output {
    param (
        $appDict
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $appDict                = [Dict] Dictionary containing the parameters essential to the application

    Description
    ----------------------------------------------------------------------------------------------------
    Builds the output window form.
    #>

    # Retrieve output.txt
    $outputText = Get-Content -Path ($appDict['dir']+'\config\output.txt')
    $outputHeight = (Get-Content -Path ($appDict['dir']+'\config\output.txt')).Length
    $formHeight = [Int](($outputHeight * 18) + 60)
    if (
        $formHeight -gt [Int]750
    ) {
        $formHeight = [Int]750
    }

    # Build Output Form
    $outputForm = New-Object System.Windows.Forms.Form
    $outputForm.Text = 'Output'
    $outputForm.Size = New-Object System.Drawing.Size @(1250, $formHeight)
    $outputForm.StartPosition = 'CenterScreen'
    $outputForm.AutoScroll = $true
    $outputForm.Icon = create_icon -UnicodeChar 0xEA37 -size 186 -theme 0 -x -25 -y 0

    # Initialize Vertical Position
    $position = 0

    # Display output.txt
    $outputText.GetEnumerator() | ForEach-Object {

        # Pack Output Label
        $outputLabel = New-Object System.Windows.Forms.Label
        $outputLabel.Location = New-Object System.Drawing.Point 0, $position
        $outputLabel.Size = New-Object System.Drawing.Point 1250, 16
        $outputForm.controls.Add($outputLabel)
        if (
            ($_.ToLower().contains('ERROR:'.ToLower())) -Or ($_.ToLower().contains('~'.ToLower()))
            ) {
            $outputLabel.ForeColor = 'Red'
            $outputLabel.Font = $outputBody
        } elseif (
            ($_.ToLower().contains('WARNING:'.ToLower()))
        ) {
            $outputLabel.ForeColor = 'Orange'
            $outputLabel.Font = $outputBody
        } elseif (
            ($_.ToLower().contains('NOTE:'.ToLower())) -Or ($_.ToLower().contains('Date Time:'.ToLower())) -Or ($_.ToLower().contains('Version  :'.ToLower()))
        ) {
            $outputLabel.Font = $outputBody
        } else {
            $outputLabel.Font = $outputHeader
        }

        # Output Text
        $outputLabel.text = $_

        # Increment Position
        $position = [Int]($position + 18)
    }

    # Display Output Form
    $outputForm.Add_Shown({$outputForm.Activate()})
    [void] $outputForm.ShowDialog()
}

function init_system_tray_menu {
    param (
        $sysTrayApp,
        $contextMenu,
        $appDict,
        $restart
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $sysTrayApp             = [System.Windows.Forms.NotifyIcon] System tray application object
    $contextMenu            = [System.Windows.Forms.ContextMenuStrip] System tray application menu object
    $appDict                = [Dict] Dictionary containing the parameters essential to the application
    $restart                = [Str] Indicates whether to clear $contextMenu

    Description
    ----------------------------------------------------------------------------------------------------
    Initializes the system tray application.
    #>

    # Remove Existing output.txt
    if (
        [System.IO.File]::Exists($appDict['dir']+'\config\output.txt')
    ) {
        Remove-Item -Path ($appDict['dir']+'\config\output.txt')
    }

    # Set Validation Flags
    $appDict['versionMismatch'] = $false
    $appDict['connectionErr'] = $false
    $appDict['invalidConfigErr'] = $false
    $appDict['missingConfigErr'] = $false
    $appDict['profileErr'] = $false
    $appDict['apoErr'] = $true
    $appDict[ 'missingPlaybackDeviceErr'] = $false
    
    # Initialize Registry
    create_reg -appDict $appDict

    # Clear System Tray Sub-Menu Items on Refresh
    if (
        $restart
    ) {
        $contextMenu.Items.Clear();
    }

    # Initialzie Logging
    write_output -returnAfter $false -appDict $appDict -str ('+'+'-'*163+'+')
    write_output -returnAfter $false -appDict $appDict -str ('|'+' '*71+$appDict['sysTrayAppName']+' '*72+'|')
    write_output -returnAfter $false -appDict $appDict -str ('+'+'-'*163+'+')
    write_output -returnAfter $false -appDict $appDict -str ('    Date Time: '+(Get-Date -Format G))
    write_output -returnAfter $true -appDict $appDict -str ('    Version  : v'+$appDict['version'])

    # Validate Internet Connection
    $appDict = eval_connection -appDict $appDict -domain 'www.github.com'

    # Validate Version Mismatch
    $appDict = eval_version -appDict $appDict

    # Validate Equalizer Profiles
    validate_equalizer_profiles -sysTrayApp $sysTrayApp -contextMenu $contextMenu -appDict $appDict
}

function validate_equalizer_profiles {
    param (
        $sysTrayApp,
        $contextMenu,
        $appDict
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $sysTrayApp             = [System.Windows.Forms.NotifyIcon] System tray application object
    $contextMenu            = [System.Windows.Forms.ContextMenuStrip] System tray application menu object
    $appDict                = [Dict] Dictionary containing the parameters essential to the application

    Description
    ----------------------------------------------------------------------------------------------------
    Validates the available equalizer profiles.
    #>

    # Validate User-Provided '.json'
    $autoEqConf, $appDict = eval_auto_eq_conf -appDict $appDict

    # Retrieve List of Equalizer Profiles
    $global:equalizerProfileLst, $appDict = retrieve_eq_profiles -appDict $appDict -autoEqConf $autoEqConf

    # Validate Equalizer Profiles
    $newEqualizerProfiles = eval_dict_against_settings -dict $global:equalizerProfileLst -settings (convert_reg_to_dict -appDict $appDict -subKey 'Equalizer Profiles')
    if (
        $newEqualizerProfiles
    ) {

        # Apply Settings to any Equalizer Profile that Already Exists
        $global:equalizerProfileLst = apply_settings_to_dict -dict $global:equalizerProfileLst -settings (convert_reg_to_dict -appDict $appDict -subKey 'Equalizer Profiles')

        # Update Equalizer Profile Registry
        push_dict_to_reg -appDict $appDict -subKey 'Equalizer Profiles' -dict $global:equalizerProfileLst -properties @('Name', 'DeviceType', 'Source')

        # Raise Equalizer Profile Settings
        build_settings -sysTrayApp $sysTrayApp -contextMenu $contextMenu -appDict $appDict -playbackDeviceLst $false -equalizerProfileLst $global:equalizerProfileLst -tab 'Equalizer Profiles'

    } else {
        
        # Update Equalizer Profile Registry (Source Parameter Only)
        push_dict_to_reg -appDict $appDict -subKey 'Equalizer Profiles' -dict $global:equalizerProfileLst -properties @('Source')

        # Update Equalizer Profile Dictionary from Settings
        $global:equalizerProfileLst = apply_settings_to_dict -dict $global:equalizerProfileLst -settings (convert_reg_to_dict -appDict $appDict -subKey 'Equalizer Profiles')

        # Validate Playback Devices
        validate_playback_devices -sysTrayApp $sysTrayApp -contextMenu $contextMenu -appDict $appDict
    }
}

function validate_playback_devices {
    param (
        $sysTrayApp,
        $contextMenu,
        $appDict
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $sysTrayApp             = [System.Windows.Forms.NotifyIcon] System tray application object
    $contextMenu            = [System.Windows.Forms.ContextMenuStrip] System tray application menu object
    $appDict                = [Dict] Dictionary containing the parameters essential to the application

    Description
    ----------------------------------------------------------------------------------------------------
    Validates the active playback devices.
    #>

    # Retrieve List of Playback Devices
    $global:playbackDeviceLst, $appDict = retrieve_playback_devices -appDict $appDict

    # Validate Playback Devices
    $newPlaybackDevices = eval_dict_against_settings -dict $global:playbackDeviceLst -settings (convert_reg_to_dict -appDict $appDict -subKey 'Playback Devices')

    if (
        $newPlaybackDevices
    ) {

        # Apply Settings to any Playback Device that Already Exists
        $global:playbackDeviceLst = apply_settings_to_dict -dict $global:playbackDeviceLst -settings (convert_reg_to_dict -appDict $appDict -subKey 'Playback Devices')

        # Update Playback Device Registry
        push_dict_to_reg -appDict $appDict -subKey 'Playback Devices' -dict $global:playbackDeviceLst -properties @('Name', 'DeviceType', 'ID', 'Type', 'DefaultProfile')

        # Raise Playback Device Settings
        build_settings -sysTrayApp $sysTrayApp -contextMenu $contextMenu -appDict $appDict -playbackDeviceLst $global:playbackDeviceLst -equalizerProfileLst $global:equalizerProfileLst -tab 'Playback Devices'

    } else {

        # Update Playback Device Registry
        push_dict_to_reg -appDict $appDict -subKey 'Playback Devices' -dict $global:playbackDeviceLst -properties @('Name', 'ID')

        # Update Playback Device Dictionary from Settings
        $global:playbackDeviceLst = apply_settings_to_dict -dict $global:playbackDeviceLst -settings (convert_reg_to_dict -appDict $appDict -subKey 'Playback Devices')

        # Build System Tray Application
        build_system_tray_menu -sysTrayApp $sysTrayApp -contextMenu $contextMenu -appDict $appDict -equalizerProfileLst $global:equalizerProfileLst -playbackDeviceLst $global:playbackDeviceLst
    }
}

function build_system_tray_menu {
    param (
        $sysTrayApp,
        $contextMenu,
        $appDict,
        $equalizerProfileLst,
        $playbackDeviceLst
    )
    <#
    Variables
    ----------------------------------------------------------------------------------------------------
    $sysTrayApp             = [System.Windows.Forms.NotifyIcon] System tray application object
    $contextMenu            = [System.Windows.Forms.ContextMenuStrip] System tray application menu object
    $appDict                = [Dict] Dictionary containing the parameters essential to the application
    $playbackDeviceLst      = [Dict] Dictionary containing the active Playback Devices
    $equalizerProfileLst    = [Dict] Dictionary containing the available equalizer profiles

    Description
    ----------------------------------------------------------------------------------------------------
    Buids the system tray application.
    #>

    # Pack System Tray Audio Device Menu and EQ Profile Sub-Menu Items
    if (
        $global:playbackDeviceLst
    ) {
        ForEach ($device in ($global:playbackDeviceLst.Keys | Sort-Object)) {

            # Add System Tray Audio Device Menu Item
            $menuDevice = $contextMenu.Items.Add($device)
            $menuDevice.Name = $global:playbackDeviceLst[$device]['Name']
            $menuDevice.Tag = $global:playbackDeviceLst[$device]['DefaultProfile']
    
            # Pack System Tray EQ Profile Sub-Menu Items
            if (
                ($global:equalizerProfileLst)
            ) {
                if (
                    ($global:playbackDeviceLst[$device]['ID'] -split '\.')[-1] -in (retrieve_apo_devices -appDict $appDict)
                ) {

                    # Assign ToolTipText to Assign Default Profile for Audio Device
                    $menuDevice.ToolTipText = 'Apply '+$global:playbackDeviceLst[$device]['DefaultProfile']

                    # Assign Action to Assign Default Profile for Audio Device
                    if (
                        $global:playbackDeviceLst[$device]['DefaultProfile'] -eq 'No Profile'
                    ) {
                        $menuDevice.add_Click(
                            {
                                add_audio_device_click_action -sysTrayApp $sysTrayApp -appDict $appDict -selectedItem $this.Tag -parentItem $this.Name -reset $true
                            }
                        )
                    } else {
                        $menuDevice.add_Click(
                            {
                                add_audio_device_click_action -sysTrayApp $sysTrayApp -appDict $appDict -selectedItem $this.Tag -parentItem $this.Name -reset $false
                            }
                        )
                    }

                    # Set Acceptable Equalizer Profile List
                    if (
                        $global:playbackDeviceLst[$device]['DeviceType'] -eq 'Headphone'
                    ) {
                        $deviceTypeProfileLst = filter_dict_keys -dict $global:equalizerProfileLst -filterOn 'DeviceType' -value 'Headphone'
                    } elseif (
                        $global:playbackDeviceLst[$device]['DeviceType'] -eq 'Speaker' 
                    ) {
                        $deviceTypeProfileLst = filter_dict_keys -dict $global:equalizerProfileLst -filterOn 'DeviceType' -value 'Speaker'
                    } else {
                        $deviceTypeProfileLst = ($global:equalizerProfileLst.Keys | Sort-Object)
                    }
                    ForEach ($profile in ($deviceTypeProfileLst | Sort-Object)) {

                        # Add System Tray ToolStrip Sub-Menu Item
                        $eqProfile = add_tool_strip_menu_item -text $profile -font $fontReg -menuItem $menuDevice

                        # Add System Tray ToolStrip Sub-Menu Click
                        $eqProfile.add_Click(
                            {
                                add_audio_device_click_action -sysTrayApp $sysTrayApp -appDict $appDict -selectedItem $this -parentItem $this.OwnerItem.Name -reset $false
                            }
                        )
                    }
    
                    # Add System Tray ToolStrip Sub-Menu Separator
                    if (
                        ($deviceTypeProfileLst)
                    ) {
                        add_tool_strip_drop_down_sep -subMenu $menuDevice
                    }
    
                    # Add System Tray ToolStrip Sub-Menu Item for No Profile
                    $noProfile = add_tool_strip_menu_item -text 'No Profile' -font $fontReg -menuItem $menuDevice
    
                    # Assign Action for Audio Device with No Profile
                    $noProfile.add_Click(
                        {
                            add_audio_device_click_action -sysTrayApp $sysTrayApp -appDict $appDict -selectedItem $this -parentItem $this.OwnerItem.Name -reset $true
                        }
                    )

                } else {
    
                    # Assign ToolTipText for Audio Device with EqualizerAPO Not Installed
                    $menuDevice.ToolTipText = 'EqualizerAPO Not Installed'
    
                    # Assign Action for Audio Device with EqualizerAPO Not Installed
                    $menuDevice.add_Click(
                        {
                            add_audio_device_click_action -sysTrayApp $sysTrayApp -appDict $appDict -selectedItem $this.Name -parentItem $false -reset $true
                        }
                    )
                }
            } else {
                if (
                    ($global:playbackDeviceLst[$device]['ID'] -split '\.')[-1] -in (retrieve_apo_devices -appDict $appDict)
                ) {
    
                    # Assign ToolTipText for Audio Device when No EQ Profiles are Found
                    $menuDevice.ToolTipText = 'No Equalizer Profiles Available'
    
                    # Assign Action for Audio Device when No EQ Profiles are Found
                    $menuDevice.add_Click(
                        {
                            add_audio_device_click_action -sysTrayApp $sysTrayApp -appDict $appDict -selectedItem $this.Name -parentItem $false -reset $true
                        }
                    )
                } else {
    
                    # Assign ToolTipText for Audio Device with EqualizerAPO Not Installed
                    $menuDevice.ToolTipText = 'EqualizerAPO Not Installed'
    
                    # Assign Action for Audio Device with EqualizerAPO Not Installed
                    $menuDevice.add_Click(
                        {
                            add_audio_device_click_action -sysTrayApp $sysTrayApp -appDict $appDict -selectedItem $this.Name -parentItem $false -reset $true
                        }
                    )
                }
            }
        }
    } else {

        # Add System Tray Item for No Audio Devices
        $noMenuDevice = $contextMenu.Items.Add('No Playback Devices')

        # Apply System Tray Format
        $noMenuDevice.Font = $fontItalic
        $noMenuDevice.Enabled = $false
    }

    # Add System Tray ToolStrip Separator
    add_tool_strip_sep -menu $contextMenu

    # Add System Tray Settings Function
    $settingsTool = $contextMenu.Items.Add('Settings')
    $settingsTool.add_Click(
        {
            # Build Settings
            build_settings -sysTrayApp $sysTrayApp -contextMenu $contextMenu -appDict $appDict -equalizerProfileLst $global:equalizerProfileLst -playbackDeviceLst $global:playbackDeviceLst -tab 'All'
        }
    )

    # Add System Tray Output Function
    $outputTool = $contextMenu.Items.Add('Output')
    $outputTool.add_Click(
        {

            # Build Output
            build_output -appDict $appDict
        }
    )

    # Add System Tray Refresh Function
    $listDevices = $contextMenu.Items.Add('Refresh')
    $listDevices.add_Click(
        {

            # Narrate
            if (
                $global:narrator
            ) {
                $synth.Speak('Refresh')
            }

            # Refresh
            init_system_tray_menu -sysTrayApp $sysTrayApp -contextMenu $contextMenu -appDict $appDict -restart $true
        }
    )

    # Add System Tray ToolStrip Separator
    add_tool_strip_sep -menu $contextMenu

    # Add System Tray Open Configurator Function
    $configuratorTool = $contextMenu.Items.Add('Open Configurator')
    $configuratorTool.add_Click(
        {
            # Narrate
            if (
                $global:narrator
            ) {
                $synth.Speak('Open Configurator')
            }

            # Open Configurator
            $tempDir = $appDict['dir']
            Start-Process "$tempDir\Configurator.exe"
        }
    )

    # Add System Tray Open Editor Function
    $editorTool = $contextMenu.Items.Add('Open Editor')
    if (
        $global:equalizerProfileLst
    ) {

        # Add System Tray ToolStrip Sub-Menu Item for Active Config
        $eqProfile = add_tool_strip_menu_item -text 'Active Config' -font $fontReg -menuItem $editorTool

        # Add System Tray ToolStrip Sub-Menu Click for Active Config
        $eqProfile.add_Click(
            {
                # Narrate
                if (
                    $global:narrator
                ) {
                    $synth.Speak('Open Active Config')
                }

                # Open Editor with Active Config
                $tempDir = $appDict['dir']
                Set-Location $tempDir; .\Editor.exe "$tempDir\config\config.txt";
            }
        )

        # Add System Tray ToolStrip Sub-Menu Separator
        add_tool_strip_drop_down_sep -subMenu $editorTool

        ForEach ($profile in ($global:equalizerProfileLst.Keys | Sort-Object)) {

            # Add System Tray ToolStrip Sub-Menu Item
            $eqProfile = add_tool_strip_menu_item -text $profile -font $fontReg -menuItem $editorTool

            # Add System Tray ToolStrip Sub-Menu Click
            $eqProfile.add_Click(
                {
                    # Narrate
                    if (
                        $global:narrator
                    ) {
                        $synth.Speak('Open '+$this)
                    }

                    # Open Editor with Selected Config
                    $tempDir = $appDict['dir']
                    Set-Location $tempDir; .\Editor.exe "$tempDir\config\Parametric_EQ_$this.txt"
                }
            )
        }
    } else {

        # Add System Tray ToolStrip Sub-Menu Item when No EQ profiles are Found
        $eqProfile = add_tool_strip_menu_item -text 'No EQ Profiles Found' -font $fontReg -menuItem $editorTool
    }

    # Add System Tray ToolStrip Separator
    add_tool_strip_sep -menu $contextMenu

    # Add System Tray Narrator Setting
    $narratorTool = $contextMenu.Items.Add('Narrator')
    $narratorTool.add_Click(
        {
            if (
                -Not $global:narrator
            ) {

                # Turn Narrator On
                $global:narrator = $true
                $synth.Speak('Narrator On')

                # Check Narrator
                $this.Checked = $true
            } else {

                # Turn Narrator Off
                $global:narrator = $false
                $synth.Speak('Narrator Off')

                # UnCheck Narrator
                $this.Checked = $false
            }
        }
    )

    # Add System Tray Notifications Setting
    $notificationTool = $contextMenu.Items.Add('Notifications')
    $notificationTool.add_Click(
        {
            if (
                -Not $global:notifications
            ) {

                # Turn Notifications On
                $global:notifications = $true

                # Narrate
                if (
                    $global:narrator
                ) {
                    $synth.Speak('Notifications On')
                }

                # Check Notifications
                $this.Checked = $true
            } else {

                # Turn Notifications Off
                $global:notifications = $false

                # Narrate
                if (
                    $global:narrator
                ) {
                    $synth.Speak('Notifications Off')
                }

                # UnCheck Notifications
                $this.Checked = $false
            }
        }
    )

    # Add System Tray ToolStrip Separator
    add_tool_strip_sep -menu $contextMenu

    # Add System Tray Resources
    if (
        $appDict['versionMismatch']
    ) {
        $resourceSelection = $contextMenu.Items.Add('Resources - New Version Available')
    } else {
        $resourceSelection = $contextMenu.Items.Add('Resources')
    }
    
    # Add System Tray ToolStrip Sub-Menu Item Resource Link to Github Repository
    $githubLink = add_tool_strip_menu_item -text ('www.github.com/~/AutoEq-Device-Manager') -font $fontItalic -menuItem $resourceSelection
    $githubLink.add_Click(
        {
            open_web_address -webAddress ('https://github.com/thomaseleff/AutoEq-Device-Manager/releases/latest') -narration ('Open Auto E Q Device Manager Git hub')
        }
    )

    # Add System Tray ToolStrip Sub-Menu Item Resource Link to AutoEq Repository
    $autoEqLink = add_tool_strip_menu_item -text ('www.github.com/~/AutoEq/~/results') -font $fontItalic -menuItem $resourceSelection
    $autoEqLink.add_Click(
        {
            open_web_address -webAddress ('https://github.com/jaakkopasanen/AutoEq/tree/master/results') -narration ('Open Auto E Q Git hub')
        }
    )

    # Add System Tray ToolStrip Sub-Menu Item Resource Link to Equalizer-APO
    $eqAPOLink = add_tool_strip_menu_item -text ('www.sourceforge.net/~/equalizerapo') -font $fontItalic -menuItem $resourceSelection
    $eqAPOLink.add_Click(
        {
            open_web_address -webAddress ('https://sourceforge.net/projects/equalizerapo/') -narration ('Open Equalizer A P O Sourceforge')
        }
    )

    # Add System Tray ToolStrip Sub-Menu Item Resource Link to Peace Editor
    $peaceLink = add_tool_strip_menu_item -text ('www.sourceforge.net/~/peace-equalizer-apo-extension') -font $fontItalic -menuItem $resourceSelection
    $peaceLink.add_Click(
        {
            open_web_address -webAddress ('https://sourceforge.net/projects/peace-equalizer-apo-extension/') -narration ('Open Peace Sourceforge')
        }
    )

    # Add System Tray ToolStrip Separator
    add_tool_strip_sep -menu $contextMenu

    # Add System Tray Window Navigation
    $exitTool = $contextMenu.Items.Add('Exit')
    $exitTool.add_Click(
        {

            # Narrate
            if (
                $global:narrator
            ) {
                $synth.Speak('Exit')
            }

            # Close and Clean-up
            $sysTrayApp.Visible = $false
            $appContext.ExitThread()
            Stop-Process $pid
        }
    )

    # Add System Tray Default Settings
    if (
        $global:narrator
    ) {
        $narratorTool.Checked = $true
    }
    if (
        $global:notifications
    ) {
        $notificationTool.Checked = $true
    }

    # Apply System Tray Default Menu Format
    $outputTool.Font = $fontReg
    $listDevices.Font = $fontReg
    $configuratorTool.Font = $fontReg
    $editorTool.Font = $fontReg
    $narratorTool.Font = $fontReg
    $resourceSelection.Font = $fontReg
    $exitTool.Font = $fontBold

    # Raise Notifications
    if (
        $global:notifications
    ) {

        # Raise Version Mismatch Note
        if (
            $appDict['versionMismatch']
        ) {
            raise_notification -sysTrayApp $sysTrayApp -appDict $appDict -type Info -str 'A new version is available on [www.github.com].'
        }

        # Raise Internet Connection Error
        if (
            $appDict['connectionErr']
        ) {
            raise_notification -sysTrayApp $sysTrayApp -appDict $appDict -type Error -str 'No Internet connection to [www.github.com]. Connect to the Internet or resolve connection issues.'
        }

        # Raise Configuration Errors
        if (
            $appDict['missingPlaybackDeviceErr']
        ) {
            raise_notification -sysTrayApp $sysTrayApp -appDict $appDict -type Error -str 'No connected Audio Device(s) found. Ensure at least one Audio Device is enabled within the system device settings.'
        } elseif (
            $appDict['apoErr']
        ) {
            raise_notification -sysTrayApp $sysTrayApp -appDict $appDict -type Warning -str "EqualizerAPO is not installed for any connected Audio Devices. Check the 'Output' for more information."
        } elseif (
            $appDict['invalidConfigErr']
        ) {
            raise_notification -sysTrayApp $sysTrayApp -appDict $appDict -type Error -str "autoeq_config.json is not a valid '.json' file. Check the 'Output' for more information."
        } elseif (
            $appDict['profileErr']
        ) {
            raise_notification -sysTrayApp $sysTrayApp -appDict $appDict -type Error -str "Error(s) found in autoeq_config.json. Check the 'Output' for more information."
        } elseif (
            $appDict['missingConfigErr']
        ) {
            raise_notification -sysTrayApp $sysTrayApp -appDict $appDict -type Error -str "autoeq_config.json not found within the \config folder. Check the 'Output' for more information."
        } else {
            raise_notification -sysTrayApp $sysTrayApp -appDict $appDict -type Info -str 'Audio Device menu generated successfully.'
        }
    }
}

# Build System Tray Icon Object
$sysTrayApp = New-Object System.Windows.Forms.NotifyIcon
$sysTrayApp.Text = $appDict['sysTrayAppName']
$sysTrayApp.Icon = create_icon -UnicodeChar 0xE9E9 -size 186 -theme $appDict['windowsTheme'] -x -25 -y 5
$sysTrayApp.Visible = $true

# Build System Tray Menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
init_system_tray_menu -sysTrayApp $sysTrayApp -contextMenu $contextMenu -appDict $appDict -restart $false
$sysTrayApp.ContextMenuStrip = $contextMenu;

# Hide PowerShell Window in the System Tray Tool
$windowCode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncWindow = Add-Type -MemberDefinition $windowCode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $asyncWindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

# Manage RAM Usage
[System.GC]::Collect()

# Initialize Application Context and Run
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)
