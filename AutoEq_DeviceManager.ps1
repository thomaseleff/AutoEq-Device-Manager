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
Set-Variable -Name 'version' -Value (New-Object System.Version('1.5.0'))
Set-Variable -Name 'sysTrayAppName' -Value 'AutoEq Device Manager'
Set-Variable -Name 'dir' -Value (split-path $MyInvocation.MyCommand.Path -Parent)
Set-Variable -Name 'fontBold' -Value ([System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold))
Set-Variable -Name 'fontReg' -Value ([System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Regular))
Set-Variable -Name 'fontLink' -Value ([System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Italic))
Set-Variable -Name 'outputHeader' -Value ([System.Drawing.Font]::new('Courier New', 9, [System.Drawing.FontStyle]::Bold))
Set-Variable -Name 'outputBody' -Value ([System.Drawing.Font]::new('Courier New', 9, [System.Drawing.FontStyle]::Regular))
Set-Variable -Name 'windowsTheme' -Value ((Get-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\ -Name "SystemUsesLightTheme").SystemUsesLightTheme)
Set-Variable -Name 'appTheme' -Value ((Get-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\ -Name "AppsUseLightTheme").AppsUseLightTheme)

# Create Text to Speech Synth Object
Set-Variable -Name 'synth' -Value (New-Object System.Speech.Synthesis.SpeechSynthesizer)
$synth.Rate = -2
$global:narrator = $false

# Enable Notifications
$global:notifications = $true

# Define Functions

function create_icon {
    param (
        $unicodeChar,
        $size,
        $theme,
        $x,
        $y
    )

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
    $bitmap = New-Object System.Drawing.Bitmap 128,128
    $bitmapGraphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $bitmapGraphics.DrawString([char]($unicodeChar), $fontIcon, $brush, $x, $y)
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
        $returnAfter,
        $dir,
        $str
    )

    # Write Str to Console
    # Write-Host ($str)

    # Write Str to output.txt
    if (
        [System.IO.File]::Exists($dir+'\config\output.txt') -eq $true
    ) {
        if (
            $returnAfter -eq $true
        ) {
            Out-File -FilePath ($dir+'\config\output.txt') -InputObject $str -Encoding utf8 -Append
            Out-File -FilePath ($dir+'\config\output.txt') -InputObject '' -Encoding utf8 -Append
        } else {
            Out-File -FilePath ($dir+'\config\output.txt') -InputObject $str -Encoding utf8 -Append
        }
    } else {
        if (
            $returnAfter -eq $true
        ) {
            Out-File -FilePath ($dir+'\config\output.txt') -InputObject $str -Encoding utf8
            Out-File -FilePath ($dir+'\config\output.txt') -InputObject '' -Encoding utf8 -Append
        } else {
            Out-File -FilePath ($dir+'\config\output.txt') -InputObject $str -Encoding utf8
        }
    }
}

function raise_notification {
    param (
        $sysTrayApp,
        $type,
        $str
    )

    # Raise Notification
    $sysTrayApp.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::$type
    $sysTrayApp.BalloonTipTitle = $sysTrayAppName
    $sysTrayApp.BalloonTipText = $str
    $sysTrayApp.ShowBalloonTip(500)
}

function list_devices {
    param (
        $dir,
        $sysTrayApp,
        $contextMenu,
        $restart
    )

    # Set Local Variables
    Set-Variable -Name 'index' -Value 0
    Set-Variable -Name 'deviceErr' -Value $false
    Set-Variable -Name 'versionMismatch' -Value $false
    Set-Variable -Name 'connectionErr' -Value $false
    Set-Variable -Name 'jsonErr' -Value $false
    Set-Variable -Name 'profileWarn' -Value $false
    Set-Variable -Name 'profileErr' -Value $false
    Set-Variable -Name 'addProfiles' -Value $false
    Set-Variable -Name 'validProfiles' -Value @()

    # Create Hash Table
    $table = $null
    $table = @{}

    # Remove Existing output.txt
    if (
        [System.IO.File]::Exists($dir+'\config\output.txt') -eq $true
    ) {
        Remove-Item -Path ($dir+'\config\output.txt')
    }

    # Initialzie Logging
    write_output -returnAfter $false -dir $dir -str ('+-------------------------------------------------------------------------------------------------------------------------------------------------------------------+')
    write_output -returnAfter $false -dir $dir -str ('|                                                                       '+$sysTrayAppName+'                                                                       |')
    write_output -returnAfter $false -dir $dir -str ('+-------------------------------------------------------------------------------------------------------------------------------------------------------------------+')
    write_output -returnAfter $false -dir $dir -str ('    Date Time: '+(Get-Date -Format G))
    write_output -returnAfter $true -dir $dir -str ('    Version  : v'+$version)

    # Validate Internet Connection and Release Version
    if (
        (Test-Connection -ComputerName www.github.com -Quiet) -eq $true
    ){
        
        # Get Latest Release Version in Two Ways Since PowerShell HTML Parsing is Broken
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        try {
            $request = Invoke-WebRequest -Uri 'https://github.com/thomaseleff/AutoEq-Device-Manager/releases/latest' -UseBasicParsing
            $html = New-Object -Com "HTMLFile"
            [string]$htmlBody = $request.Content
            $html.write([ref]$htmlBody)
        } catch {
            $request = Invoke-WebRequest -Uri 'https://github.com/thomaseleff/AutoEq-Device-Manager/releases/latest'
            $html = $request.ParsedHtml
        }
        $versionLatest = $html.getElementsByClassName('ml-1') | Where-Object{$_.textContent -match '[a-zA-Z]{1}\d{1,2}\.\d{1,2}\.\d{1,3}'}
        $versionLatest = $versionLatest.textContent -Replace "[^0-9.]"
        $versionLatest = New-Object System.Version($versionLatest)

        # Logging
        if (
            ($version.CompareTo($versionLatest)) -lt 0
        ) {
            # Capture Version Mismatch for Notification
            $versionMismatch = $true
            write_output -returnAfter $true -dir $dir -str ('NOTE: A new version, v'+$versionLatest+', is available on [www.github.com].')
        }

    } else {

        # Capture Connection Error for Notification
        $connectionErr = $true
        write_output -returnAfter $true -dir $dir -str ('ERROR: No Internet connection to [www.github.com]. Connect to the Internet or resolve connection issues.')
    }

    # Validate Json
    if (
        [System.IO.File]::Exists($dir+'\config\eq_profiles.json') -eq $true
    ) {

        # Logging
        write_output -returnAfter $true -dir $dir -str ('NOTE: eq_profiles.json exists within the \config folder as expected.')

        try {
            $json = Get-Content -Raw -Path ($dir+'\config\eq_profiles.json') | ConvertFrom-Json
            $json = $json.GetEnumerator() | Sort-Object device
            # $json | ConvertTo-Json | Out-File ($dir+'\config\eq_profiles.json')
        } catch {

            # Capture Json Error for Notification
            $jsonErr = $true
            write_output -returnAfter $false -dir $dir -str ('ERROR: eq_profiles.json is not a valid json file.')

            # Replace eq_profiles.json
            if (
                $connectionErr -eq $false
            ) {

                # Backup Existing eq_profiles.json
                Copy-Item -Path ($dir+'\config\eq_profiles.json') -Destination ($dir+'\config\eq_profiles_backup_'+(Get-Date -Format 'MM_dd_yyyy_HH_mm_ss')+'.json')

                # Download Latest Valid eq_profiles.json
                Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/thomaseleff/AutoEq-Device-Manager/main/config/eq_profiles.json' -OutFile ($dir+'\config\eq_profiles.json')

                # Logging
                write_output -returnAfter $false -dir $dir -str ('     ~ The current eq_profiles.json has been backed up as eq_profiles_backup_'+(Get-Date -Format 'MM_dd_yyyy_HH_MM_SS')+'.json within the \config folder.')
                write_output -returnAfter $false -dir $dir -str ('     ~ The latest eq_profiles.json has been downloaded from [www.github.com].')
                write_output -returnAfter $true -dir $dir -str ("     ~ Modify the new template eq_profiles.json and then click 'Refresh Devices' from the tool menu.")
            }
        }
    } else {

        # Capture Profile Warning for Notification
        $profileWarn = $true
        write_output -returnAfter $false -dir $dir -str ('ERROR: eq_profiles.json not found within the \config folder.')

        # Download eq_profiles.json
        if (
            $connectionErr -eq $false
        ) {

            # Download Latest Valid eq_profiles.json
            Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/thomaseleff/AutoEq-Device-Manager/main/config/eq_profiles.json' -OutFile ($dir+'\config\eq_profiles.json')

            # Logging
            write_output -returnAfter $false -dir $dir -str ('     ~ The latest eq_profiles.json has been downloaded from [www.github.com].')
            write_output -returnAfter $true -dir $dir -str ("     ~ Modify the new template eq_profiles.json and then click 'Refresh Devices' from the tool menu.")
        }
    }

    # Retrive List of Parametric EQ Profiles
    if (
        ($connectionErr -eq $false -And $jsonErr -eq $false -And $profileWarn -eq $false)
    ) {

        # Retrieve
        $json | ForEach-Object {
            if (
                ('device' -notin $_.PSobject.Properties.Name -And 'parametricConfig' -notin $_.PSobject.Properties.Name)
            ) {

                # Capture Profile Error for Notification
                if (
                    $profileErr -eq $false
                ) {
                    $profileErr = $true
                }
                write_output -returnAfter $false -dir $dir -str ('Invalid EQ Profile Entry')
                write_output -returnAfter $false -dir $dir -str ('+----------------------+')
                write_output -returnAfter $false -dir $dir -str ('    ERROR: Invalid EQ Profile entry in eq_profiles.json.')
                write_output -returnAfter $true -dir $dir -str ('         ~ [device] and [parametricConfig] parameters not found.')
            } elseif (
                'device' -notin $_.PSobject.Properties.Name
            ) {

                # Capture Profile Error for Notification
                if (
                    $profileErr -eq $false
                ) {
                    $profileErr = $true
                }
                write_output -returnAfter $false -dir $dir -str ('Invalid EQ Profile Entry')
                write_output -returnAfter $false -dir $dir -str ('+----------------------+')
                write_output -returnAfter $false -dir $dir -str ('    ERROR: Invalid EQ Profile entry in eq_profiles.json.')
                write_output -returnAfter $true -dir $dir -str ('         ~ [device] parameter not found.')
            } elseif (
                'parametricConfig' -notin $_.PSobject.Properties.Name
            ) {

                # Capture Profile Error for Notification
                if (
                    $profileErr -eq $false
                ) {
                    $profileErr = $true
                }
                if (
                    ([string]::IsNullOrWhiteSpace($_.device)) 
                ) {
                    write_output -returnAfter $false -dir $dir -str ('Invalid EQ Profile Entry')
                    write_output -returnAfter $false -dir $dir -str ('+----------------------+')
                    write_output -returnAfter $false -dir $dir -str ('    ERROR: Invalid EQ Profile entry in eq_profiles.json.')
                    write_output -returnAfter $false -dir $dir -str ('         ~ [device] parameter exists but is null, empty or contains only white space.')
                    write_output -returnAfter $true -dir $dir -str ('         ~ [parametricConfig] parameter not found.')
                } else {
                    write_output -returnAfter $false -dir $dir -str ($_.device)
                    write_output -returnAfter $false -dir $dir -str ('+'+('-' * ($_.device.Length-2)+'+'))
                    write_output -returnAfter $false -dir $dir -str ('    ERROR: Invalid EQ Profile entry in eq_profiles.json.')
                    write_output -returnAfter $true -dir $dir -str ('         ~ [parametricConfig] parameter not found.')
                }

            } elseif (
                ([string]::IsNullOrWhiteSpace($_.device))
            ) {
                
                # Capture Profile Error for Notification
                if (
                    $profileErr -eq $false
                ) {
                    $profileErr = $true
                }
                write_output -returnAfter $false -dir $dir -str ('Invalid EQ Profile Entry')
                write_output -returnAfter $false -dir $dir -str ('+----------------------+')
                write_output -returnAfter $false -dir $dir -str ('    ERROR: Invalid EQ Profile entry in eq_profiles.json.')
                write_output -returnAfter $true -dir $dir -str ('         ~ [device] parameter exists but is null, empty or contains only white space.')
            } else {

                # Logging
                write_output -returnAfter $false -dir $dir -str ($_.device)
                write_output -returnAfter $false -dir $dir -str ('+'+('-' * ($_.device.Length-2)+'+'))
                if (
                    ($_.parametricConfig.ToLower().Contains('https://raw.githubusercontent.com/jaakkopasanen/AutoEq/master/results/'.ToLower()) -And $_.parametricConfig.ToLower().Contains('ParametricEQ.txt'.ToLower()))
                ) {
                    if (
                        [System.IO.File]::Exists($dir+'\config\Parametric_EQ_'+$_.device+'.txt') -eq $false
                    ) {
                        try {
                            Invoke-WebRequest -Uri $_.parametricConfig -OutFile ($dir+'\config\Parametric_EQ_'+$_.device+'.txt')
    
                            # Logging
                            write_output -returnAfter $false -dir $dir -str ('    NOTE: URL to [parametricConfig] is valid.')
                            write_output -returnAfter $true -dir $dir -str ('    NOTE: Parametric EQ Config retrieved successfully.')
    
                            # Add Profile to Valid List
                            $validProfiles += $_.device
                        } catch {
    
                            # Capture Profile Error for Notification
                            if (
                                $profileErr -eq $false
                            ) {
                                $profileErr = $true
                            }
                            write_output -returnAfter $true -dir $dir -str ('    ERROR: URL to [parametricConfig] not found.')
                        }     
                    } else {
                        write_output -returnAfter $true -dir $dir -str ('    NOTE: Passing, Parametric EQ Config already exists within the \config folder.')
    
                        # Add Profile to Valid List
                        $validProfiles += $_.device
                    }
                } else {
    
                    # Capture Profile Error for Notification
                    if (
                        $profileErr -eq $false
                    ) {
                        $profileErr = $true
                    }
                    write_output -returnAfter $false -dir $dir -str ('    ERROR: URL to [parametricConfig] is invalid or is null, empty or contains only white space.')
                    write_output -returnAfter $false -dir $dir -str ('         ~ The URL must begin with [https://raw.githubusercontent.com/] and link to a')
                    write_output -returnAfter $true -dir $dir -str ('         ~ Parametric EQ profile within the [jaakkopasanen/AutoEQ] Github project.')
                }
            }
        }

        # Set Flag
        if (
            $validProfiles.count -gt 0
        ) {
            $addProfiles = $true
        }
    }

    # Clear System Tray Sub-Menu Items
    if (
        $restart -eq $true
    ) {
        $contextMenu.Items.Clear();
    }

    while (

        # Use for Debugging
        # $index -lt 6

        # Use to End Loop
        $deviceErr -eq $false
    ) {
        $index = $index + 1;
        # write_output -returnAfter $true -dir $dir -str $deviceErr
        # write_output -returnAfter $true -dir $dir -str $index

        # Check If an Audio Device is Found at Each Index
        try {
            Set-Variable -Name 'audioDevice' -Value (Get-AudioDevice -Index $index)
        } catch {
            Set-Variable -Name 'returnedVal' -Value $_
            Set-Variable -Name 'returnedVal' -Value ($returnedVal -Replace '\s','')
            if (
                $returnedVal -eq 'NoAudioDevicewiththatIndex'
            ) {
                Set-Variable -Name 'deviceErr' -Value $true
            } else {
                continue
            }
        }

        # Record Audio Device Parameters for All Playback Devices
        if (
            $audioDevice.Type -eq 'Playback'
        ) {
            $tableAdd = $null
            $tableAdd = @{
                'Index' = $index
                'Name' = $audioDevice.Name
                'Default' = $audioDevice.Default
                'Type' = $audioDevice.Type
                'ID' = $audioDevice.ID
            }

            $table.Add($audioDevice.Name, $tableAdd)
            # $table[$index]

            # Add System Tray Menu Item
            Set-Variable -Name 'audioDeviceName' -Value $audioDevice.Name
            Set-Variable -Name 'deviceName' -Value ($index, $audioDeviceName -join ': ')
            $menuDevice = $contextMenu.Items.Add($deviceName)

            # Pack System Tray Sub-Menu Items
            if (
                $addProfiles -eq $true
            ) {
                $validProfiles | ForEach-Object {

                    # Add System Tray Sub-Menu Items
                    $eqProfile = New-Object System.Windows.Forms.ToolStripMenuItem
                    $eqProfile.Text = $_
                    $eqProfile.Font = $fontReg
                    $eqProfileNew = $menuDevice.DropDownItems.Add($eqProfile);

                    # Add System Tray Sub-Menu Click
                    $eqProfile.add_Click(
                        {

                            # Check Active Audio Device
                            Set-Variable -Name 'activeDevice' -Value (Get-AudioDevice -Playback)
                            Set-Variable -Name 'activeDeviceName' -Value $activeDevice.Name
                            Set-Variable -Name 'selectedDeviceName' -Value ($this.OwnerItem)
                            Set-Variable -Name 'deviceIndex' -Value (($selectedDeviceName -split ': ')[0])
                            Set-Variable -Name 'deviceUserName' -Value $selectedDeviceName
                            if (
                                [System.IO.File]::Exists($dir+'\config\Parametric_EQ_'+$this+'.txt') -eq $true
                            ) {
                                if (
                                    $activeDeviceName -eq ($selectedDeviceName -split ': ')[1]
                                ) {

                                    # Switch Parametric EQ Profile
                                    Copy-Item -Path ($dir+'\config\Parametric_EQ_'+$this+'.txt') -Destination ($dir+'\config\config.txt')
                                    write_output -returnAfter $false -dir $dir -str ('NOTE: '+$this+' parametric EQ profile successfully assigned.')

                                    # Narrate
                                    if (
                                        $global:narrator -eq $true
                                    ) {
                                        $synth.Speak($this)
                                    }
                                } else {

                                    # Switch Devices and Parametric EQ Profile
                                    Set-AudioDevice -Index $deviceIndex
                                    Copy-Item -Path ($dir+'\config\Parametric_EQ_'+$this+'.txt') -Destination ($dir+'\config\config.txt')
                                    write_output -returnAfter $false -dir $dir -str ('NOTE: '+($selectedDeviceName -split ': ')[1]+' successfully assigned with the parametric EQ profile for '+$this+'.')

                                    # Narrate
                                    if (
                                        $global:narrator -eq $true
                                    ) {
                                        $synth.Speak($deviceUserName + 'with' + $this)
                                    }
                                }
                            }

                            # Update App Title
                            $sysTrayApp.Text = $sysTrayAppName+' - '+$deviceUserName+' ('+$this+')'
                        }
                    )
                }

                # Add System Try Sub-Menu Item for No Profile
                $profileSepObj = New-Object System.Windows.Forms.ToolStripSeparator
                $profileSep = $menuDevice.DropDownItems.Add($profileSepObj);
                $eqProfile = New-Object System.Windows.Forms.ToolStripMenuItem
                $eqProfile.Text = 'No Profile'
                $eqProfile.Font = $fontReg
                $eqProfileNew = $menuDevice.DropDownItems.Add($eqProfile);

                # Add System Tray Sub-Menu Click for No Profile
                $eqProfile.add_Click(
                    {

                        # Check Active Audio Device
                        Set-Variable -Name 'activeDevice' -Value (Get-AudioDevice -Playback)
                        Set-Variable -Name 'activeDeviceName' -Value $activeDevice.Name
                        Set-Variable -Name 'selectedDeviceName' -Value ($this.OwnerItem)
                        Set-Variable -Name 'deviceIndex' -Value (($selectedDeviceName -split ': ')[0])
                        Set-Variable -Name 'deviceUserName' -Value $selectedDeviceName

                        if (
                            [System.IO.File]::Exists($dir+'\config\config.txt') -eq $true
                        ) {
                            if (
                                $activeDeviceName -eq ($selectedDeviceName -split ': ')[1]
                            ) {

                                # Remove the Parametric EQ Profile
                                Out-File -FilePath ($dir+'\config\config.txt')
                                write_output -returnAfter $false -dir $dir -str ('NOTE: Parametric EQ profile successfully unassigned.')

                                # Narrate
                                if (
                                    $global:narrator -eq $true
                                ) {
                                    $synth.Speak('No profile')
                                }
                            } else {

                                # Switch Devices and Remove the Parametric EQ Profile
                                Set-AudioDevice -Index $deviceIndex
                                Out-File -FilePath ($dir+'\config\config.txt')
                                write_output -returnAfter $false -dir $dir -str ('NOTE: '+($selectedDeviceName -split ': ')[1]+' successfully assigned with no parametric EQ profile.')

                                # Narrate
                                if (
                                    $global:narrator -eq $true
                                ) {
                                    $synth.Speak($deviceUserName + 'with no profile.')
                                }
                            }
                        } else {

                            if (
                                $activeDeviceName -eq ($selectedDeviceName -split ': ')[1]
                            ) {

                                # Remove the Parametric EQ Profile
                                write_output -returnAfter $false -dir $dir -str ('NOTE: Parametric EQ profile successfully unassigned.')

                                # Narrate
                                if (
                                    $global:narrator -eq $true
                                ) {
                                    $synth.Speak('No profile')
                                }
                            } else {

                                # Switch Devices and Remove the Parametric EQ Profile
                                Set-AudioDevice -Index $deviceIndex
                                write_output -returnAfter $false -dir $dir -str ('NOTE: '+($selectedDeviceName -split ': ')[1]+' successfully assigned with no parametric EQ profile.')

                                # Narrate
                                if (
                                    $global:narrator -eq $true
                                ) {
                                    $synth.Speak($deviceUserName + 'with no profile.')
                                }
                            }
                        }

                        # Update App Title
                        $sysTrayApp.Text = $sysTrayAppName+' - '+$deviceUserName+' ('+$this+')'
                    }
                )
            } else {

                # Add System Tray Sub-Menu Item for No EQ Profiles Found
                $eqProfile = New-Object System.Windows.Forms.ToolStripMenuItem
                $eqProfile.Text = 'No EQ Profiles Found'
                $eqProfile.Font = $fontReg
                $eqProfileNew = $menuDevice.DropDownItems.Add($eqProfile);

            } 
        }
    }

    # Raise Notification
    if (
        $global:notifications -eq $true
    ) {
        if (
            $connectionErr -eq $true
        ) {
            raise_notification -sysTrayApp $sysTrayApp -type Error -str 'ERROR: No Internet connection to [www.github.com]. Connect to the Internet or resolve connection issues.'
        } elseif (
            $jsonErr -eq $true
        ) {
            raise_notification -sysTrayApp $sysTrayApp -type Error -str "ERROR: eq_profiles.json is not a valid .json file. Check the 'Output' for more information."
        } elseif (
            $profileErr -eq $true
        ) {
            raise_notification -sysTrayApp $sysTrayApp -type Error -str "ERROR: Error(s) found in eq_profiles.json. Check the 'Output' for more information."
        } elseif (
            $profileWarn -eq $true
        ) {
            raise_notification -sysTrayApp $sysTrayApp -type Error -str "ERROR: eq_profiles.json not found within the \config folder. Check the 'Output' for more information."
        } else {
            if (
                $restart -eq $true
            ) {
                raise_notification -sysTrayApp $sysTrayApp -type Info -str 'NOTE: Device lists re-generated successfully.'
            } else {
                raise_notification -sysTrayApp $sysTrayApp -type Info -str 'NOTE: Device lists generated successfully.'
            }
        }
        if (
            $versionMismatch -eq $true
        ) {
            raise_notification -sysTrayApp $sysTrayApp -type Info -str 'NOTE: A new version is available on [www.github.com].'
        } 
    }

    $deviceSepObj = New-Object System.Windows.Forms.ToolStripSeparator
    $deviceSep = $contextMenu.Items.Add($deviceSepObj);

    # Add Functions
    $outputTool = $contextMenu.Items.Add('Output');
    $listDevices = $contextMenu.Items.Add('Refresh Devices');
    $toolSepObj = New-Object System.Windows.Forms.ToolStripSeparator
    $toolSep = $contextMenu.Items.Add($toolSepObj);

    # Add EqualizerAPO-Related Functions
    $configuratorTool = $contextMenu.Items.Add('Open Configurator');
    $editorTool = $contextMenu.Items.Add('Open Editor');

    # Pack Editor Sub-Menu Items
    if (
        $addProfiles -eq $true
    ) {

        # Add System Try Sub-Menu Item for Current Active Config
        $eqProfile = New-Object System.Windows.Forms.ToolStripMenuItem
        $eqProfile.Text = 'Current Active Config'
        $eqProfile.Font = $fontReg
        $eqProfileNew = $editorTool.DropDownItems.Add($eqProfile);
        $profileSepObj = New-Object System.Windows.Forms.ToolStripSeparator
        $profileSep = $editorTool.DropDownItems.Add($profileSepObj);

        # Add System Tray Sub-Menu Click for Current Active Config
        $eqProfile.add_Click(
            {
                cd $dir; .\Editor.exe "$dir\config\config.txt";
            }
        )

        $validProfiles | ForEach-Object {

            # Add System Tray Sub-Menu Items
            $eqProfile = New-Object System.Windows.Forms.ToolStripMenuItem
            $eqProfile.Text = $_
            $eqProfile.Font = $fontReg
            $eqProfileNew = $editorTool.DropDownItems.Add($eqProfile);

            # Add System Tray Sub-Menu Click
            $eqProfile.add_Click(
                {
                    cd $dir; .\Editor.exe "$dir\config\Parametric_EQ_$this.txt"
                }
            )
        }
    } else {

        # Add System Tray Sub-Menu Item for No EQ Profiles Found
        $eqProfile = New-Object System.Windows.Forms.ToolStripMenuItem
        $eqProfile.Text = 'No EQ Profiles Found'
        $eqProfile.Font = $fontReg
        $eqProfileNew = $editorTool.DropDownItems.Add($eqProfile);
        
    }
    $apoSepObj = New-Object System.Windows.Forms.ToolStripSeparator
    $apoSep = $contextMenu.Items.Add($apoSepObj);

    # Add Settings
    $narratorTool = $contextMenu.Items.Add('Narrator');
    $notificationTool = $contextMenu.Items.Add('Notifications');
    $audioSepObj = New-Object System.Windows.Forms.ToolStripSeparator
    $audioSep = $contextMenu.Items.Add($audioSepObj);

    # Add Resource Links
    if (
        $versionMismatch -eq $true
    ) {
        $resourceSelection = $contextMenu.Items.Add('Resources - New Version Available');
    } else {
        $resourceSelection = $contextMenu.Items.Add('Resources');
    }
    $githubLink = New-Object System.Windows.Forms.ToolStripMenuItem
    $githubLink.Text = ('www.github.com/~/AutoEq-Device-Manager')
    $githubLink.Font = $fontLink
    $githubLinkNew = $resourceSelection.DropDownItems.Add($githubLink);

    $autoEqLink = New-Object System.Windows.Forms.ToolStripMenuItem
    $autoEqLink.Text = ('www.github.com/~/AutoEq/~/results')
    $autoEqLink.Font = $fontLink
    $autoEqLinkNew = $resourceSelection.DropDownItems.Add($autoEqLink);

    $eqAPOLink = New-Object System.Windows.Forms.ToolStripMenuItem
    $eqAPOLink.Text = ('www.sourceforge.net/~/equalizerapo')
    $eqAPOLink.Font = $fontLink
    $eqAPOLinkNew = $resourceSelection.DropDownItems.Add($eqAPOLink);

    $peaceLink = New-Object System.Windows.Forms.ToolStripMenuItem
    $peaceLink.Text = ('www.sourceforge.net/~/peace-equalizer-apo-extension')
    $peaceLink.Font = $fontLink
    $peaceLinkNew = $resourceSelection.DropDownItems.Add($peaceLink);

    $githubSepObj = New-Object System.Windows.Forms.ToolStripSeparator
    $githubSep = $contextMenu.Items.Add($githubSepObj);

    # Add Window Navigation
    $exitTool = $contextMenu.Items.Add('Exit');

    # Add Menu Checks
    if (
        $global:narrator -eq $true
    ) {
        $narratorTool.Checked = $true
    }
    if (
        $global:notifications -eq $true
    ) {
        $notificationTool.Checked = $true
    }

    # Format System Tray Menu
    $outputTool.Font = $fontReg
    $listDevices.Font = $fontReg
    $configuratorTool.Font = $fontReg
    $editorTool.Font = $fontReg
    $narratorTool.Font = $fontReg
    $resourceSelection.Font = $fontReg
    $exitTool.Font = $fontBold

    # Build Menu Function Actions
    $outputTool.add_Click(
        {

            # Narrate
            if (
                $global:narrator -eq $true
            ) {
                $synth.Speak('Output')
            }

            # Retrieve output.txt
            $outputText = Get-Content -Path ($dir+'\config\output.txt')
            $outputHeight = (Get-Content -Path ($dir+'\config\output.txt')).Length
            $formHeight = [Int](($outputHeight * 18) + 60)
            if (
                $formHeight -gt [Int]750
            ) {
                $formHeight = [Int]750
            }

            # Build Output Form
            $outputObj = New-Object System.Windows.Forms.Form
            $outputObj.Text = $sysTrayAppName+' Output'
            $outputObj.Size = New-Object System.Drawing.Size @(1250, $formHeight)
            $outputObj.StartPosition = 'CenterScreen'
            $outputObj.AutoScroll = $true
            $outputObj.Icon = create_icon -UnicodeChar 0xEA37 -size 100 -theme 1 -x -10 -y 12

            # Initialize Vertical Position
            $position = 0

            # Display output.txt
            $outputText.GetEnumerator() | ForEach-Object {

                # Pack Output Label
                $outputLabel = New-Object System.Windows.Forms.Label
                $outputLabel.Location = New-Object System.Drawing.Point 0, $position
                $outputLabel.Size = New-Object System.Drawing.Point 1250, 16
                $outputObj.controls.add($outputLabel)
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
            $outputObj.Topmost = $true
            $outputObj.Add_Shown({$outputObj.Activate()})
            [void] $outputObj.ShowDialog()
        }
    )
    $listDevices.add_Click(
        {

            # Narrate
            if (
                $global:narrator -eq $true
            ) {
                $synth.Speak('Refresh Devices')
            }

            # Refresh Devices
            list_devices -dir $dir -sysTrayApp $sysTrayApp -contextMenu $contextMenu -restart $true
        }
    )
    $configuratorTool.add_Click(
        {
            Start-Process "$dir\Configurator.exe"
        }
    )
    $narratorTool.add_Click(
        {
            if (
                $global:narrator -eq $false
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
    $notificationTool.add_Click(
        {
            if (
                $global:notifications -eq $false
            ) {

                # Turn Notifications On
                $global:notifications = $true

                # Narrate
                if (
                    $global:narrator -eq $true
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
                    $global:narrator -eq $true
                ) {
                    $synth.Speak('Notifications Off')
                }

                # UnCheck Notifications
                $this.Checked = $false
            }
        }
    )
    $githubLink.add_Click(
        {
            Start-Process 'https://github.com/thomaseleff/AutoEq-Device-Manager/releases/latest'
        }
    )
    $autoEqLink.add_Click(
        {
            Start-Process 'https://github.com/jaakkopasanen/AutoEq/tree/master/results'
        }
    )
    $eqAPOLink.add_Click(
        {
            Start-Process 'https://sourceforge.net/projects/equalizerapo/'
        }
    )
    $peaceLink.add_Click(
        {
            Start-Process 'https://sourceforge.net/projects/peace-equalizer-apo-extension/'
        }
    )
    $exitTool.add_Click(
        {

            # Narrate
            if (
                $global:narrator -eq $true
            ) {
                $synth.Speak('Exit')
            }

            # Close and Clean-up
            $sysTrayApp.Visible = $false
            $appContext.ExitThread()
            Stop-Process $pid
        }
    )

    # Output Audio Devices
    $table = $table.GetEnumerator() | Sort-Object Index
    $table.GetEnumerator() | ForEach-Object {
         $_.Value
    }
}

# Build System Tray Icon Object
$sysTrayApp = New-Object System.Windows.Forms.NotifyIcon
$sysTrayApp.Text = $sysTrayAppName

# Assign Icon
#   Music Info: E90B, difficult to make out the resolution
#   Music Album: E93C, difficult to interpret
#   Music Note: EC4F
#   Music Sharing: F623, difficult to make out the resolution
#   Audio: E8D6, appears nicely
#   Equalizer: E9E9, appears nicely
#   Earbud: F4C0
#   Mix Volumes: F4C3, difficult to make out the resolution
#   Speakers: E7F5
#   Headphone: E7F6, appears nicely

$iconCode = 0xE9E9
$sysTrayApp.Icon = create_icon -UnicodeChar $iconCode -size 100 -theme $windowsTheme -x -25 -y 5
$sysTrayApp.Visible = $true

# Build System Tray Menu Object
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Pack System Tray Menu Items
list_devices -dir $dir -sysTrayApp $sysTrayApp -contextMenu $contextMenu -restart $false

# Pack System Tray Menu
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
