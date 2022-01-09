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

# Assign Global Varibles
Set-Variable -Name 'dir' -Value (split-path $MyInvocation.MyCommand.Path -Parent)
Set-Variable -Name 'fontBold' -Value ([System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold))
Set-Variable -Name 'fontReg' -Value ([System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Regular))
Set-Variable -Name 'outputHeader' -Value ([System.Drawing.Font]::new('Courier New', 9, [System.Drawing.FontStyle]::Bold))
Set-Variable -Name 'outputBody' -Value ([System.Drawing.Font]::new('Courier New', 9, [System.Drawing.FontStyle]::Regular))

# Define Functions

function create_icon {
    param (
        $unicodeChar,
        $font,
        $size,
        $color,
        $x,
        $y
    )

    # Music-Related Icons and Segoe MDL2 Asset Unicode


    # Create Icon Bitmap from Segoe MDL2 Assets
    Set-Variable -Name 'fontIcon' -Value ([System.Drawing.Font]::new('Segoe MDL2 Assets', $size, [System.Drawing.FontStyle]::Regular))
    $brush = [System.Drawing.Brushes]::$color
    $bitmap = New-Object System.Drawing.Bitmap 16,16
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

function list_devices {
    param (
        $dir,
        $contextMenu,
        $restart
    )

    # Set Local Variables
    Set-Variable -Name 'index' -Value 0
    Set-Variable -Name 'err' -Value $false
    Set-Variable -Name 'defaultIndex' -Value 2
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
    write_output -returnAfter $false -dir $dir -str ('|                                                                       AutoEq Device Manager                                                                       |')
    write_output -returnAfter $false -dir $dir -str ('+-------------------------------------------------------------------------------------------------------------------------------------------------------------------+')
    write_output -returnAfter $true -dir $dir -str ('    Date Time: '+(Get-Date -Format G))

    # Retrive List of Parametric EQ Profiles
    if (
        [System.IO.File]::Exists($dir+'\config\eq_profiles.json') -eq $true
    ) {

        # Logging
        write_output -returnAfter $true -dir $dir -str ('NOTE: eq_profiles.json exists within the \config folder as expected.')

        # Retrieve
        $json = Get-Content -Raw -Path ($dir+'\config\eq_profiles.json') | ConvertFrom-Json
        $json = $json.GetEnumerator() | Sort-Object device
        $json | ForEach-Object {

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
                        write_output -returnAfter $true -dir $dir -str ('    ERROR: URL to [parametricConfig] not found.')
                    }     
                } else {
                    write_output -returnAfter $true -dir $dir -str ('    NOTE: Passing, Parametric EQ Config already exists within the \config folder.')
                    
                    # Add Profile to Valid List
                    $validProfiles += $_.device
                }
            } else {
                write_output -returnAfter $false -dir $dir -str ('    ERROR: URL to [parametricConfig] is invalid.')
                write_output -returnAfter $false -dir $dir -str (         '~ The URL must begin with [https://raw.githubusercontent.com/] and link to a')
                write_output -returnAfter $true -dir $dir -str (         '~ Parametric EQ profile within the [jaakkopasanen/AutoEQ] Github project.')
            }
        }

        # Set Flag
        Set-Variable -Name 'addProfiles' -Value $true
        
    } else {

        # Set Flag
        Set-Variable -Name 'addProfiles' -Value $false
        write_output -returnAfter $true -dir $dir -str ('ERROR: eq_profiles.json not found within the \config folder.')
    }

    # Clear System Tray Sub-Menu Items
    if (
        $restart -eq $true
    ) {
        $contextMenu.Items.Clear();
    }

    do {
        $index = $index + 1;
        # write_output -returnAfter $true -dir $dir -str $err
        # write_output -returnAfter $true -dir $dir -str $index
        try {

            # Check If an Audio Device is Found at Each Index
            Set-Variable -Name 'audioDevice' -Value (Get-AudioDevice -Index $index)

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
                                if (
                                    [System.IO.File]::Exists($dir+'\config\Parametric_EQ_'+$this+'.txt') -eq $true
                                ) {
                                    if (
                                        $activeDeviceName -eq ($selectedDeviceName -split ': ')[1]
                                    ) {
                                        # Switch Parametric EQ Profile
                                        Copy-Item -Path ($dir+'\config\Parametric_EQ_'+$this+'.txt') -Destination ($dir+'\config\config.txt')
                                        write_output -returnAfter $false -dir $dir -str ('Note: '+$this+' parametric EQ profile successfully assigned.')
                                    } else {

                                        # Switch Devices and Parametric EQ Profile
                                        Set-AudioDevice -Index $deviceIndex
                                        Copy-Item -Path ($dir+'\config\Parametric_EQ_'+$this+'.txt') -Destination ($dir+'\config\config.txt')
                                        write_output -returnAfter $false -dir $dir -str ('Note: '+($selectedDeviceName -split ': ')[1]+' successfully assigned with the parametric EQ profile for '+$this+'.')
                                    }
                                }
                            }
                        )
                    }

                    # Add System Try Sub-Menu Item for None
                    $eqProfile = New-Object System.Windows.Forms.ToolStripMenuItem
                    $eqProfile.Text = 'None'
                    $eqProfile.Font = $fontReg
                    $eqProfileNew = $menuDevice.DropDownItems.Add($eqProfile);

                    # Add System Tray Sub-Menu Click for None
                    $eqProfile.add_Click(
                        {
                            # Check Active Audio Device
                            Set-Variable -Name 'activeDevice' -Value (Get-AudioDevice -Playback)
                            Set-Variable -Name 'activeDeviceName' -Value $activeDevice.Name
                            Set-Variable -Name 'selectedDeviceName' -Value ($this.OwnerItem)
                            Set-Variable -Name 'deviceIndex' -Value (($selectedDeviceName -split ': ')[0])
                            if (
                                [System.IO.File]::Exists($dir+'\config\config.txt') -eq $true
                            ) {
                                if (
                                    $activeDeviceName -eq ($selectedDeviceName -split ': ')[1]
                                ) {
                                    # Remove the Parametric EQ Profile
                                    Remove-Item -Path ($dir+'\config\config.txt')
                                    write_output -returnAfter $false -dir $dir -str ('Note: Parametric EQ profile successfully unassigned.')
                                } else {

                                    # Switch Devices and Remove the Parametric EQ Profile
                                    Set-AudioDevice -Index $deviceIndex
                                    Remove-Item -Path ($dir+'\config\config.txt')
                                    write_output -returnAfter $false -dir $dir -str ('Note: '+($selectedDeviceName -split ': ')[1]+' successfully assigned with no parametric EQ profile.')
                                }
                            } else {

                                # Switch Devices
                                Set-AudioDevice -Index $deviceIndex
                                write_output -returnAfter $false -dir $dir -str ('Note: '+($selectedDeviceName -split ': ')[1]+' successfully assigned with no parametric EQ profile.')
                            }
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
        } catch {
            Set-Variable -Name 'returnedVal' -Value $_
            Set-Variable -Name 'returnedVal' -Value ($returnedVal -Replace '\s','')
            if (
                $returnedVal -eq 'NoAudioDevicewiththatIndex'
            ) {
                Set-Variable -Name 'err' -Value $true
                # write_output -returnAfter $true -dir $dir -str ('ERROR: Invalid Index.')
            } else {
                write_output -returnAfter $true -dir $dir -str ('WARNING: No Audio Device Found at Index ['+$index+'].')
            }
        }
    }
        until (
            # Use for Debugging
            # $index -eq 5

            # Use to End Loop
            $err -eq $true
        )

    # Add System Tray Menu Functions
    $toolSepObj = New-Object System.Windows.Forms.ToolStripSeparator
    $toolSep = $contextMenu.Items.Add($toolSepObj);
    $outputTool = $contextMenu.Items.Add('Output');
    $listDevices = $contextMenu.Items.Add('Refresh Devices');
    $exitSepObj = New-Object System.Windows.Forms.ToolStripSeparator
    $exitSep = $contextMenu.Items.Add($exitSepObj)
    $exitTool = $contextMenu.Items.Add('Exit');

    # Format System Tray Menu
    $outputTool.Font = $fontReg
    $listDevices.Font = $fontReg
    $exitTool.Font = $fontBold

    # Build Menu Function Actions
    $outputTool.add_Click(
        {
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
            $outputObj.Text = 'AutoEq Device Manager Output'
            $outputObj.Size = New-Object System.Drawing.Size @(1250, $formHeight)
            $outputObj.StartPosition = 'CenterScreen'
            $outputObj.AutoScroll = $true
            $outputObj.Icon = create_icon -UnicodeChar 0xEA37 -font $fontIcon -size 12 -color 'Black' -x -5 -y 0

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
                    ($_.ToLower().contains('NOTE:'.ToLower())) -Or ($_.ToLower().contains('Date Time:'.ToLower()))
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
            $outputObj.Topmost = $True
            $outputObj.Add_Shown({$outputObj.Activate()})  
            [void] $outputObj.ShowDialog()
        }
    )
    $listDevices.add_Click(
        {
            list_devices -dir $dir -contextMenu $contextMenu -restart $true
            [System.Windows.Forms.MessageBox]::Show('NOTE: Audio Device list re-generated successfully.')
        }
    )
    $exitTool.add_Click(
        {
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
$sysTrayApp.Text = 'AutoEq Device Manager'

# Assign Icon
# Music Info: E90B, difficult to make out the resolution
# Music Album: E93C, difficult to interpret
# Music Note: EC4F
# Music Sharing: F623, difficult to make out the resolution
# Audio: E8D6, appears nicely
# Equalizer: E9E9, appears nicely
# Earbud: F4C0
# Mix Volumes: F4C3, difficult to make out the resolution
# Speakers: E7F5
# Headphone: E7F6, appears nicely

$sysTrayApp.Icon = create_icon -UnicodeChar 0xE9E9 -font $fontIcon -size 12 -color 'White' -x -4 -y 0
$sysTrayApp.Visible = $true

# Build System Tray Menu Object
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Pack System Tray Menu Items
list_devices -dir $dir -contextMenu $contextMenu -restart $false

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
