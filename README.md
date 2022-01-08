# AutoEQ Device Manager
The AutoEQ Device Manager integrates the Equalizer APO parametric / graphic equalizer for Windows with the AutoEQ project, providing a system-tray interface to switch playback devices and apply audio device EQ profiles.
# Usage
Instructions for setting-up the AutoEQ Device Manager.
## Requirements
- Windows Operating System (> v7)
- Equalizer APO (> v1.2.1)
- PowerShell (> v5.1)
## Installing Equalizer APO
Equalizer APO is a system-wide windows parametric / graphic equalizer, which allows for applying a single equalizer profile to a given playback device at a given time. Equalizer APO can be downloaded from [SourceForge](https://sourceforge.net/projects/equalizerapo/).

Once downloaded, follow the steps of the Setup Wizard to complete the installation.
Once installation is complete, enable Equalizer APO to control any playback device you wish through the Configurator window.
1. Within the Configurator window, check the selection box next to the Connector name in order to enable Equalizer APO for that playback device.
2. Then, click on the Connector name, highlighting it. Next, check the selection box next to "Troubleshooting options (only use in case of problems)". An additional panel of settings should appear.
3. Next, check the selection boxes for "Install APO" next to both Pre-mix and Post-mix. In the dropdown selector, ensure "Install as SFX/EFX (experimental)" is selected.
4. Repeat steps 1-3 for all playback devices you wish to enable Equalizer APO.
Once these steps are completed for all desired playback devices, select "OK" and restart your computer.
### Verifying the Equalizer APO Installation
1. To verify whether Equalizer APO was correctly installed, navigate to the audio settings in the system-tray and select a playback device that has Equalizer APO enabled.
2. Next, navigate to the Equalizer APO install directory. Default directory below.
```
C:/Program Files/EqualizerAPO
```
3. Then, navigate to the "/config" folder and open config.txt.
   - config.txt contains the parameters which set the system-wide equalizer profile
4. Next, play audio from your computer.
5. With config.txt open, delete the "-" character from the "Preamp" setting, then click save. You should experience hearing the audio becoming noticeably louder. If you do, you have successfully set up Equalizer APO.
## Verifying the PowerShell Version
1. To verify the version of PowerShell installed on your computer, navigate to the search box in the windows task bar.
2. Search for "Windows PowerShell".
3. When results appear, right-click on "Windows PowerShell" and select "Run as Administrator".
4. In the Windows PowerShell window, run the following command.
```
$PSVersionTable
```
5. The Windows PowerShell version should be displayed to the right of the PSVersion parameter in the output.
To upgrade the version of PowerShell, follow the instructions in the "Upgrading existing Windows PowerShell" sub-section of the official Microsoft documentation, [Installing Windows PowerShell](https://docs.microsoft.com/en-us/powershell/scripting/windows-powershell/install/installing-windows-powershell?view=powershell-7.2).
## Installing AutoEQ Device Manager
