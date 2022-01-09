# AutoEq Device Manager
The AutoEq Device Manager integrates the [Equalizer APO](https://sourceforge.net/projects/equalizerapo/) parametric equalizer for Windows with the [AutoEq](https://github.com/jaakkopasanen/AutoEq) project, providing a system-tray tool to switch playback devices and apply audio device EQ profiles. This tool is written in Windows PowerShell and utilizes the [AudioDeviceCmdlets](https://github.com/frgnca/AudioDeviceCmdlets) module to manage Windows playback devices.

# Features
- Ability to switch between any connected playback device
- Ability to change or remove parametric EQ profiles for any number of user-configured audio devices
- Output window for viewing status of the application tasks

# Installation
Instructions for setting-up the AutoEq Device Manager.

## Requirements
- Windows Operating System (minimum Windows 10)
- Equalizer APO (minimum v1.2.1)
- PowerShell (minimum v5.1)

  _Minimum requirements are known working versions. Any previous version may also be valid, but has not been verified to be working._

## Installing Equalizer APO
Equalizer APO is a system-wide Windows parametric equalizer, which allows for applying a single equalizer profile to a given playback device at a given time. Equalizer APO can be downloaded from [SourceForge](https://sourceforge.net/projects/equalizerapo/).

Once downloaded, follow the steps of the Setup Wizard to complete the installation.
Once installation is complete, enable Equalizer APO to control any playback device you wish through the Configurator window.

1. Within the Configurator window, check the selection box next to the Connector name in order to enable Equalizer APO for that playback device.
2. Then, click on the Connector name, highlighting it. Next, check the selection box next to "Troubleshooting options (only use in case of problems)". An additional panel of settings should appear.
3. Next, check the selection boxes for "Install APO" next to both "Pre-mix" and "Post-mix". In the dropdown selector, ensure "Install as SFX/EFX (experimental)" is selected.
4. Repeat steps 1-3 for all playback devices you wish to enable Equalizer APO.
5. Once these steps are completed for all desired playback devices, select "OK" and restart your computer.

At any time in the future, if you wish to configure a new playback device or disable Equalizer APO for an existing playback device, you can open and change settings within the configurator by running the following program. Default file path below.

```
C:/Program Files/EqualizerAPO/Configurator.exe
```
### Verifying the Equalizer APO Installation
1. To verify whether Equalizer APO was correctly installed, navigate to the audio settings in the system-tray and select a playback device that has Equalizer APO enabled.
2. Next, navigate to the Equalizer APO install directory. Default directory below.

      ```
      C:/Program Files/EqualizerAPO
      ```
3. Then, navigate to the /config folder and open config.txt.
   - config.txt contains the parameters that set the system-wide equalizer profile. These parameters can be manually edited, or they can be replaced with Parametric EQ profiles specific to a given audio device, like those created within the [AutoEq](https://github.com/jaakkopasanen/AutoEq) project.
4. Next, play audio from your computer.
5. With config.txt open, delete the "-" character from the "Preamp" setting, then save. You should hear the audio become noticeably louder. If you do, you have successfully set up Equalizer APO.
6. Repeat steps 1-5 for all playback devices you have enabled Equalizer APO.

## Verifying Windows PowerShell
1. To verify the version of Windows PowerShell installed on your computer, navigate to the search box in the Windows Taskbar.
2. Search for "Windows PowerShell".
3. When results appear, right-click on "Windows PowerShell" and select "Run as Administrator".
4. In the Windows PowerShell window, run the following command.

      ```
      $PSVersionTable
      ```
5. The Windows PowerShell version should be displayed to the right of the "PSVersion" parameter in the output.

To upgrade the version of PowerShell, follow the instructions in the "Upgrading existing Windows PowerShell" sub-section of the official Microsoft documentation, [Installing Windows PowerShell](https://docs.microsoft.com/en-us/powershell/scripting/windows-powershell/install/installing-windows-powershell?view=powershell-7.2).

## Downloading AutoEq Device Manager
The AutoEq Device Manager consists of two components,
1. AutoEq_DeviceManager.ps1
2. eq_profiles.json

Navigate to [AutoEq_DeviceManager.ps1](AutoEq_DeviceManager.ps1), copying the contents into a ".ps1" file within the same directory as the Equalizer APO installation. Default directory below.
- File name **IS NOT** important.

   ```
   C:/Program Files/EqualizerAPO
   ```

Navigate to [eq_profiles.json](config/eq_profiles.json), copying the contents into a file named "eq_profiles.json" within the /config sub-directory of the Equalizer APO installation. Default directory below.
- File name **IS** important, the file **MUST** be named eq_profiles.json.

   ```
   C:/Program Files/EqualizerAPO/config
   ```

# Usage
Configuring and running the AutoEq Device Manager.

## Configuring eq_profiles.json
eq_profiles.json contains the user-configured names of your audio devices and the corresponding URL path to the Parametric EQ profiles from [AutoEq/results](https://github.com/jaakkopasanen/AutoEq/tree/master/results). Any number of devices and profiles can be configured within eq_profiles.json. The AutoEq Device Manager will validate all URL paths to ensure they source valid parametric EQ profiles and will download the profiles automatically into the /config sub-directory. Default sub-directory below.

   ```
   C:/Program Files/EqualizerAPO/config
   ```
To modify eq_profiles.json,
1. Open eq_profiles.json in a text editor.
2. Modify existing entries or add new entries in or to eq_profiles.json for all audio devices you wish to specify unique Parametric EQ profiles for through Equalizer APO.
3. Each entry requires two parameters, "device" and "parametricConfig", to be assigned values. Each parameter is described below.
   - device: Quoted string for the user-provided name of the audio device. Not required to match the name of the device from [AutoEq/results](https://github.com/jaakkopasanen/AutoEq/tree/master/results).
   - parametricConfig: Quoted string for the URL path to the Raw Parametric EQ profile from [AutoEq/results](https://github.com/jaakkopasanen/AutoEq/tree/master/results). This **MUST** be the URL path to a **Parametric** EQ profile pointing directly to the **Raw** file. Example URL path to the Parametric EQ profile for Sennheiser HD 600 below.

        ```
        https://raw.githubusercontent.com/jaakkopasanen/AutoEq/master/results/oratory1990/harman_over-ear_2018/Sennheiser%20HD%20600/Sennheiser%20HD%20600%20ParametricEQ.txt
        ```
      - The Raw file can be found by clicking the "Raw" button on the right-hand side of a given Parametric EQ file's options, highlighted below.


   ![Locating the Raw File URL Path](https://user-images.githubusercontent.com/49733042/148661224-cf7d3091-4abc-4016-8360-da792e6efadb.png)


## Running AutoEq Device Manager
AutoEq Device Manager can be run manually through a Windows PowerShell session or scheduled through the Windows Task Scheduler to run at log-on.

### Running Manually Through a Windows PowerShell Session
It is recommended to first run the AutoEq Device Manager manually in order to verify set-up was completed successfully.

1. Navigate to the search box in the Windows Taskbar.
2. Search for "Windows PowerShell".
3. When results appear, right-click on "Windows PowerShell" and select "Run as Administrator".
4. When a new session window opens, run AutoEq_DeviceManager.ps1. Default command below is based on the default directory and filename.

      ```
      PowerShell -File "C:/Program Files/EqualizerAPO/AutoEq_DeviceManager.ps1"
      ```
5. Accept all prompts to install the required AudioDeviceCmdlets module.

### Scheduling Through Windows Task Scheduler to Run at Log-On
Once the AutoEq Device Manager is verified to be working successfully, create a new task in the Windows Task Scheduler to automatically run the AutoEq Device Manager at log-on.

1. Navigate to the search box in the Windows Taskbar.
2. Search for "Task Scheduler".
3. When results appear, open "Task Scheduler".
4. When the application window appears, click "Create Task" within the "Actions" panel on the right-hand side of the window.
5. Within the "General" tab of the pop-up window, enter a "Name" for the task and select "Run only when user is logged on."
6. Within the "Triggers" tab of the pop-up window, select "New".
7. Within the "New Trigger" pop-up window, select "At log on" from the "Begin the task" drop-down options.
8. Click "OK".
9. Within the "Actions" tab of the pop-up window, select "New".
10. Within the "New Action" pop-up window, enter the following into the "Program/script" field.

      ```
      PowerShell
      ```
11. Then, enter the following into the "Add arguments (optional)" field.

      ```
      -File "C:/Program Files/EqualizerAPO/AutoEq_DeviceManager.ps1" -WindowStyle Hidden
      ```
12. Click "OK".
13. Within the "Settings" tab of the pop-up, un-check the selection box next to the "Stop the task if it runs longer than" field.
14. Click "OK".
15. The newly created task should now appear within the Task Scheduler Library, and can be run manually by selecting the task, then by clicking Run within the "Actions" panel on the righ-hand side of the window.
17. Verify that the newly created task was setup correctly, log-off and on again. Once logged-on, AutoEq Device Manager should be running in the system-tray, appearing with a notepad icon.

# Attribution
Windows PowerShell scripting methods for creating system-tray tools were sourced from the following two articles from [https://www.systanddeploy.com](https://www.systanddeploy.com):
- [Create your own PowerShell systray/taskbar tool](https://www.systanddeploy.com/2018/12/create-your-own-powershell.html), Damien Van Robaeys
- [Build a PowerShell systray tool with menus, sub menus and pictures](https://www.systanddeploy.com/2020/09/build-powershell-systray-tool-with.html), Damien Van Robaeys
