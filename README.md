# Overview
A cutdown version of D2rLoader specifically aimed at tracking time spent on offline characters. Useful for tracking total time spent across each character on Grail :)

## Features
Work in progress but plan to enable the following features:
**Main features**
- Track time spent per character by assessing save game files for changes
- Launches in offline mode for those who want to help their mates fill games online but still play single player (uses parameter '-region xxx')
- Option to skip intro videos and in game videos (works with mods or if using -direct -txt)
- Ability to start Grail App of your choosing (if not already running at game launch).
- Ability to start Run timer app of your choosing (if not already running at game launch).
- Ability to show only characters your interested in on main screen (allows you to hide mules from display)
- Ability to set launch options for seed, enablerespec, playersX, resetofflinemaps.

**Swap Character Sets**<br>
Load/Unload Save game sets (eg if switching between Grail characters and edited test characters).<br>
This works by looking in subfolders within your saved game folder (excluding folders called backups or mods) for D2r save game files (.d2s) and allows you to swap these out.<br>
Note that character names should be unique if you want to ensure play time tracking for the character is accurate.<br>

**Backup**<br>
Local Backups - Ability to manually or automatically backup your save game files into "C:\Users\\%USERNAME%\Saved Games\Diablo II Resurrected\Backups". Uses my repurposed [Offline Backup script](https://github.com/shupershuff/FolderBackup)<br><br>
Cloud Backups - Setup wizard in script to allow save games folder to be backed up to the cloud. This is handy if your computer explodes or if you want to play the same characters across different devices.<br>
Uses my [D2rSinglePlayerBackup script](https://github.com/shupershuff/D2rSinglePlayerBackup). 

# Setup Steps
This has a much easier setup than Diablo2RLoader as this is an offline script and as such doesn't need any accounts added.<br>
Essentially you can just download it and run it but there are some optional config options.

## 1. Download
1. Download the latest [release here](https://github.com/shupershuff/D2rSPLoader/releases). Click on D2RSPLoader-<version>.zip (eg D2RSPLoader-1.0.0.zip) to download.
2. One downloaded, extract the zip file to a folder of your choosing.
3. Right click on D2rSPLoader.ps1 and open properties.
4. Check the "Unblock" box and click apply.<br>
![image](https://user-images.githubusercontent.com/63577525/234503557-22b7b8d4-0389-48fa-8ff4-f8a7870ccd82.png)

## 2. Script Config (Mostly Optional)
Default settings within config.xml *should* be ok but can be optionally changed. Recommend checking out the features here.
Open the .xml file in a text editor such as notepad, Powershell ISE, Notepad++ etc.
- **Most importantly**, if you have a game path that's not the default ("C:\Program Files (x86)\Diablo II Resurrected"), then you'll need to edit this to wherever you chose to install the game.<br>

All other config options below this are strictly optional:<br>
- Optionally set 'CustomLaunchArguments' to any custom launch arguments you previously used in the battlenet client for loading mods or for loading files directly.
- Optionally set 'GrailAppExecutablePath' if you want Holy Grail software to run when D2r.exe launches.
- Optionally set 'RunTimerAppExecutablePath' if you have any run timer apps you want to run when the D2r.exe launches.
- Optionally set 'CreateDesktopShortcut' to False if you don't want a handy dandy shortcut on your desktop. Enabled by default.
- Optionally set 'ShortcutCustomIconPath' to the location of a custom icon file if you want the desktop icon to be something else (eg the old D2LOD logo). Uses D2r logo by default.
- Optionally set 'ForceWindowedMode' to True if you want to force windowed mode each time. This causes issues with Diablo remembering resolution settings, so I recommend leaving this as False and manually setting your game to windowed in your game settings. Disabled by default.
- Optionally set 'AutoBackup' to true if you want the script to automatically locally backup your save game files every half hour.
- Optionally set 'DisableVideos' to true if you want to skip intro videos or videos between acts.
- Optionally set 'ManualSettingSwitcherEnabled' to True if you want the ability to be able to choose a settings profile to load from. Once enabled, this is toggleable from the script using 's'. See the [Manual Setting Switcher](#6-manual-settings-switcher-optional) section below for more info. Disabled by default.
<br>
Some of the above options can be set within the scripts option menu also. Make sure to save the file when finished editing.<br>

## 3. Run the script manually for the first time
1. Browse to the folder, right click on D2rSPLoader.ps1 and choose run (if you see an option for 'Run with PowerShell', use this).
2. If you get prompted to change the execution policy so you can run the script, type y and press enter.
   ![image](https://user-images.githubusercontent.com/63577525/234580880-e78df284-edea-4a5e-b4c6-4825f6031b4e.png)   
   a) If the script opens up and immediately closes or you instead get a message about "D2rSPLoader.ps1 cannot be loaded because running scripts is disabled on this system" then you will need to perform the following steps:   
   b) Open the start menu and type in powershell. Right click on PowerShell and click "Run as administrator".<br>
   c) Enter the following command: **Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser**<br>
   d) Type in "y" and press enter to confirm.<br>
   e) Run the D2rSPLoader.ps1 script again.<br>
3. If the script prompts to trust it and add it to the unblock list, type in y and press enter to confirm.
4. This will perform the first time setup to compile settext.exe, populate characters.csv and create a shortcut on your desktop.

## 4. Manual Settings Switcher (Optional)
Do you want to manually choose which settings to use when launching the game? This is for you! This feature is disabled by default, as this needs to be setup first and understood this first.<br>
<br>
To enable this feature in the setting 'ManualSettingSwitcherEnabled' must be set to True. You can do this from the options menu or from editing the .

**Setting up alternate Game settings to switch too**.<br>
1. Set 'ManualSettingSwitcherEnabled' to True in your [config file](#2-script-config-mostly-optional) or from the options menu in the script.
2. Launch the Game (via the Loader or via Bnet client, doesn't matter).
3. Make the required graphics/audio/game changes via the menu.
4. Close the game.
5. Browse to "C:\Users\\\<yourusername>\Saved Games\Diablo II Resurrected"
6. Copy the Settings.json file and paste into the same folder.
7. Rename this copied file to Settings._name_.json (eg ensure the name is inside two fullstops eg settings.1440pHigh.json or settings.PotatoGraphics.json)
8. Press S on the main screen of the loader to enable prompting for settings. 
9. Launch the game via the loader, choose settings and proceed find all of the high runes. All of them.
<br>
<br>
<br>
Page views as of 2nd Oct 2024:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[![Hits](https://hits.seeyoufarm.com/api/count/incr/badge.svg?url=https%3A%2F%2Fgithub.com%2Fshupershuff%2FD2rSPLoader&count_bg=%2379C83D&title_bg=%23555555&icon=&icon_color=%23E7E7E7&title=hits&edge_flat=false)](https://www.youtube.com/watch?v=dQw4w9WgXcQ)<br>
