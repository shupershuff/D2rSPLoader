<#
Author: Shupershuff
Usage:
Happy for you to make any modifications to this script for your own needs providing:
- Any variants of this script are never sold.
- Any variants of this script published online should always be open source.
Purpose:
	Script is mainly orientated around tracking character playtime and total game time for single player.
	Script will track character details from CSV.
Instructions: See GitHub readme https://github.com/shupershuff/D2rSPLoader

Changes since 1.0.0 (next version edits):
- Changed Skipped backup message
- Removed "multibox" from welcome banner.
- Fixed typo in config.xml
- Added mitigation if 'Saved Games' isn't in default location

1.0.0+ to do list
Couldn't write :) in release notes without it adding a new line, some minor issue with formatfunction regex
Fix whatever I broke or poorly implemented in the last update :)
#>


$CurrentVersion = "1.0.1"
###########################################################################################################################################
# Script itself
###########################################################################################################################################
$host.ui.RawUI.WindowTitle = "Diablo 2 Resurrected: Single Player Loader"
#run script as admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")){ Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $ScriptArguments"  -Verb RunAs;exit }
#DebugMode
#$DebugMode = $True # Uncomment to enable
if ($DebugMode -eq $True){
	$DebugPreference = "Continue"
	$VerbosePreference = "Continue"
}
#set window size
[console]::WindowWidth=77; #script has been designed around this width. Adjust at your own peril.
$WindowHeight=50 #Can be adjusted to preference, but not less than 42
do {
	Try{
		[console]::WindowHeight = $WindowHeight;
		$HeightSuccessfullySet = $True
	}
	Catch {
		$WindowHeight --
	}
} Until ($HeightSuccessfullySet -eq $True)
[console]::BufferWidth=[console]::WindowWidth
#set misc vars
$Script:X = [char]0x1b #escape character for ANSI text colors
$ProgressPreference = "SilentlyContinue"
$Script:WorkingDirectory = ((Get-ChildItem -Path $PSScriptRoot)[0].fullname).substring(0,((Get-ChildItem -Path $PSScriptRoot)[0].fullname).lastindexof('\')) #Set Current Directory path.
$Script:CharacterSavePath = (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" -name "{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}")."{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}" #Get Saved Games folder from registry rather than assume it's in C:\Users\Username\Saved Games
$Script:SettingsProfilePath = $Script:CharacterSavePath
$Script:StartTime = Get-Date #Used for elapsed time. Is reset when script refreshes.
$Script:LastBackup = $Script:StartTime.addminutes(-60)
$Script:MOO = "%%%"
$Script:JobIDs = @()
$MenuRefreshRate = 30 #How often the script refreshes in seconds. This should be set to 30, don't change this please.
$Script:ScriptFileName = Split-Path $MyInvocation.MyCommand.Path -Leaf #find the filename of the script in case a user renames it.
$Script:SessionTimer = 0 #set initial session timer to avoid errors in info menu.
$Script:NotificationHasBeenChecked = $False
#Baseline of acceptable characters for ReadKey functions. Used to prevents receiving inputs from folk who are alt tabbing etc.
$Script:AllowedKeyList = @(48,49,50,51,52,53,54,55,56,57) #0 to 9
$Script:AllowedKeyList += @(96,97,98,99,100,101,102,103,104,105) #0 to 9 on numpad
$Script:AllowedKeyList += @(65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90) # A to Z
$Script:MenuOptions = @(65,66,67,68,71,73,74,79,82,83,84,88) #a, b, c, d, g, i, j, o, r, s, t and x. Used to detect singular valid entries where script can have two characters entered.
$EnterKey = 13
Function RemoveMaximiseButton { # I'm removing the maximise button on the script as sometimes I misclick maximise instead of minimise and it annoys me. Copied this straight out of ChatGPT lol.
	Add-Type @"
		using System;
		using System.Runtime.InteropServices;
		public class WindowAPI {
			public const int GWL_STYLE = -16;
			public const int WS_MAXIMIZEBOX = 0x10000;
			public const int WS_THICKFRAME = 0x40000;  // Window has a sizing border
			[DllImport("user32.dll")]
			public static extern IntPtr GetForegroundWindow();
			[DllImport("user32.dll", SetLastError = true)]
			public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
			[DllImport("user32.dll", SetLastError = true)]
			public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
		}
"@
	# Get the handle for the current window (PowerShell window)
	$hWnd = [WindowAPI]::GetForegroundWindow()
	# Get the current window style
	$style = [WindowAPI]::GetWindowLong($hWnd, [WindowAPI]::GWL_STYLE)
	# Disable the maximize button by removing the WS_MAXIMIZEBOX style. Also disable resizing the width.
	$newStyle = $style -band -bnot ([WindowAPI]::WS_MAXIMIZEBOX -bor [WindowAPI]::WS_THICKFRAME)
	[WindowAPI]::SetWindowLong($hWnd, [WindowAPI]::GWL_STYLE, $newStyle) | out-null
}
Function ReadKey([string]$message=$Null,[bool]$NoOutput,[bool]$AllowAllKeys){#used to receive user input
	$Script:key = $Null
	$Host.UI.RawUI.FlushInputBuffer()
	if (![string]::IsNullOrEmpty($message)){
		Write-Host -NoNewLine $message
	}
	$AllowedKeyList = $Script:AllowedKeyList + @(13,27) #Add Enter & Escape to the allowedkeylist as acceptable inputs.
	while ($Null -eq $key){
	if ($Host.UI.RawUI.KeyAvailable){
			$key_ = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp")
			if ($True -ne $AllowAllKeys){
				if ($key_.KeyDown -and $key_.VirtualKeyCode -in $AllowedKeyList){
					$key = $key_
				}
			}
			else {
				if ($key_.KeyDown){
					$script:key = $key_
				}
			}
		}
		else {
			Start-Sleep -m 200  # Milliseconds
		}
	}
	if ($key_.VirtualKeyCode -ne $EnterKey -and -not ($Null -eq $key) -and [bool]$NoOutput -ne $true){
		Write-Host ("$X[38;2;255;165;000;22m" + "$($key.Character)" + "$X[0m") -NoNewLine
	}
	if (![string]::IsNullOrEmpty($message)){
		Write-Host "" # newline
	}
	return $(
		if ($Null -eq $key -or $key.VirtualKeyCode -eq $EnterKey){
			""
		}
		ElseIf ($key.VirtualKeyCode -eq 27){ #if key pressed was escape
			"Esc"
		}
		else {
			$key.Character
		}
	)
}
Function ReadKeyTimeout([string]$message=$Null, [int]$timeOutSeconds=0, [string]$Default=$Null, [object[]]$AdditionalAllowedKeys = $null, [bool]$TwoDigitAcctSelection = $False){
	$Script:key = $Null
	$inputString = ""
	$Host.UI.RawUI.FlushInputBuffer()
	if (![string]::IsNullOrEmpty($message)){
		Write-Host -NoNewLine $message
	}
	$Counter = $timeOutSeconds * 1000 / 250
	$AllowedKeyList = $Script:AllowedKeyList + $AdditionalAllowedKeys #Add any other specified allowed key inputs (eg Enter).
	while ($Null -eq $key -and ($timeOutSeconds -eq 0 -or $Counter-- -gt 0)){
		if ($TwoDigitAcctSelection -eq $True -and $inputString.length -ge 1){
			$AllowedKeyList = $AllowedKeyList + 13 + 8 # Allow enter and backspace to be used if 1 character has been typed.
		}
		if (($timeOutSeconds -eq 0) -or $Host.UI.RawUI.KeyAvailable){
			$key_ = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp")
			if ($key_.KeyDown -and $key_.VirtualKeyCode -in $AllowedKeyList){
				if ($key_.VirtualKeyCode -eq [System.ConsoleKey]::Backspace){
					$Counter = $timeOutSeconds * 1000 / 250 #reset counter
					if ($inputString.Length -gt 0){
						$inputString = $inputString.Substring(0, $inputString.Length - 1) #remove last added character/number from variable
						# Clear the last character from the console
						$Host.UI.RawUI.CursorPosition = @{
							X = [Math]::Max($Host.UI.RawUI.CursorPosition.X - 1, 0)
							Y = $Host.UI.RawUI.CursorPosition.Y
						}
						Write-Host -NoNewLine " " #-ForegroundColor Black
						$Host.UI.RawUI.CursorPosition = @{
							X = [Math]::Max($Host.UI.RawUI.CursorPosition.X - 1, 0)
							Y = $Host.UI.RawUI.CursorPosition.Y
						}
					}
				}
				ElseIf ($TwoDigitAcctSelection -eq $True -and $key_.VirtualKeyCode -notin $Script:MenuOptions + 27){
					$Counter = $timeOutSeconds * 1000 / 250 #reset counter
					if ($key_.VirtualKeyCode -eq $EnterKey -or $key_.VirtualKeyCode -eq 27){
						break
					}
					$inputString += $key_.Character
					Write-Host ("$X[38;2;255;165;000;22m" + $key_.Character + "$X[0m") -nonewline
					if ($inputString.length -eq 2){#if 2 characters have been entered
						break
					}
				}
				Else {
					$script:key = $key_
					$inputString = $key_.Character
				}
			}
		}
		else {
			Start-Sleep -m 250 # Milliseconds
		}
	}
	if ($Counter -le 0){
		if ($InputString.Length -gt 0){# if it timed out, revert to no input if one character was entered.
			$InputString = "" #remove last added character/number from variable
		}
	}
	if ($TwoDigitAcctSelection -eq $False -or ($TwoDigitAcctSelection -eq $True -and $key_.VirtualKeyCode -in $Script:MenuOptions)){
		Write-Host ("$X[38;2;255;165;000;22m" + "$inputString" + "$X[0m")
	}
	if (![string]::IsNullOrEmpty($message) -or $TwoDigitAcctSelection -eq $True){
		Write-Host "" # newline
	}
	Write-Host #prevent follow up text from ending up on the same line.
	return $(
		If ($key.VirtualKeyCode -eq $EnterKey -and $EnterKey -in $AllowedKeyList){
			""
		}
		ElseIf ($key.VirtualKeyCode -eq 27){ #if key pressed was escape
			"Esc"
		}
		ElseIf ($inputString.Length -eq 0){
			$Default
		}
		else {
			$inputString
		}
	)
}
Function PressTheAnyKey {#Used instead of Pause so folk can hit any key to continue
	Write-Host "  Press any key to continue..." -nonewline
	readkey -NoOutput $True -AllowAllKeys $True | out-null
	Write-Host
}
Function PressTheAnyKeyToExit {#Used instead of Pause so folk can hit any key to exit
	Write-Host "  Press Any key to exit..." -nonewline
	readkey -NoOutput $True -AllowAllKeys $True | out-null
	remove-job * -force
	Exit
}
Function Red {
	process { Write-Host $_ -ForegroundColor Red }
}
Function Yellow {
	process { Write-Host $_ -ForegroundColor Yellow }
}
Function Green {
	process { Write-Host $_ -ForegroundColor Green }
}
Function NormalText {
	process { Write-Host $_ }
}
Function FormatFunction { # Used to get long lines formatted nicely within the CLI. Possibly the most difficult thing I've created in this script. Hooray for Regex!
	param (
		[string] $Text,
		[int] $Indents,
		[int] $SubsequentLineIndents,
		[switch] $IsError,
		[switch] $IsWarning,
		[switch] $IsSuccess
	)
	if ($IsError -eq $True){
		$Colour = "Red"
	}
	ElseIf ($IsWarning -eq $True){
		$Colour = "Yellow"
	}
	ElseIf ($IsSuccess -eq $True){
		$Colour = "Green"
	}
	Else {
		$Colour = "NormalText"
	}
	$MaxLineLength = 76
	If ($Indents -ge 1){
		while ($Indents -gt 0){
			$Indent += " "
			$Indents --
		}
	}
	If ($SubsequentLineIndents -ge 1){
		while ($SubsequentLineIndents -gt 0){
			$SubsequentLineIndent += " "
			$SubsequentLineIndents --
		}
	}
	$Text -split "`n" | ForEach-Object {
		$Line = " " + $Indent + $_
		$SecondLineDeltaIndent = ""
		if ($Line -match '^[\s]*-'){ #For any line starting with any preceding spaces and a dash.
			$SecondLineDeltaIndent = "  "
		}
		if ($Line -match '^[\s]*\d+\.\s'){ #For any line starting with any preceding spaces, a number, a '.' and a space. Eg "1. blah".
			$SecondLineDeltaIndent = "   "
		}
		Function Formatter ([string]$line){
			$pattern = "[\e]?[\[]?[`"-,`.!']?\b[\w\-,'`"]+(\S*)" # Regular expression pattern to find the last word including any trailing non-space characters. Also looks to include any preceding special characters or ANSI escape character.
			$WordMatches = [regex]::Matches($Line, $pattern) # Find all matches of the pattern in the string
			# Initialize variables to track the match with the highest index
			$highestIndex = -1
			$SelectedMatch = $Null
			$PatternLengthCount = 0
			$ANSIPatterns = "\x1b\[38;\d{1,3};\d{1,3};\d{1,3};\d{1,3};\d{1,3}m","\x1b\[0m","\x1b\[4m"
			ForEach ($WordMatch in $WordMatches){# Iterate through each match (match being a block of characters, ie each word).
				ForEach ($ANSIPattern in $ANSIPatterns){ #iterate through each possible ANSI pattern to find any text that might have ANSI formatting.
					$ANSIMatches = $WordMatch.value | Select-String -Pattern $ANSIPattern -AllMatches
					ForEach ($ANSIMatch in $ANSIMatches){
						$Script:ANSIUsed = $True
						$PatternLengthCount = $PatternLengthCount + (($ANSIMatch.matches | ForEach-Object {$_.Value}) -join "").length #Calculate how many characters in the text are ANSI formatting characters and thus won't be displayed on screen, to prevent skewing word count.
					}
				}
				$matchIndex = $WordMatch.Index
				$matchLength = $WordMatch.Length
				$matchEndIndex = $matchIndex + $matchLength - 1
				if ($matchEndIndex -lt ($MaxLineLength + $PatternLengthCount)){# Check if the match ends within the first $MaxLineLength characters
					if ($matchIndex -gt $highestIndex){# Check if this match has a higher index than the current highest
						$highestIndex = $matchIndex # This word has a higher index and is the winner thus far.
						$SelectedMatch = $WordMatch
						$lastspaceindex = $SelectedMatch.Index + $SelectedMatch.Length - 1 #Find the index (the place in the string) where the last word can be used without overflowing the screen.
					}
				}
			}
			try {
				$script:chunk = $Line.Substring(0, $lastSpaceIndex + 1) #Chunk of text to print to screen. Uses all words from the start of $line up until $lastspaceindex so that only text that fits on a single line is printed. Prevents words being cut in half and prevents loss of indenting.
			}
			catch {
				$script:chunk = $Line.Substring(0, [Math]::Min(($MaxLineLength), ($Line.Length))) #If the above fails for whatever reason. Can't exactly remember why I put this in here but leaving it in to be safe LOL.
			}
		}
		Formatter $Line
		if ($Script:ANSIUsed -eq $True){ #if fancy pants coloured text (ANSI) is used, write out the first line. Check if ANSI was used in any overflow lines.
			do {
				$Script:ANSIUsed = $False
				Write-Output $Chunk | out-host #have to use out-host due to pipeline shenanigans and at this point was too lazy to do things properly :)
				$Line = " " + $SubsequentLineIndent + $Indent + $Line.Substring($chunk.Length).trimstart() #$Line is equal to $Line but without the text that's already been outputted.
				Formatter $Line
			} until ($Script:ANSIUsed -eq $False)
			if ($Chunk -ne " " -and $Chunk.lenth -ne 0){#print any remaining text.
				Write-Output $Chunk | out-host
			}
		}
		Else { #if line has no ANSI formatting.
			Write-Output $Chunk | &$Colour
		}
		$Line = $Line.Substring($chunk.Length).trimstart() #remove the string that's been printed on screen from variable.
		if ($Line.length -gt 0){ # I see you're reading my comment. How thorough of you! This whole function was an absolute mindf#$! to come up with and took probably 30 hours of trial, error and rage (in ascending order of frequency). Odd how the most boring of functions can take up the most time :)
				Write-Output ($Line -replace "(.{1,$($MaxLineLength - $($Indent.length) - $($SubsequentLineIndent.length) -1 - $($SecondLineDeltaIndent.length))})(\s+|$)", " $SubsequentLineIndent$SecondLineDeltaIndent$Indent`$1`n").trimend() | &$Colour
		}
	}
}
Function CommaSeparatedList {
	param (
		[object] $Values,
		[switch] $NoOr,
		[switch] $AndText
	)
	ForEach ($Value in $Values){ #write out each account option, comma separated but show each option in orange writing. Essentially output overly complicated fancy display options :)
		if ($Value -ne $Values[-1]){
			Write-Host "$X[38;2;255;165;000;22m$Value$X[0m" -nonewline
			if ($Value -ne $Values[-2]){Write-Host ", " -nonewline}
		}
		else {
			if ($Values.count -gt 1){
				$AndOr = "or"
				if ($AndText -eq $True){
					$AndOr = "and"
				}
				if ($NoOr -eq $False){
					Write-Host " $AndOr " -nonewline
				}
				Else {
					Write-Host ", " -nonewline
				}
			}
			Write-Host "$X[38;2;255;165;000;22m$Value$X[0m" -nonewline
		}
	}
}
Function InitialiseCurrentStats {
	if ((Test-Path -Path "$Script:WorkingDirectory\Stats.csv") -ne $true){#Create Stats CSV if it doesn't exist
		$Null = {} | Select-Object "TotalGameTime","TimesLaunched","LastUpdateCheck","HighRunesFound","UniquesFound","SetItemsFound","RaresFound","MagicItemsFound","NormalItemsFound","Gems","CowKingKilled","PerfectGems" | Export-Csv "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation
		Write-Host " Stats.csv created!"
	}
	do {
		Try {
			$Script:CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv" #Get current stats csv details
		}
		Catch {
			Write-Host " Unable to import stats.csv. File corrupt or missing." -foregroundcolor red
		}
		if ($null -ne $CurrentStats){
			#Todo: In the Future add CSV validation checks
			$StatsCSVImportSuccess = $True
		}
		else {#Error out and exit if there's a problem with the csv.
				if ($StatsCSVRecoveryAttempt -lt 1){
					try {
						Write-Host " Attempting Autorecovery of stats.csv from backup..." -foregroundcolor red
						Copy-Item -Path $Script:WorkingDirectory\Stats.backup.csv -Destination $Script:WorkingDirectory\Stats.csv -ErrorAction stop
						Write-Host " Autorecovery successful!" -foregroundcolor Green
						$StatsCSVRecoveryAttempt ++
						PressTheAnyKey
					}
					Catch {
						$StatsCSVImportSuccess = $False
					}
				}
				Else {
					$StatsCSVRecoveryAttempt = 2
				}
				if ($StatsCSVImportSuccess -eq $False -or $StatsCSVRecoveryAttempt -eq 2){
					Write-Host "`n Stats.csv is corrupted or empty." -foregroundcolor red
					Write-Host " Replace with data from stats.backup.csv or delete stats.csv`n" -foregroundcolor red
					PressTheAnyKeyToExit
				}
			}
	} until ($StatsCSVImportSuccess -eq $True)
	if (-not ($CurrentStats | Get-Member -Name "LastUpdateCheck" -MemberType NoteProperty -ErrorAction SilentlyContinue)){#For update 1.8.1+. If LastUpdateCheck column doesn't exist, add it to the CSV data
		$Script:CurrentStats | ForEach-Object {
			$_ | Add-Member -NotePropertyName "LastUpdateCheck" -NotePropertyValue "2000.06.28 12:00:00" #previously "28/06/2000 12:00:00 pm"
		}
	}
	ElseIf ($CurrentStats.LastUpdateCheck -eq "" -or $CurrentStats.LastUpdateCheck -like "*/*"){# If script has just been freshly downloaded or has the old Date format.
		$Script:CurrentStats.LastUpdateCheck = "2000.06.28 12:00:00" #previously "28/06/2000 12:00:00 pm"
		$CurrentStats | Export-Csv "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation
	}
}
Function CheckForUpdates {
	#Only Check for updates if updates haven't been checked in last 8 hours. Reduces API requests.
	if ($Script:CurrentStats.LastUpdateCheck -lt (Get-Date).addHours(-8).ToString('yyyy.MM.dd HH:mm:ss')){# Compare current date and time to LastUpdateCheck date & time.
		try {
			# Check for Updates
			$Releases = Invoke-RestMethod -Uri "https://api.github.com/repos/shupershuff/D2rSPLoader/releases"
			$ReleaseInfo = ($Releases | Sort-Object id -desc)[0] #find release with the highest ID.
			$Script:LatestVersion = [version[]]$ReleaseInfo.Name.Trim('v')
			if ($Script:LatestVersion -gt $Script:CurrentVersion){ #If a newer version exists, prompt user about update details and ask if they want to update.
				Write-Host "`n Update available! See Github for latest version and info" -foregroundcolor Yellow -nonewline
				if ([version]$CurrentVersion -in (($Releases.name.Trim('v') | ForEach-Object { [version]$_ } | Sort-Object -desc)[2..$releases.count])){
					Write-Host ".`n There have been several releases since your version." -foregroundcolor Yellow
					Write-Host " Checkout Github releases for fixes/features added. " -foregroundcolor Yellow
					Write-Host " $X[38;2;69;155;245;4mhttps://github.com/shupershuff/D2rSPLoader/releases/$X[0m`n"
				}
				Else {
					Write-Host ":`n $X[38;2;69;155;245;4mhttps://github.com/shupershuff/D2rSPLoader/releases/latest$X[0m`n"
				}
				FormatFunction -Text $ReleaseInfo.body #Output the latest release notes in an easy to read format.
				Write-Host; Write-Host
				Do {
					Write-Host " Your Current Version is v$CurrentVersion."
					Write-Host (" Would you like to update to v"+ $Script:LatestVersion + "? $X[38;2;255;165;000;22mY$X[0m/$X[38;2;255;165;000;22mN$X[0m: ") -nonewline
					$ShouldUpdate = ReadKey
					if ($ShouldUpdate -eq "y" -or $ShouldUpdate -eq "yes" -or $ShouldUpdate -eq "n" -or $ShouldUpdate -eq "no"){
						$UpdateResponseValid = $True
					}
					Else {
						Write-Host "`n Invalid response. Choose $X[38;2;255;165;000;22mY$X[0m $X[38;2;231;072;086;22mor$X[0m $X[38;2;255;165;000;22mN$X[0m.`n" -ForegroundColor red
					}
				} Until ($UpdateResponseValid -eq $True)
				if ($ShouldUpdate -eq "y" -or $ShouldUpdate -eq "yes"){#if user wants to update script, download .zip of latest release, extract to temporary folder and replace old D2rSPLoader.ps1 with new D2rSPLoader.ps1
					Write-Host "`n Updating... :)" -foregroundcolor green
					try {
						New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") -ErrorAction stop | Out-Null #create temporary folder to download zip to and extract
					}
					Catch {#if folder already exists for whatever reason.
						Remove-Item -Path ($Script:WorkingDirectory + "\UpdateTemp\") -Recurse -Force
						New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") | Out-Null #create temporary folder to download zip to and extract
					}
					$ZipURL = $ReleaseInfo.zipball_url #get zip download URL
					$ZipPath = ($WorkingDirectory + "\UpdateTemp\D2rSPLoader_" + $ReleaseInfo.tag_name + "_temp.zip")
					Invoke-WebRequest -Uri $ZipURL -OutFile $ZipPath
					if ($Null -ne $releaseinfo.assets.browser_download_url){#Check If I didn't forget to make a version.zip file and if so download it. This is purely so I can get an idea of how many people are using the script or how many people have updated. I have to do it this way as downloading the source zip file doesn't count as a download in github and won't be tracked.
						Invoke-WebRequest -Uri $releaseinfo.assets.browser_download_url -OutFile $null | out-null #identify the latest file only.
					}
					$ExtractPath = ($Script:WorkingDirectory + "\UpdateTemp\")
					Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
					$FolderPath = Get-ChildItem -Path $ExtractPath -Directory -Filter "shupershuff*" | Select-Object -ExpandProperty FullName
					Copy-Item -Path ($FolderPath + "\D2rSPLoader.ps1") -Destination ($Script:WorkingDirectory + "\" + $Script:ScriptFileName) #using $Script:ScriptFileName allows the user to rename the file if they want
					Remove-Item -Path ($Script:WorkingDirectory + "\UpdateTemp\") -Recurse -Force #delete update temporary folder
					Write-Host " Updated :)" -foregroundcolor green
					Start-Sleep -milliseconds 850
					& ($Script:WorkingDirectory + "\" + $Script:ScriptFileName)
					exit
				}
			}
			$Script:CurrentStats.LastUpdateCheck = (get-date).tostring('yyyy.MM.dd HH:mm:ss')
			$Script:LatestVersionCheck = $CurrentStats.LastUpdateCheck
			$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update stats.csv with the new time played.
		}
		Catch {
			Write-Host "`n Couldn't check for updates. GitHub API limit may have been reached..." -foregroundcolor Yellow
			Start-Sleep -milliseconds 3500
		}
	}
	#Update (or replace missing) SetTextV2.bas file. This is an newer version of SetText (built by me and ChatGPT) that allows windows to be closed by process ID.
	if ((Test-Path -Path ($workingdirectory + '\SetText\SetTextv2.bas')) -ne $True){#if SetTextv2.bas doesn't exist, download it.
			try {
				New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") -ErrorAction stop | Out-Null #create temporary folder to download zip to and extract
			}
			Catch {#if folder already exists for whatever reason.
				Remove-Item -Path ($Script:WorkingDirectory + "\UpdateTemp\") -Recurse -Force
				New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") | Out-Null #create temporary folder to download zip to and extract
			}
			$Releases = Invoke-RestMethod -Uri "https://api.github.com/repos/shupershuff/D2rSPLoader/releases"
			$ReleaseInfo = ($Releases | Sort-Object id -desc)[0] #find release with the highest ID.
			$ZipURL = $ReleaseInfo.zipball_url #get zip download URL
			$ZipPath = ($WorkingDirectory + "\UpdateTemp\D2rSPLoader_" + $ReleaseInfo.tag_name + "_temp.zip")
			Invoke-WebRequest -Uri $ZipURL -OutFile $ZipPath
			if ($Null -ne $releaseinfo.assets.browser_download_url){#Check If I didn't forget to make a version.zip file and if so download it. This is purely so I can get an idea of how many people are using the script or how many people have updated. I have to do it this way as downloading the source zip file doesn't count as a download in github and won't be tracked.
				Invoke-WebRequest -Uri $releaseinfo.assets.browser_download_url -OutFile $null | out-null #identify the latest file only.
			}
			$ExtractPath = ($Script:WorkingDirectory + "\UpdateTemp\")
			Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
			$FolderPath = Get-ChildItem -Path $ExtractPath -Directory -Filter "shupershuff*" | Select-Object -ExpandProperty FullName
			Copy-Item -Path ($FolderPath + "\SetText\SetTextv2.bas") -Destination ($Script:WorkingDirectory + "\SetText\SetTextv2.bas")
			Write-Host "  SetTextV2.bas was missing and was downloaded."
			Remove-Item -Path ($Script:WorkingDirectory + "\UpdateTemp\") -Recurse -Force #delete update temporary folder
	}
}
Function ImportXML { #Import Config XML
	try {
		$Script:Config = ([xml](Get-Content "$Script:WorkingDirectory\Config.xml" -ErrorAction Stop)).D2SPLoaderConfig
		Write-Verbose "Config imported successfully."
	}
	Catch {
		Write-Host "`n Config.xml Was not able to be imported. This could be due to a typo or a special character such as `'&`' being incorrectly used." -foregroundcolor red
		Write-Host " The error message below will show which line in the config.xml is invalid:" -foregroundcolor red
		Write-Host (" " + $PSitem.exception.message + "`n") -foregroundcolor red
		PressTheAnyKeyToExit
	}
}
Function ValidationAndSetup {
	#
	#	Note to self, enter in any future additions/removals from config.xml here.
	#
	#Perform some validation on config.xml. Helps avoid errors for people who may be on older versions of the script and are updating. Will look to remove all of this in a future update.
	$Script:Config = ([xml](Get-Content "$Script:WorkingDirectory\Config.xml" -ErrorAction Stop)).D2SPLoaderConfig #import config.xml again for any updates made by the above.
	#check if there's any missing config.xml options, if so user has out of date config file.
	$AvailableConfigs = #add to this if adding features.
	"GamePath",
	"CustomLaunchArguments",
	"ShortcutCustomIconPath"
	$BooleanConfigs =
	"ManualSettingSwitcherEnabled",
	"DisableVideos",
	"AutoBackup",
	"CreateDesktopShortcut",
	"ForceWindowedMode"
	$AvailableConfigs = $AvailableConfigs + $BooleanConfigs
	$ConfigXMLlist = ($Config | Get-Member | Where-Object {$_.membertype -eq "Property" -and $_.name -notlike "#comment"}).name
	Write-Host
	ForEach ($Option in $AvailableConfigs){#Config validation
		if ($Option -notin $ConfigXMLlist){
			Write-Host " Config.xml file is missing a config option for $Option." -foregroundcolor yellow
			Start-Sleep 1
			PressTheAnyKey
		}
	}
	if ($Option -notin $ConfigXMLlist){
		Write-Host "`n Make sure to grab the latest version of config.xml from GitHub" -foregroundcolor yellow
		Write-Host " $X[38;2;69;155;245;4mhttps://github.com/shupershuff/D2rSPLoader/releases/latest$X[0m`n"
		PressTheAnyKey
	}
	if ($Config.GamePath -match "`""){#Remove any quotes from path in case someone ballses this up.
		$Script:GamePath = $Config.GamePath.replace("`"","")
	}
	else {
		$Script:GamePath = $Config.GamePath
	}
	ForEach ($ConfigCheck in $BooleanConfigs){#validate all configs that require "True" or "False" as the setting.
		if ($Null -ne $Config.$ConfigCheck -and ($Config.$ConfigCheck -ne $true -and $Config.$ConfigCheck -ne $false)){#if config is invalid
			Write-Host " Config option '$ConfigCheck' is invalid." -foregroundcolor Red
			Write-Host " Ensure this is set to either True or False.`n" -foregroundcolor Red
			PressTheAnyKeyToExit
		}
	}
	if ($Config.ShortcutCustomIconPath -match "`""){#Remove any quotes from path in case someone ballses this up.
		$ShortcutCustomIconPath = $Config.ShortcutCustomIconPath.replace("`"","")
	}
	else {
		$ShortcutCustomIconPath = $Config.ShortcutCustomIconPath
	}
	#Check Windows Game Path for D2r.exe is accurate.
	if ((Test-Path -Path "$GamePath\d2r.exe") -ne $True){
		Write-Host " Gamepath is incorrect. Looks like you have a custom D2r install location!" -foregroundcolor red
		Write-Host " Edit the GamePath variable in the config file.`n" -foregroundcolor red
		PressTheAnyKeyToExit
	}
	#Check Grail app path actually exists and if not throw an error
	if ("" -ne $Config.GrailAppExecutablePath){
		if ((Test-Path -Path $Config.GrailAppExecutablePath) -ne $true){ 
			Write-Host " Grail app '$(split-path $Config.GrailAppExecutablePath -leaf)' not found." -foregroundcolor red
			formatfunction -IsError -Text "Couldn't find the Grail application in '$(split-path $Config.GrailAppExecutablePath)'"
			Write-Host " Double check the grail application path and update config.xml to fix." -foregroundcolor red
			PressTheAnyKeyToExit
		}
}
	#Check Run Timer app path actually exists and if not throw an error
	if ("" -ne $Config.RunTimerAppExecutablePath){
		if ((Test-Path -Path $Config.RunTimerAppExecutablePath) -ne $true){ 
			Write-Host " Run Timer app '$(split-path $Config.RunTimerAppExecutablePath -leaf)' not found." -foregroundcolor red
			formatfunction -IsError -Text "Couldn't find the Run Timer application in '$(split-path $Config.RunTimerAppExecutablePath)'"
			Write-Host " Double check the Run Timer application path and update config.xml to fix." -foregroundcolor red
			PressTheAnyKeyToExit
		}
	}
	# Create Shortcut
	if ($Config.CreateDesktopShortcut -eq $True){
		$DesktopPath = [Environment]::GetFolderPath("Desktop")
		$Targetfile = "-ExecutionPolicy Bypass -File `"$WorkingDirectory\$ScriptFileName`""
		$ShortcutFile = "$DesktopPath\D2R Single Player.lnk"
		$WScriptShell = New-Object -ComObject WScript.Shell
		$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
		$Shortcut.TargetPath = "powershell.exe"
		$Shortcut.Arguments = $TargetFile
		if ($ShortcutCustomIconPath.length -eq 0){
			$Shortcut.IconLocation = "$Script:GamePath\D2R.exe"
		}
		Else {
			$Shortcut.IconLocation = $ShortcutCustomIconPath
		}
		$Shortcut.Save()
	}
	#Check if SetTextv2.exe exists, if not, compile from SetTextv2.bas. SetTextv2.exe is what's used to rename the windows.
	if ((Test-Path -Path ($workingdirectory + '\SetText\SetTextv2.exe')) -ne $True){ #-PathType Leaf check windows renamer is configured.
		Write-Host "`n First Time run!`n" -foregroundcolor Yellow
		Write-Host " SetTextv2.exe not in .\SetText\ folder and needs to be built."
		if ((Test-Path -Path "C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe") -ne $True){#check that .net4.0 is actually installed or compile will fail.
			Write-Host " .Net v4.0 not installed. This is required to compile the Window Renamer for Diablo." -foregroundcolor red
			Write-Host " Download and install it from Microsoft here:" -foregroundcolor red
			Write-Host " https://dotnet.microsoft.com/en-us/download/dotnet-framework/net40" #actual download link https://dotnet.microsoft.com/en-us/download/dotnet-framework/thank-you/net40-web-installer
			PressTheAnyKeyToExit
		}
		Write-Host " Compiling SetTextv2.exe from SetTextv2.bas..."
		& "C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe" -target:winexe -out:"`"$WorkingDirectory\SetText\SetTextv2.exe`"" "`"$WorkingDirectory\SetText\SetTextv2.bas`"" | out-null #/verbose  #actually compile the bastard
		if ((Test-Path -Path ($workingdirectory + '\SetText\SetTextv2.exe')) -ne $True){#if it fails for some reason and settextv2.exe still doesn't exist.
			Write-Host " SetTextv2 Could not be built for some reason :/"
			PressTheAnyKeyToExit
		}
		Write-Host " Successfully built SetTextv2.exe for Diablo 2 Launcher script :)" -foregroundcolor green
		Start-Sleep -milliseconds 4000 #a small delay so the first time run outputs can briefly be seen
	}
	#Check Handle64.exe downloaded and placed into correct folder
	$Script:WorkingDirectory = ((Get-ChildItem -Path $PSScriptRoot)[0].fullname).substring(0,((Get-ChildItem -Path $PSScriptRoot)[0].fullname).lastindexof('\'))
	if ((Test-Path -Path ($workingdirectory + '\Handle\Handle64.exe')) -ne $True){ #-PathType Leaf check windows renamer is configured.
		try {
			Write-Host "`n Handle64.exe not in .\Handle\ folder. Downloading now..." -foregroundcolor Yellow
			try {
				New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\Handle\ExtractTemp\") -ErrorAction stop | Out-Null #create temporary folder to download zip to and extract
			}
			Catch {#if folder already exists for whatever reason.
				Remove-Item -Path ($Script:WorkingDirectory + "\Handle\ExtractTemp\") -Recurse -Force
				New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\Handle\ExtractTemp\") | Out-Null #create temporary folder to download zip to and extract
			}
			$ZipURL = "https://download.sysinternals.com/files/Handle.zip" #get zip download URL
			$ZipPath = ($WorkingDirectory + "\Handle\ExtractTemp\")
			Invoke-WebRequest -Uri $ZipURL -OutFile ($ZipPath + "\Handle.zip")
			Expand-Archive -Path ($ZipPath + "\Handle.zip") -DestinationPath $ZipPath -Force
			Copy-Item -Path ($ZipPath + "Handle64.exe") -Destination ($Script:WorkingDirectory + "\Handle\")
			Remove-Item -Path ($Script:WorkingDirectory + "\Handle\ExtractTemp\") -Recurse -Force #delete update temporary folder
			Write-Host " Successfully downloaded Handle64.exe :)" -ForeGroundcolor Green
			Start-Sleep -milliseconds 2024
		}
		Catch {
			Write-Host " Handle.zip couldn't be downloaded." -foregroundcolor red
			FormatFunction -text "It's possible the download link changed. Try checking the Microsoft page or SysInternals.com site for a download link and ensure that handle64.exe is placed in the .\Handle\ folder." -IsError
			Write-Host "`n $X[38;2;69;155;245;4mhttps://learn.microsoft.com/sysinternals/downloads/handle$X[0m"
			Write-Host " $X[38;2;69;155;245;4mhttps://download.sysinternals.com/files/Handle.zip$X[0m`n"
			PressTheAnyKeyToExit
		}
	}
}
Function DisableVideos {
	#Diable Videos feature
	$VideoFiles = @(
		"blizzardlogos.webm",
		"d2intro.webm",
		"logoanim.webm",
		"d2x_intro.webm",
		"act2\act02start.webm",
		"act3\act03start.webm",
		"act4\act04start.webm",
		"act4\act04end.webm",
		"act5\d2x_out.webm"
	)
	if ($Config.DisableVideos -eq $True){
		if ($Config.CustomLaunchArguments -match "-mod"){
			$pattern = "-mod\s+(\S+)" #pattern to find the first word after -mod
			if ($Config.CustomLaunchArguments -match $pattern){
				$ModName = $matches[1]	
				$ModPath = $Config.GamePath + "\mods\$ModName\$ModName.mpq\data\hd\global\video"
				if (-not (Test-Path "$ModPath\blizzardlogos.webm")){
					Write-Host " You've opted to disable game videos however a mod is already being used." -ForegroundColor Yellow
					Do {
						Write-Host " Attempt to update the current mod ($ModName) to disable videos? $X[38;2;255;165;000;22mY$X[0m/$X[38;2;255;165;000;22mN$X[0m: " -nonewline
						$ShouldUpdate = ReadKey
						if ($ShouldUpdate -eq "y" -or $ShouldUpdate -eq "n"){
							$UpdateResponseValid = $True
							write-host
						}
						Else {
							Write-Host "`n Invalid response. Choose $X[38;2;255;165;000;22mY$X[0m $X[38;2;231;072;086;22mor$X[0m $X[38;2;255;165;000;22mN$X[0m.`n" -ForegroundColor red
						}
					} Until ($UpdateResponseValid -eq $True)
					if ($ShouldUpdate -eq "y"){	
						if (-not (Test-Path "$ModPath\act2")){
							Write-Debug " Creating folders required for disabling D2r videos..."
							New-Item -ItemType Directory -Path $ModPath -ErrorAction stop | Out-Null
							New-Item -ItemType Directory -Path "$ModPath\act2" -ErrorAction stop | Out-Null
							New-Item -ItemType Directory -Path "$ModPath\act3" -ErrorAction stop | Out-Null
							New-Item -ItemType Directory -Path "$ModPath\act4" -ErrorAction stop | Out-Null
							New-Item -ItemType Directory -Path "$ModPath\act5" -ErrorAction stop | Out-Null
							Write-Debug " Created folder: $ModPath"
							start-sleep -milliseconds 213
						}
						foreach ($File in $VideoFiles){
							New-Item -ItemType File -Path "$ModPath\$File" | Out-Null
						}
						Write-Debug " Created dummy D2r videos."
						start-sleep -milliseconds 222
					}
					else {
						Write-Host " D2r videos have not been disabled.`n" -foregroundcolor red
						start-sleep -milliseconds 256
					}
				}
			}
		}
		elseif ($Config.CustomLaunchArguments -match "-direct -txt"){ #if user has extracted files.
			foreach ($File in $VideoFiles){
				if ((Get-Item "$($Config.GamePath)\Data\hd\global\video\$File").Length -gt 0){ #check if file is larger than 0 bytes and if so backup original file and replace with 0 byte file.
					$FileName = $File -replace "^[^\\]+\\", "" #remove "act2\" from string if needed.
					try { #try renaming file. If it can't be renamed, it must already exist, therefore delete the video file.
						Rename-Item -Path "$($Config.GamePath)\Data\hd\global\video\$File" -NewName "$FileName.backup" -erroraction stop | Out-Null
					}
					Catch {
						Remove-Item -Path "$($Config.GamePath)\Data\hd\global\video\$File"
					}
					New-Item -ItemType File -Path "$($Config.GamePath)\Data\hd\global\video\$File" | Out-Null
				}
			}
		}
		else { #if user has not extracted files, launch with mod
			$ModPath = $Config.GamePath + "\mods\DisableVideos\DisableVideos.mpq\data\hd\global\video"
			if (-not (Test-Path "$ModPath")){
				Write-Host "  Creating needed folders for disabling D2r videos..."
				New-Item -ItemType Directory -Path $ModPath -ErrorAction stop | Out-Null
				New-Item -ItemType Directory -Path "$ModPath\act2" -ErrorAction stop | Out-Null
				New-Item -ItemType Directory -Path "$ModPath\act3" -ErrorAction stop | Out-Null
				New-Item -ItemType Directory -Path "$ModPath\act4" -ErrorAction stop | Out-Null
				New-Item -ItemType Directory -Path "$ModPath\act5" -ErrorAction stop | Out-Null
				start-sleep -milliseconds 213
				foreach ($File in $VideoFiles){
					$FileName = $File -replace "^[^\\]+\\", "" #remove "act2\" from string if needed.
					New-Item -ItemType File -Path "$ModPath\$File" | Out-Null
				}
				$data = @{
					name     = "DisableVideos"
					savepath = "../"
				}
				$json = $data | ConvertTo-Json -Depth 1 -Compress # Convert the hashtable to JSON
				Set-Content -Path ($Config.GamePath + "\mods\DisableVideos\DisableVideos.mpq\modinfo.json") -Value $json -Encoding UTF8 # Write the JSON content to the file
				Write-Host "  Created dummy D2r videos.`n" -ForegroundColor Green
				start-sleep -milliseconds 222
			}
			$Script:StartWithDisableVideosMod = $True
		}
	}
	else {
		Write-Debug "Videos are enabled"
		$Script:StartWithDisableVideosMod = $False
		if ($Config.CustomLaunchArguments -match "-direct -txt"){ #if user has extracted files.
			foreach ($File in $VideoFiles){
				if ((Get-Item "$($Config.GamePath)\Data\hd\global\video\$File").Length -eq 0){ #check if file is l0 bytes and if so remove 0 byte file and restore original file.
					$FileName = $File -replace "^[^\\]+\\", "" #remove "act2\" from string if needed.
					Remove-Item -Path "$($Config.GamePath)\Data\hd\global\video\$File"
					write-debug "removed $($Config.GamePath)\Data\hd\global\video\$File"
					Rename-Item -Path "$($Config.GamePath)\Data\hd\global\video\$File.backup" -NewName "$FileName" -erroraction stop | Out-Null
					write-debug "renamed $($Config.GamePath)\Data\hd\global\video\$File.backup"
				}
			}
		}
		elseif ($Config.CustomLaunchArguments -match "mod"){ #if user has extracted files.
			$pattern = "-mod\s+(\S+)" #pattern to find the first word after -mod
			if ($Config.CustomLaunchArguments -match $pattern){
				$ModName = $matches[1]	
				$ModPath = $Config.GamePath + "\mods\$ModName\$ModName.mpq\data\hd\global\video"
				if (-not(Test-Path "$ModPath\act2\act02start.webm.backup")){#figure out if we should try rename files from backup or just delete.
					$Replace = $True
				}
				foreach ($File in $VideoFiles){
					if ((Get-Item "$ModPath\$File").Length -eq 0){ #check if file is l0 bytes and if so remove 0 byte file and restore original file.
						$FileName = $File -replace "^[^\\]+\\", "" #remove "act2\" from string if needed.
						if ($Replace = $True){
							Rename-Item -Path "$ModPath\$File.backup" -NewName "$FileName" -erroraction stop | Out-Null
						}
						Else {
							Remove-Item -Path "$ModPath\$File"
							if ((Get-Item "$($Config.GamePath)\Data\hd\global\video\$File").Length -eq 0){#check to see if we can rely on original game files. If this is true, original files are empty and can't be used
								try {
									Rename-Item -Path "$ModPath\$File.backup" -NewName "$FileName" -erroraction stop | Out-Null #Attempt to see if there's a backup we can restore from in the mod folder. Unlikely.
								}
								Catch {
									formatfunction -IsError -indent 2 "Couldn't restore $($Config.GamePath)\Data\hd\global\video\$File.`nYou may need to repair your game from the Battlenet client."
								}
							}
						}
					}
				}
			}
		}
	}
}
Function ImportCSV { #Import Character CSV
	do {
		try {
			$Script:CharactersCSV = import-csv "$Script:WorkingDirectory\characters.csv" #import all characters from csv
		}
		Catch {
			FormatFunction -text "`ncharacters.csv does not exist. Make sure you create this file. Redownload from Github if needed." -IsError
			PressTheAnyKeyToExit
		}
		if ($Null -ne $Script:CharactersCSV){
			if ($Null -ne ($CharactersCSV | Where-Object {$_.CharacterName -eq ""})){
				$Script:CharactersCSV = $Script:CharactersCSV | Where-Object {$_.CharacterName -ne ""} # To account for user error, remove any empty lines from characters.csv
			}
			$CharacterCSVImportSuccess = $True
		}
		else {
			if (Test-Path ($Script:WorkingDirectory + "\characters.backup.csv")){ #Figure out if script is being run for first time by checking if characters.backup.csv doesn't exist, if so, don't try perform recovery.
				if ($CharacterCSVRecoveryAttempt -lt 1){#Error out and exit if there's a problem with the csv.
					try {
						Write-Host " Issue with characters.csv. Attempting Autorecovery from backup..." -foregroundcolor red
						Copy-Item -Path $Script:WorkingDirectory\characters.backup.csv -Destination $Script:WorkingDirectory\characters.csv -erroraction stop
						Write-Host " Autorecovery successful!" -foregroundcolor Green
						$CharacterCSVRecoveryAttempt ++
						PressTheAnyKey
					}
					Catch {
						$CharacterCSVImportSuccess = $False
					}
				}
				Else {
					$CharacterCSVRecoveryAttempt = 2
				}
				if ($CharacterCSVImportSuccess -eq $False -or $CharacterCSVRecoveryAttempt -eq 2){
					Write-Host "`n There's an issue with characters.csv." -foregroundcolor red
					Write-Host " Please ensure that this is filled out correctly and rerun the script." -foregroundcolor red
					Write-Host " Alternatively, rebuild CSV from scratch or restore from characters.backup.csv`n" -foregroundcolor red
					PressTheAnyKeyToExit
				}
			}
			else {
				$CharacterCSVImportSuccess = $True
			}
		}
	} until ($CharacterCSVImportSuccess -eq $True)
	$Script:CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
	([int]$Script:CurrentStats.TimesLaunched) ++
	if ($CurrentStats.TotalGameTime -eq ""){
		$Script:CurrentStats.TotalGameTime = 0 #prevents errors from happening on first time run.
	}
	try {
		$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv with Total Time played.
	}
	Catch {
		Write-Host "  Couldn't update stats.csv" -foregroundcolor yellow
	}
	#Make Backup of CSV.
	 # Added this in as I had BSOD on my PC and noticed that this caused the files to get corrupted.
	Copy-Item -Path ($Script:WorkingDirectory + "\stats.csv") -Destination ($Script:WorkingDirectory + "\stats.backup.csv")
	if ($Null -ne $Script:CharactersCSV){#Don't create backup csv if characters file isn't populated yet. prevents issues with first time run or running on a computer that doesnt have D2r char saves yet.
		Copy-Item -Path ($Script:WorkingDirectory + "\characters.csv") -Destination ($Script:WorkingDirectory + "\characters.backup.csv")
	}
}
Function CloudBackupSetup {
	<#
	Author: Shupershuff
	Version: 1.0
	Usage:
		Close D2r if open. Run script to move Diablo 2 save folder from "C:\Users\<USERNAME>\Saved Games\Diablo II Resurrected" to your chosen cloud storage.
	Purpose:
		Quick and easy way for folk to ensure game saves are saved to the cloud instead of local only.
		
	Instructions: See GitHub readme https://github.com/shupershuff/D2rSinglePlayerBackup
	Notes: 
	- Google Drive only works if it's configured to store files locally AND on the cloud. In other words, the "Mirror files" option is chosen instead of "Stream files"/
	- If ya want to do this yourself (for anything) without a script, just copy the data to a cloud sync'd path and use this command in CMD: mklink /J <DefaultSaveGamePath> <CloudSaveGamePath>
	#>
	##################
	# Config Options #
	##################
	$SaveFolderName = "Saved Games" #Name of the folder which will be created.
	##########
	# Script #
	##########
	write-host
	formatfunction -indents 1 "This will ensure your game files are saved in a cloud sync'd location."
	write-host
	$DefaultSaveGamePath = ($Script:CharacterSavePath + "\Diablo II Resurrected")
	$OneDriveSavePath = ("C:\Users\" + $env:username + "\OneDrive\")
	$DropboxSavePath = ("C:\Users\" + $env:username + "\Dropbox\")
	$GoogleDriveSavePath =  ("C:\Users\" + $env:username + "\My Drive\")
	###Check if junction has already been created ###
	$SavedGamesFolder = $Script:CharacterSavePath
	# Run cmd's dir command to get junction info, ensuring the path is quoted
	$junctionInfo = cmd /c "dir `"$SavedGamesFolder`" /AL"
	# Define a regex pattern to match the "Diablo II Resurrected" junction and its target path
	$regexPattern = "\s+<JUNCTION>\s+Diablo II Resurrected\s+\[([^\]]+)\]"
	# Check if the output contains the specific junction target path for "Diablo II Resurrected"
	if ("$junctionInfo" -match $regexPattern){ #Warn user if they've already ran this script and moved the game folder.
		# Extract the target path using the match
		$JunctionTarget = $matches[1]
		$RecreateJunction = $True
		formatfunction -indents 1 -IsWarning -text "Warning, the D2r Savegame folder is already redirected."
		formatfunction -indents 1 -IsWarning -text "This is currently pointing to: '$JunctionTarget'"
		Write-host
		Write-host "  If you're happy with game files already being saved to this cloud folder," -foregroundcolor yellow
		Write-host "  choose cancel ($X[38;2;255;165;000;22mc$X[0m" -nonewline -foregroundcolor yellow; write-host ")" -foregroundcolor yellow
		formatfunction -indents 1 -IsWarning -text "Otherwise, continue with the script to point it to the new cloud location."
		Write-Host "`n  Press '$X[38;2;255;165;000;22mc$X[0m' to cancel or any other key to proceed: "  -nonewline
		if (readkey -eq "c"){
			return $False
		}
		Write-host
	}
	else {
		Write-Verbose "No junction for 'Diablo II Resurrected' found in $SavedGamesFolder."
	}
	do {
		$LastOption = 3
		$OptionText = " or 3"
		write-host "`n  Options are:"
		write-host "   1 - OneDrive"
		write-host "   2 - Dropbox"
		write-host "   3 - Google Drive"
		if ($RecreateJunction -eq $True) {
			write-host "   4 - Move folder back to default (local) location"
			$LastOption ++
			$OptionText = ", 3 or 4"
		}
		write-host
		$CloudOption = [int](ReadKey "  Enter the option would you like to choose (1, 2$OptionText): ").tostring()	
		if ($CloudOption -notin 1..$LastOption){
			write-host "  Please choose option 1, 2$OptionText.`n" -foregroundcolor red
		}
	} until ($CloudOption -in 1..$LastOption)
	if ($CloudOption -eq 1){
		write-host "  Configuring for OneDrive...`n"
		$CloudSaveGamePath = ($OneDriveSavePath + $SaveFolderName + "\Diablo II Resurrected")
	}
	if ($CloudOption -eq 2){
		write-host "  Configuring for Dropbox...`n"
		$CloudSaveGamePath = ($DropboxSavePath + $SaveFolderName + "\Diablo II Resurrected")
	}
	if ($CloudOption -eq 3){
		write-host "  Configuring for Google Drive...`n"
		$CloudSaveGamePath = ($GoogleDriveSavePath + $SaveFolderName + "\Diablo II Resurrected")
		formatfunction -indents 1 -IsWarning "Note, for this to save to the cloud, you need to configure Google Drive to store files locally AND on the cloud"
		formatfunction -indents 1 -IsWarning "To do this, go into Google Drive preferences and change My Drive syncing options from 'Stream files' to 'Mirror files'.`n"
		PressTheAnyKey
	}
	if ($RecreateJunction -eq $True){
		write-verbose "Removing Existing Junction"
		Remove-Item -Path $DefaultSaveGamePath -recurse -Force
		New-Item $DefaultSaveGamePath -type directory | out-null
		write-verbose "Moving Savegame data from previous junction target back to default location"
		Get-ChildItem -Path $JunctionTarget | Copy-Item -Destination $DefaultSaveGamePath -Force -recurse
		if ($CloudOption -eq 4){
			Write-Host "  Moved Saved Game data back to default location.`n " -nonewline -foregroundcolor green
			$DefaultSaveGamePath
			Write-Host
			Return $True
		}
	}
	if (!(Test-Path -path $CloudSaveGamePath)) {
		New-Item $CloudSaveGamePath -type directory | out-null
		Write-host "  Created Directory: $CloudSaveGamePath" -foregroundcolor green
	}
	try {
		Get-ChildItem -Path $DefaultSaveGamePath | Copy-Item -Destination $CloudSaveGamePath -Force -recurse
		Remove-Item -Path $DefaultSaveGamePath -recurse -Force
		formatfunction -indents 1 -IsSuccess "Moved D2r Saves to $CloudSaveGamePath "
		cmd /c "mklink /J `"$DefaultSaveGamePath`" `"$CloudSaveGamePath`"" | out-null
		formatfunction -indents 1 -IsSuccess "Junction created in Saved Games folder to $CloudSaveGamePath."
		Write-host "`n  CloudBackup configured succesfully.`n" -foregroundcolor green
		return $True
	}
	Catch {
		Write-host "`n  Oh stink, something went wrong.`n" -foregroundcolor red
	}
}
Function LocalBackup {# Pillaged my own script but I'm lazy and used chatgpt to rewrite to account for folder exclusions https://github.com/shupershuff/FolderBackup
	Write-Host "  Backing up save games, please wait..." -foregroundcolor yellow
	CheckForModSavePath
	$PathToBackup = $Script:CharacterSavePath
	# Define the folders to exclude
	$ExcludedFolders = @("mods", "backup", "backups")
	# Initialize results array as a System.Collections.ArrayList for better performance
	$Results = [System.Collections.ArrayList]@()
	# Function to recursively collect files and directories while excluding specific folders
	function Get-FilteredItems {
		param (
			[string]$Path,
			[string[]]$ExcludedFolders
		)
		# Get all items in the current directory
		$Items = Get-ChildItem -Path $Path -Force
		# Add the current directory to results if it's not excluded
		if (-not ($ExcludedFolders -contains (Split-Path -Leaf $Path))) {
			$null = $Results.Add((Get-Item $Path))
		}
		foreach ($Item in $Items) {
			# Skip excluded folders
			if ($Item.PSIsContainer -and ($ExcludedFolders -contains $Item.Name)) {
				continue
			}
			# Add files directly to results
			if (-not $Item.PSIsContainer) {
				$null = $Results.Add($Item)
			}
			# If it's a directory, recurse into it
			elseif ($Item.PSIsContainer) {
				Get-FilteredItems -Path $Item.FullName -ExcludedFolders $ExcludedFolders
			}
		}
	}
	# Start collecting items from the root path
	Get-FilteredItems -Path $PathToBackup -ExcludedFolders $ExcludedFolders
	$PathToSaveBackup = Join-Path -Path $PathToBackup -ChildPath "Backups"
	if (-not (Test-Path $PathToSaveBackup)) {
		New-Item -ItemType Directory -Path $PathToSaveBackup -Force | Out-Null
	}
	# Helper function to calculate a folder hash. Prevents backup from rerunning and wasting time, storage and IO if no files have changed.
	Function Get-FolderHash {
		param ([string]$folderPath)
		# Initialize excluded folders (can be adjusted based on your needs)
		$ExcludedFolders = @("mods", "backup", "backups")
		# Get all files recursively, ensuring excluded folders are excluded
		$files = Get-ChildItem -Path $folderPath -Recurse -Force | Where-Object {
			# Check if any part of the full path is in the excluded list
			$exclude = $false
			# Split the full path into individual parts (folder and file name)
			$pathParts = $_.FullName.Split('\')
			# Check if any part of the path matches an excluded folder
			foreach ($part in $pathParts) {
				if ($ExcludedFolders -contains $part) {
					$exclude = $true
					break
				}
			}
			# Only include files that aren't in the excluded folders
			-not $exclude
		} | Sort-Object FullName
		$combinedHashes = ""
		# Calculate hash for each file
		foreach ($file in $files) {
			if (-not $file.PSIsContainer) {
				# Only hash files, not directories
				$fileHash = Get-FileHash -Path $file.FullName -Algorithm SHA256
				$combinedHashes += $fileHash.Hash
			}
		}
		# Final hash computation from all included files
		$finalHash = [System.BitConverter]::ToString((New-Object Security.Cryptography.SHA256Managed).ComputeHash([System.Text.Encoding]::UTF8.GetBytes($combinedHashes)))
		return $finalHash.Replace("-", "")
	}
	# Get the current date and time
	$currentDateTime = Get-Date
	$year = $currentDateTime.Year
	$month = $currentDateTime.ToString("MMMM")
	$day = $currentDateTime.ToString("dd")
	$hour = $currentDateTime.ToString("HHmm")
	$HashFilePath = Join-Path -Path $PathToSaveBackup -ChildPath "last_backup_hash.txt"
	$PreviousHash = if (Test-Path $HashFilePath) { Get-Content $HashFilePath } else { "" }
	$CurrentHash = Get-FolderHash -folderPath $PathToBackup
	if ($CurrentHash -eq $PreviousHash) {
		formatfunction -IsSuccess -indent 1 "Backup: Characters have not changed since the last backup. Backup skipped."	
		return "Skipped"
	}
	$destinationPath = Join-Path -Path $PathToSaveBackup -ChildPath "$year\$month\$day\$hour"
	if (-not (Test-Path $destinationPath)) {
		Write-Verbose "Backup: Creating Backup Folder in $PathToSaveBackup"
		New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
	}
	# Copy each item while respecting exclusions
	foreach ($Item in $Results) {
		$RelativePath = $Item.FullName.Substring($PathToBackup.Length).TrimStart('\')
		$Destination = Join-Path -Path $destinationPath -ChildPath $RelativePath
		if ($Item.PSIsContainer) {
			if (-not (Test-Path $Destination)) {
				New-Item -ItemType Directory -Path $Destination -Force | Out-Null
			}
		}
		else {
			$Dir = ([System.IO.Path]::GetDirectoryName($Destination))
			if (-not (Test-Path ([System.IO.Path]::GetDirectoryName($Destination)))) {
				New-Item -ItemType Directory -Path ([System.IO.Path]::GetDirectoryName($Destination)) -Force | Out-Null
			}
			Copy-Item -Path $Item.FullName -Destination $Destination -Force | Out-Null
		}
	}
	Write-verbose "Backup: Save Data copied to: $destinationPath" 
	write-host "  Backup: D2r Save Game data has been backed up :)" -ForegroundColor Green
	start-sleep 1
	$CurrentHash | Out-File -FilePath $HashFilePath -Force

	#Start Cleanup Tasks
	Write-verbose "Backup: Checking for old backups that can be cleaned up..."
	$DirectoryArray = New-Object -TypeName System.Collections.ArrayList
	Get-ChildItem -Path "$PathToSaveBackup\" -Directory -recurse -Depth 3 | Where-Object {$_.FullName -match '\\\d{4}\\\w+\\\d+\\\d{4}$'} | ForEach-Object {
		$DirectoryObject = New-Object -TypeName PSObject
		$pathComponents = $_.FullName -split '\\'
		$year = $pathComponents[-4]
		$month = $pathComponents[-3]
		$month = [datetime]::ParseExact($month, 'MMMM', $null).Month # convert month from text to number. EG February to 02
		$day = $pathComponents[-2]
		$time = $pathComponents[-1]
		$hour = $time[0]+$time[1]
		$minute = $time[2]+$time[3]
		$dateInFolder = Get-Date -Year $year -Month $month -Day $day -Hour $hour -minute $minute -second 00 #$minute can be changed to 00 if we want all the folders to be nicely named.
		$ShortFolderDate = (Get-Date -Year $year -Month $month -Day $day).ToString("d")
		Add-Member -InputObject $DirectoryObject -MemberType NoteProperty -Name FullPath -Value $_.FullName
		Add-Member -InputObject $DirectoryObject -MemberType NoteProperty -Name FolderDate -Value $dateInFolder
		Add-Member -InputObject $DirectoryObject -MemberType NoteProperty -Name ShortDate -Value $ShortFolderDate
		[VOID]$DirectoryArray.Add($DirectoryObject)
	}
	$DirectoryArray = $DirectoryArray | Sort-Object {[datetime]$_.FolderDate} -Descending
	$HourliesToKeep = $DirectoryArray | Group-Object -Property ShortDate | Select-Object -First 7 | select -expandproperty group #hourlies. These aren't necessarily hourly, can be taken every few minutes if desired
	$DailiesToKeep = $DirectoryArray | Group-Object -Property ShortDate | ForEach-Object { $_.Group[0] } | Select-Object -skip 7 -First 24 #this is actually useful for capturing the last backup of each day
	$MonthliesToKeep = $DirectoryArray | Group-Object -Property { ($_.ShortDate -split '/')[1] } | ForEach-Object { $_.Group[0] }
	#Perform steps to remove any old backups that aren't needed anymore. Keep all backups within last 7 days (even if last 7 days aren't contiguous). For the last 30 days, keep only the last backup taken on that day (Note that again, 30 days aren't necessarily contiguous). For all older backups, only keep the last backup taken that month.
	foreach ($Folder in $DirectoryArray){
		if ($MonthliesToKeep.FullPath -notcontains $Folder.FullPath -and $DailiesToKeep.FullPath -notcontains $Folder.FullPath -and $HourliesToKeep.FullPath -notcontains $Folder.FullPath){
			$Folder | Add-Member -MemberType NoteProperty -Name KeepFolder -Value "Deleted"
			Remove-Item -Path $Folder.FullPath -Recurse -Force
			Write-verbose "Backup: Tidied up by removing $($Folder.FullPath)"
			$Cleanup = $True
		}
		Else {
			$Folder | Add-Member -MemberType NoteProperty -Name KeepFolder -Value $True
		}
	}
	#Perform steps to Cleanup any empty directories.
	Function IsDirectoryEmpty($directory) { #Function to check each directory and subdirectory to determine if it's actually empty.
		$files = Get-ChildItem -Path $directory -File
		if ($files.Count -eq 0) { #directory has no files in it, checking subdirectories.
			$subdirectories = Get-ChildItem -Path $directory -Directory
			foreach ($subdirectory in $subdirectories) {
				if (-not (IsDirectoryEmpty $subdirectory.FullName)) {
					return $false #subdirectory has files in it
				}
			}
			return $true #directory is empty
		}
		return $false #directory has files in it.
	}
	$subdirectories = Get-ChildItem -Path $PathToSaveBackup -recurse -Directory
	foreach ($subdirectory in $subdirectories) {
		if (IsDirectoryEmpty $subdirectory.FullName) { # Check if the subdirectory is empty (no files)
			Remove-Item -Path $subdirectory.FullName -Force -Recurse # Remove the subdirectory
			Write-verbose "Backup: Deleted empty folder: $($subdirectory.FullName)"
			$Cleanup = $True
		}
	}
	if ($Cleanup -eq $True){
		Write-verbose "Backup: Backup cleanup complete."
	}
	Else {
		Write-verbose "Backup: No cleanup required."
	}
}
Function SetQualityRolls {
	#Set item quality array for randomizing quote colours. A stupid addition to script but meh.
	$Script:QualityArray = @(#quality and chances for things to drop based on 0MF values in D2r (I think?)
		[pscustomobject]@{Type='HighRune';Probability=1}
		[pscustomobject]@{Type='Unique';Probability=50}
		[pscustomobject]@{Type='SetItem';Probability=124}
		[pscustomobject]@{Type='Rare';Probability=200}
		[pscustomobject]@{Type='Magic';Probability=588}
		[pscustomobject]@{Type='Normal';Probability=19036}
	)
	if ($Script:GemActivated -eq $True){#small but noticeable MF boost
		$Script:QualityArray | Where-Object { $_.Type -eq 'Normal' } | ForEach-Object {
			$_.Probability = 16384  # New probability value
		}
	}
	Else {
		$Script:QualityArray | Where-Object { $_.Type -eq 'Normal' } | ForEach-Object {
			$_.Probability = 19036  # Original probability value
		}
	}
	if ($Script:CowKingActivated -eq $True){#big MF boost
		$Script:QualityArray | Where-Object { $_.Type -eq 'Normal' } | ForEach-Object {
			$_.Probability = 2048  # New probability value
			$Script:MOO = "MOO"
		}
	}
	if ($Script:PGemActivated -eq $True){#huuge MF boost
		$Script:QualityArray | Where-Object { $_.Type -eq 'Normal' } | ForEach-Object {
			$_.Probability = 192  # New probability value
		}
	}
	$QualityHash = @{};
	ForEach ($Object in $Script:QualityArray | select-object type,probability){#convert PSOobjects to hashtable for enumerator
		$QualityHash.add($Object.type,$Object.probability) #add each PSObject to hash
	}
	$Script:ItemLookup = ForEach ($Entry in $QualityHash.GetEnumerator()){
		[System.Linq.Enumerable]::Repeat($Entry.Key, $Entry.Value) #This creates a hash table with 19036 normal items, 588 magic items, 200 rare items etc etc. Used later as a list to randomly pick from.
	}
}
Function CowKingKilled {
	Write-Host "`n                          You Killed the Cow King!" -foregroundcolor green
	Write-Host "                                $X[38;2;165;146;99;22mMoo.$X[0m"
	Write-Host "                                    $X[38;2;165;146;99;22mMoooooooo!$X[0m"
	$voice = New-Object -ComObject Sapi.spvoice
	$voice.rate = -4 #How quickly the voice message should be
	$voice.volume = 50
	$voice.voice = $voice.getvoices() | Where-Object {$_.id -like "*David*"}
	$voice.speak("MooMoo-moo moo-moo, moo-moo, moo moo.") | out-null
	$CowLogo = @"
 -)[<-
   +[[[*
     *)}]>>=-----                   +:                   =-
      :=]#}[[[]))])>-               [>-     :------     =)*:
     :<]<<<][[]))][<*=              <]<>]]<>[%%%##}<><<<]]*
   :=*)))**)))[[]]>:                 -<)[}#%%%%%%%%%%%%[<
     -*><)]))))<)][[*:                    -[#%%%%%%%%%}-
       =*>><<<)<->[}}[*:                 *}#<>%%%)=)%%%}*
        :+>><)]])]]<[##[*:              *@@%%#[[[[}%%%%%}
          :-<<)))<)*  *#}[):            -}%@#[]]]]#%%%%%%>
              -=+><-   :>[#})+          :}%%#[[[]]###%%%%%>
                          -]#}}-:      -]%#}}[}}[[[[}#%%%%):
                            :>}})-    +#%%%#}}[][[)[}#%%%%#+
                               =##%%<=}}}}}}}}}}[[])}#%%%%%#+
                            :>[[##}}[[}[[[[[][]]]))))]]]]}#%%-
                            :]##}]<)][}[}#}}}}}[[[]])]][[}##%}:
                             =<}]>****+<[[}}}}[[[]]]]]][}}}###<
                                       :#%#[[[]]])))))][%%}[}}[-
                                       :#%###}[]]]))))][#%}]][]-
                                       :#%##}}}}[[[]]][[#}))<>-:
                                       :#####}}}}}[[[[]<>)}[:
                                         [#}}}}}}}}[}}##}%%+
                                         :#}}}}}}}}}[[[}}}}>-
                                         >##}}}}}}}}[[[[[}#[[<
                                        =}}}##}}}[[[[[[])]##<+
                                        >#}}#}}[[[[][]]]))#%*
                                       :]#}}###}[[[[[][[)]#[:
                                        *#}}#}}}}}}}[[[}[}#[:
          (__)                          *###}##}}}}}}}}}}#%):
          (oo)                          :[}}}}<]][[[}###}##+
   /-------\/                            :}}}}>    :)###}}#]-
  / |     ||                              +}}}>      +}%#}#%>
  * ||----||                              >}}[:        -=]}#<
    ||    ||                             -####*          )}]<:
                                        +#%%%}           -[]]>
                                        :-*+-:           -][])
                                                         =[]))+
                                                         =}##}):
"@
	Write-Host "  $X[38;2;255;165;0;22m$CowLogo$X[0m"
	$Script:CowKingActivated = $True
	([int]$Script:CurrentStats.CowKingKilled) ++
	SetQualityRolls
	start-sleep -milliseconds 4550
	try {
		$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv with Total Time played and cow stats.
	}
	Catch {
		Write-Host "  Couldn't update stats.csv" -foregroundcolor yellow
	}
}
Function HighRune {
	process { Write-Host "  $X[38;2;255;165;000;48;2;1;1;1;4m$_$X[0m"}
}
Function Unique {
	process { Write-Host "  $X[38;2;165;146;99;48;2;1;1;1;4m$_$X[0m"}
}
Function SetItem {
	process { Write-Host "  $X[38;2;0;225;0;48;2;1;1;1;4m$_$X[0m"}
}
Function Rare {
	process { Write-Host "  $X[38;2;255;255;0;48;2;1;1;1;4m$_$X[0m"}
}
Function Magic {#ANSI text colour formatting for "magic" quotes. The variable $X (for the escape character) is defined earlier in the script.
	process { Write-Host "  $X[38;2;65;105;225;48;2;1;1;1;4m$_$X[0m" }
}
Function Normal {
	process { Write-Host "  $X[38;2;255;255;255;48;2;1;1;1;4m$_$X[0m"}
}
Function QuoteRoll {#stupid thing to draw a random quote but also draw a random quality.
	$Quality = get-random $Script:ItemLookup #pick a random entry from ItemLookup hashtable.
	Write-Host
	$LeQuote = (Get-Random -inputobject $Script:quotelist) #pick a random quote.
	$ConsoleWidth = $Host.UI.RawUI.BufferSize.Width
	$DesiredIndent = 2  # indent spaces
	$ChunkSize = $ConsoleWidth - $DesiredIndent
	[RegEx]::Matches($LeQuote, ".{$ChunkSize}|.+").Groups.Value | ForEach-Object {
		Write-Output $_ | &$Quality #write the quote and write it in the quality colour
	}
	if ($LeQuote -match "Moo" -and $Quality -eq "Unique"){
		CowKingKilled
	}
	$Script:CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
	if ($Quality -eq "HighRune"){([int]$Script:CurrentStats.HighRunesFound) ++}
	if ($Quality -eq "Unique"){([int]$Script:CurrentStats.UniquesFound) ++}
	if ($Quality -eq "SetItem"){([int]$Script:CurrentStats.SetItemsFound) ++}
	if ($Quality -eq "Rare"){([int]$Script:CurrentStats.RaresFound) ++}
	if ($Quality -eq "Magic"){([int]$Script:CurrentStats.MagicItemsFound) ++}
	if ($Quality -eq "Normal"){([int]$Script:CurrentStats.NormalItemsFound) ++}
	try {
		$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv
	}
	Catch {
		Write-Host "  Couldn't update stats.csv" -foregroundcolor yellow
		Start-Sleep -Milliseconds 256
	}
}
Function Inventory {#Info screen
	Clear-Host
	Write-Host "`n          Stay a while and listen! Here's your D2r Loader info.`n`n" -foregroundcolor yellow
	Write-Host "  $X[38;2;255;255;255;4mNote:$X[0m D2r Playtime is based on the time the script has been running"
	Write-Host "  whilst D2r is running. In other words, if you use this script when you're"
	Write-Host "  playing the game, it will give you a reasonable idea of the total time"
	Write-Host "  you've spent receiving disappointing drops from Mephisto :)`n"
	$QualityArraySum = 0
	$Script:QualityArray | ForEach-Object {
		$QualityArraySum += $_.Probability
	}
	$NormalProbability = ($QualityArray | where-object {$_.type -eq "Normal"} | Select-Object Probability).probability
	$Script:CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
	$Line1 =   "                    ----------------------------------"
	$Line2 =  ("                   |  $X[38;2;255;255;255;22mD2r Playtime (Hours):$X[0m " +  ((($time =([TimeSpan]::Parse($CurrentStats.TotalGameTime))).hours + ($time.days * 24)).tostring() + ":" + ("{0:D2}" -f $time.minutes)))
	$Line3 =  ("                   |  $X[38;2;255;255;255;22mCurrent Session (Hours):$X[0m" + ((($time =([TimeSpan]::Parse($Script:SessionTimer))).hours + ($time.days * 24)).tostring() + ":" + ("{0:D2}" -f $time.minutes)))
	$Line4 =  ("                   |  $X[38;2;255;255;255;22mScript Launch Counter:$X[0m " + $CurrentStats.TimesLaunched)
	$Line5 =   "                    ----------------------------------"
	$Line6 =  ("                   |  $X[38;2;255;165;000;22mHigh Runes$X[0m Found: " + $(if ($CurrentStats.HighRunesFound -eq ""){"0"} else {$CurrentStats.HighRunesFound}))
	$Line7 =  ("                   |  $X[38;2;165;146;99;22mUnique$X[0m Quotes Found: " + $(if ($CurrentStats.UniquesFound -eq ""){"0"} else {$CurrentStats.UniquesFound}))
	$Line8 =  ("                   |  $X[38;2;0;225;0;22mSet$X[0m Quotes Found: " + $(if ($CurrentStats.SetItemsFound -eq ""){"0"} else {$CurrentStats.SetItemsFound}))
	$Line9 =  ("                   |  $X[38;2;255;255;0;22mRare$X[0m Quotes Found: " + $(if ($CurrentStats.RaresFound -eq ""){"0"} else {$CurrentStats.RaresFound}))
	$Line10 = ("                   |  $X[38;2;65;105;225;22mMagic$X[0m Quotes Found: " + $(if ($CurrentStats.MagicItemsFound -eq ""){"0"} else {$CurrentStats.MagicItemsFound}))
	$Line11 = ("                   |  $X[38;2;255;255;255;22mNormal$X[0m Quotes Found: " + $(if ($CurrentStats.NormalItemsFound -eq ""){"0"} else {$CurrentStats.NormalItemsFound}))
	$Line12 =  "                    ----------------------------------"
	$Line13 = ("                   |  $X[38;2;165;146;99;22mCow King Killed:$X[0m " + $(if ($CurrentStats.CowKingKilled -eq ""){"0"} else {$CurrentStats.CowKingKilled}))
	$Line14 = ("                   |  $X[38;2;255;0;255;22mGems Activated:$X[0m  " + $(if ($CurrentStats.Gems -eq ""){"0"} else {$CurrentStats.Gems}))
	$Line15 = ("                   |  $X[38;2;255;0;255;22mPerfect Gem Activated:$X[0m " + $(if ($CurrentStats.PerfectGems -eq ""){"0"} else {$CurrentStats.PerfectGems}))
	$Line16 =  "                    ----------------------------------"
	$Lines = @($Line1,$Line2,$Line3,$Line4,$Line5,$Line6,$Line7,$Line8,$Line9,$Line10,$Line11,$Line12,$Line13,$Line14,$Line15,$Line16)
	# Loop through each object in the array to find longest line (for formatting)
	ForEach ($Line in $Lines){
		if (($Line -replace '\[.*?22m', '' -replace '\[0m','').Length -gt $LongestLine){
			$LongestLine = ($Line -replace '\[.*?22m', '' -replace '\[0m','').Length
		}
	}
	ForEach ($Line in $Lines){#Formatting nonsense to indent things nicely
		$Indent = ""
		$Dash = ""
		if (($Line -replace '\[.*?22m', '' -replace '\[0m','').Length -lt $LongestLine + 2){
			if ($Line -notmatch "-"){
				while ((($Line -replace '\[.*?22m', '' -replace '\[0m','').Length + $Indent.length) -lt ($LongestLine + 2)){
					$Indent = $Indent + " "
				}
				Write-Host $Line.replace(":$X[0m",":$X[0m$Indent").replace(" Found:"," Found:$Indent") -nonewline
				Write-Host "  |" -nonewline
				Write-Host
			}
			else {
				while (($Line.Length + $Dash.length) -le ($LongestLine +1)){
					$Dash = $Dash + "-"
				}
				Write-Host $Line -nonewline
				Write-Host $Dash
			}
		}
		Else {
			Write-Host " |"
		}
	}
	Write-Host ("`n  Chance to find $X[38;2;65;105;225;22mMagic$X[0m quality quote or better: " + [math]::Round((($QualityArraySum - $NormalProbability + 1) * (1/$QualityArraySum) * 100),2) + "%" )
	Write-Host ("`n  $X[4mD2r Game Version:$X[0m    " + (Get-Command "$GamePath\D2R.exe").FileVersionInfo.FileVersion)
	Write-Host "  $X[4mScript Install Path:$X[0m " -nonewline
	Write-Host ("`"$Script:WorkingDirectory`"" -replace "((.{1,52})(?:\\|\s|$)|(.{1,53}))", "`n                        `$1").trim() #add two spaces before any line breaks for indenting. Add line break for paths that are longer than 53 characters.
	Write-Host "  $X[4mYour Script Version:$X[0m v$CurrentVersion"
	Write-Host "  $X[38;2;69;155;245;4mhttps://github.com/shupershuff/D2rSPLoader/releases/v$CurrentVersion$X[0m"
	if ($null -eq $Script:LatestVersionCheck -or $Script:LatestVersionCheck.tostring() -lt (Get-Date).addhours(-2).ToString('yyyy.MM.dd HH:mm:ss')){ #check for updates. Don't check if this has been checked in the couple of hours.
		try {
			$Releases = Invoke-RestMethod -Uri "https://api.github.com/repos/shupershuff/D2rSPLoader/releases"
			$ReleaseInfo = ($Releases | Sort-Object id -desc)[0] #find release with the highest ID.
			$Script:LatestVersionCheck = (get-date).tostring('yyyy.MM.dd HH:mm:ss')
			$Script:LatestVersion = [version[]]$ReleaseInfo.Name.Trim('v')
		}
		Catch {
			Write-Output "  Couldn't check for updates :(" | Red
		}
	}
	if ($Null -ne $Script:LatestVersion -and $Script:LatestVersion -gt $Script:CurrentVersion){
		Write-Host "`n  $X[4mLatest Script Version:$X[0m v$LatestVersion" -foregroundcolor yellow
		Write-Host "  $X[38;2;69;155;245;4mhttps://github.com/shupershuff/D2rSPLoader/releases/latest$X[0m"
	}
	Write-Host "`n  $X[38;2;0;225;0;22mConsider donating as a way to say thanks via an option below:$X[0m"
	Write-Host "    - $X[38;2;69;155;245;4mhttps://www.buymeacoffee.com/shupershuff$X[0m"
	Write-Host "    - $X[38;2;69;155;245;4mhttps://paypal.me/Shupershuff$X[0m"
	Write-Host "    - $X[38;2;69;155;245;4mhttps://github.com/sponsors/shupershuff?frequency=one-time&amount=5$X[0m`n"
	if ($Script:NotificationsAvailable -eq $True){
		Write-Host "  -------------------------------------------------------------------------"
		Write-Host "  $X[38;2;255;165;000;48;2;1;1;1;4mNotification:$X[0m" -nonewline
		Notifications -check $False
		$Script:NotificationHasBeenChecked = $True
		Write-Host "  -------------------------------------------------------------------------"
	}
	Write-Host
	PressTheAnyKey
}
Function CheckForModSavePath {
	param (
		[switch] $Settings,
		[switch] $CheckOnly
	)
	if ($Settings -eq $True){
		$SettingsOrCharString = "settings"
		$SettingsOrCharString2 = "settings.json"
	}
	else {
		$SettingsOrCharString = "character"
		$SettingsOrCharString2 = "character saves"
	}
	if ($Config.CustomLaunchArguments -match "-mod"){
		$pattern = "-mod\s+(\S+)" #pattern to find the first word after -mod
		if ($Config.CustomLaunchArguments -match $pattern){
			$ModName = $matches[1]
			try {
				Write-Verbose "Trying to get Mod Content..."
				try {
					$Modinfo = ((Get-Content "$($Config.GamePath)\Mods\$ModName\$ModName.mpq\Modinfo.json" -ErrorAction silentlycontinue | ConvertFrom-Json).savepath).Trim("/")
				}
				catch {
					try {
						$Modinfo = ((Get-Content "$($Config.GamePath)\Mods\$ModName\Modinfo.json" -ErrorAction stop -ErrorVariable ModReadError | ConvertFrom-Json).savepath).Trim("/")
					}
					catch {
						FormatFunction -Text "Using standard $SettingsOrCharString save path. Couldn't find Modinfo.json in '$($Config.GamePath)\Mods\$ModName\$ModName.mpq'" -IsWarning
						start-sleep -milliseconds 1500
					}
				}
				If ($Null -eq $Modinfo){
					Write-Verbose " No Custom Save Path Specified for this mod."
				}
				ElseIf ($Modinfo -ne "../"){
					$Script:CharacterSavePath += "mods\$Modinfo\"
					$Script:SettingsProfilePath = $Script:CharacterSavePath
					if ($CheckOnly -ne $True){
						if (-not (Test-Path $CharacterSavePath)){
							Write-Host "  Mod Save Folder doesn't exist yet. Creating folder..."
							New-Item -ItemType Directory -Path $CharacterSavePath -ErrorAction stop | Out-Null
							Write-Host "  Created folder: $CharacterSavePath" -ForegroundColor Green
						}
						Write-Host "  Mod: '$ModName' detected. Using custom path for $SettingsOrCharString2." -ForegroundColor Green
					}
					Write-Verbose " $CharacterSavePath"
				}
				Else {
					Write-Verbose "  Mod used but save path is standard."
				}
			}
			Catch {
				Write-Verbose "  Mod used but custom save path not specified."
			}
		}
		else {
			Write-Host "  Couldn't detect Mod name. Standard path to be used for $SettingsOrCharString2." -ForegroundColor Red
		}
	}	
}
Function Options {
	ImportXML
	Clear-Host
	Write-Host "`n This screen allows you to change script config options."
	Write-Host " Note that you can also change these settings (and more) in config.xml."
	Write-Host " Options you can change/toggle below:"
	CheckForModSavePath -CheckOnly
	# Get all directories in the specified path, excluding "mods"
	$D2rDirectories = Get-ChildItem -Path $CharacterSavePath -Directory | Where-Object {$_.Name -ne "mods" -and $_.Name -ne "backup" -and $_.Name -ne "backups"}
	$NonEmptyDirectories = @()
	foreach ($Directory in $D2rDirectories){
		if ((Get-ChildItem -Path $directory.FullName -File -Filter "*.d2s").Count -eq 0){#Check if there are any folders with no .d2s files. If there is an 'empty' folder, this must be the Character set currently in use.
			$CurrentProfile = $directory.Name
		}
		else {
			$NonEmptyDirectories += $Directory
		}
	}
	if ($Null -eq $CurrentProfile){
			$CurrentProfile = "Main Profile" #If no empty folders
			if (-not (Test-Path "$CharacterSavePath\$CurrentProfile")){
				New-Item -ItemType Directory -Path "$CharacterSavePath\$CurrentProfile" -ErrorAction stop | Out-Null
			}
	}
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	write-host "`n  ------------------------------------------------------------------------- "
	FormatFunction -indents 1 -SubsequentLineIndents 4 -text "Launch Arguments: (Currently '$X[38;2;255;165;000;22m$(if($Script:Config.CustomLaunchArguments -ne ''){$Script:Config.CustomLaunchArguments}else{'No launch parameters configured'})$X[0m')"
	FormatFunction -indents 1 -SubsequentLineIndents 4 -text " $X[38;2;255;165;000;22m1$X[0m - $X[4m-seed - Launch with map seed$X[0m"
	FormatFunction -indents 1 -SubsequentLineIndents 4 -text " $X[38;2;255;165;000;22m2$X[0m - $X[4m-playersX - Set default players value on launch$X[0m"
	FormatFunction -indents 1 -SubsequentLineIndents 4 -text " $X[38;2;255;165;000;22m3$X[0m - $X[4m-enablerespec - Enable Infinite Respecs$X[0m"
	FormatFunction -indents 1 -SubsequentLineIndents 4 -text " $X[38;2;255;165;000;22m4$X[0m - $X[4m-resetofflinemaps - Enable rolling new maps on each new game$X[0m"
	write-host "  ------------------------------------------------------------------------- "
	write-host "  Other Options: "
	Write-Host "   $X[38;2;255;165;000;22m5$X[0m - $X[4mShow/Hide Characters on main screen$X[0m" #Whitelist chars for main screen
	Write-Host "   $X[38;2;255;165;000;22m6$X[0m - $X[4mManualSettingSwitcherEnabled$X[0m (Currently $X[38;2;255;165;000;22m$(if($Script:Config.ManualSettingSwitcherEnabled -eq 'True'){'Enabled'}else{'Disabled'})$X[0m)"
	Write-Host "   $X[38;2;255;165;000;22m7$X[0m - $X[4mDisableVideos$X[0m (Currently $X[38;2;255;165;000;22m$($Script:Config.DisableVideos)$X[0m)"
	Write-Host "   $X[38;2;255;165;000;22m8$X[0m - $X[4mBack Up Saved Game Folder$X[0m"
	if ($null -ne $D2rDirectories){
		Write-Host "   $X[38;2;255;165;000;22m9$X[0m - $X[4mSwap Character Packs$X[0m (Current Profile: $X[38;2;255;165;000;22m$CurrentProfile$X[0m)"
	}
	Write-Host "`n Enter one of the above options to change the setting."
	Write-Host " Otherwise, press any other key to return to main menu... " -nonewline
	$Option = readkey
	Write-Host;Write-Host
	Function OptionSubMenu {
		param (
			[String]$Description,
			[hashtable]$OptionsList,
			[String]$OptionsText,
			[String]$Current,
			[String]$ConfigName,
			[switch]$OptionInteger
		)
		$XML = Get-Content "$Script:WorkingDirectory\Config.xml" -Raw
		FormatFunction -indents 1 -text "Changing setting for $X[4m$($ConfigName)$X[0m (Currently $X[38;2;255;165;000;22m$($Current)$X[0m).`n"
		FormatFunction -text $Description -indents 1
		Write-Host;
		FormatFunction -indents 3 $OptionsText
		do {
			if ($OptionInteger -eq $True){
				Write-Host "   Enter a number between $X[38;2;255;165;000;22m1$X[0m and $X[38;2;255;165;000;22m99$X[0m or '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
				$AcceptableOptions = 1..99 # Allow user to enter 1 to 99
				$NewOptionValue = (ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27 -TwoDigitAcctSelection $True).tostring()
				if ($NewOptionValue -notin $AcceptableOptions + "c" + "Esc"){
					Write-Host "   Invalid Input. Please enter one of the options above.`n" -foregroundcolor red
				}
				else {
					$NewValue = $NewOptionValue
				}
			}
			Else {
				Write-Host "   Enter " -nonewline;CommaSeparatedList -NoOr ($OptionsList.keys | sort-object); Write-Host " or '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
				$AcceptableOptions = $OptionsList.keys
				$NewOptionValue = (ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27).tostring()
				$NewValue = $($OptionsList[$NewOptionValue])
				if ($NewOptionValue -eq "c"){
					break
				}
				if ($NewOptionValue -notin $AcceptableOptions + "Esc"){
					Write-Host "   Invalid Input. Please enter one of the options above.`n" -foregroundcolor red
				}
				if ($Option -in @("1","2","3","4")){# if option is for Changing custom launch parameters, we need a sub option to obtain seed number.
					$customLaunchArguments = $config.SelectSingleNode("//CustomLaunchArguments")	
					if ($Option -in @("1","2")){
						Function GetNumber {
							if ($Option -eq "1"){
								$Type = "seed"
								$Message = "Enter the seed you want to use"
								$RangeMax = 4294967294
							}
							Elseif ($Option -eq "2"){
								$Type = "player count"
								$Message = "Enter the players count you want the game to default to"
								$RangeMax = 8
							}
							do {
								$NumberInput = Read-Host "   Enter a number between 1 and $RangeMax"
								$ParsedNumber = $Null
								if ([long]::TryParse($NumberInput, [ref]$ParsedNumber) -and $ParsedNumber -ge 1 -and $ParsedNumber -le $RangeMax) {# Try to parse the input as an integer and ensure it's valid input (a number) within the range
									break
								}
								else {
									Write-Host "   Invalid $type. Please enter a valid number between 1 and $RangeMax." -ForegroundColor Red
								}
							} while ($true)
							$Parameter = [string]$($OptionsList[$NewOptionValue])
							return "$Parameter$ParsedNumber"
						}
					}
					if ($customLaunchArguments.InnerText -match $pattern){ #Remove/Edit config - If current config matches pattern, it must already exist.
						$newSubValue = ""
						if ($Option -in @("1","2")){
							if ($NewOptionValue -eq "2"){ #if sub menu option for seed or player count.
								$newSubValue = GetNumber
							}
						}
						$customLaunchArguments.InnerText = $customLaunchArguments.InnerText -replace $pattern, $newSubValue
					}
					else { #Add config
						if ($Option -in @("1","2")){
							$newSubValue = GetNumber
						}
						else {
							$newSubValue = [string]$($OptionsList[$NewOptionValue])
						}
						$customLaunchArguments.InnerText += " " + $newSubValue
					}
					$NewValue = ($customLaunchArguments.InnerText -replace "  ", " ").trim()
				}
				ElseIf ($Option -eq "8"){
					if ($NewOptionValue -eq "2"){
						if (LocalBackup -eq "Skipped"){
							start-sleep -milliseconds 2800 #If it was skipped allow short time for the message to appear before refreshing screen
						}
						return $False
					}
					elseif ($NewOptionValue -eq "3"){
						if (CloudBackupSetup -eq $True){
							PressTheAnyKey
						}
						return $false
					}
				}
			}
		} until ($NewOptionValue -in $AcceptableOptions + "c" + "Esc")
		if ($NewOptionValue -in $AcceptableOptions){
			try {
				$Pattern = "(<$ConfigName>)([^<]*)(</$ConfigName>)"
				$ReplaceString = '{0}{1}{2}' -f '${1}', $NewValue, '${3}'
				$NewXML = ([regex]::Replace($Xml, $Pattern, $ReplaceString)).trim()
				$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
				return $True
			}
			Catch {
				Write-Host "`n  Was unable to update config :(" -foregroundcolor red
				start-sleep -milliseconds 2500
				return $False
			}
		}
		else {
			Return $False
		}
	} #end of OptionSubMenu function
	If ($Option -eq "1"){ #custom map seed - SinglePlayerLaunch Options to add/remove from custom launch arguments
		$Pattern = "-seed \d{1,}" # Matches "-seed " followed by one or more digits
		If ($Script:Config.CustomLaunchArguments -notmatch "-seed "){
			$Options = @{"1" = "-seed "}
			$OptionsSubText = "enable"
			$CurrentState = "$($Script:Config.CustomLaunchArguments)"
		}
		Else {
			$Options = @{"1" = "PlaceholderValue Only :)";"2" = "-seed "}	
			if ($Script:Config.CustomLaunchArguments -match "-seed (\d+)"){
				$Seed = $matches[1]
			}
			$ExtraOptionsText = "Choose '$X[38;2;255;165;000;22m2$X[0m' to change seed (Currently $X[38;2;255;165;000;22m$Seed$X[0m)`n"
			$OptionsSubText = "disable"
			$CurrentState = "$($Script:Config.CustomLaunchArguments)"
		}
		$XMLChanged = OptionSubMenu -ConfigName "CustomLaunchArguments" -OptionsList $Options -Current $CurrentState `
		-Description "Choose to $OptionsSubText launching with a specified map seed for the game." `
		-OptionsText "Choose '$X[38;2;255;165;000;22m1$X[0m' to $OptionsSubText launching with a custom map seed.`n$ExtraOptionsText"
	}
	If ($Option -eq "2"){ #playersX - SinglePlayerLaunch Options to add/remove from custom launch arguments
		$Pattern = "-players \d{1,}" # Matches "-players " followed by one or more digits
		If ($Script:Config.CustomLaunchArguments -notmatch "-players"){
			$Options = @{"1" = "-players "}
			$OptionsSubText = "enable"
			$CurrentState = "$($Script:Config.CustomLaunchArguments)"
		}
		Else {
			$Options = @{"1" = "PlaceholderValue Only :)";"2" = "-players "}	
			if ($Script:Config.CustomLaunchArguments -match "-players (\d+)"){
				$PlayerCount = $matches[1]
			}
			$ExtraOptionsText = "Choose '$X[38;2;255;165;000;22m2$X[0m' to change the player count (Currently $X[38;2;255;165;000;22m$PlayerCount$X[0m)`n"
			$OptionsSubText = "disable"
			$CurrentState = "$($Script:Config.CustomLaunchArguments)"
		}
		$XMLChanged = OptionSubMenu -ConfigName "CustomLaunchArguments" -OptionsList $Options -Current $CurrentState `
		-Description "Choose to $OptionsSubText launching with the game with /playersX already set. Saves you having to type out the same thing at launch :)" `
		-OptionsText "Choose '$X[38;2;255;165;000;22m1$X[0m' to $OptionsSubText launching with /playersX.`n$ExtraOptionsText"
	}
	If ($Option -eq "3"){ #enablerespec - SinglePlayerLaunch Options to add/remove from custom launch arguments
		$Pattern = "-enablerespec" # Matches "-enablerespec"
		if ($Script:Config.CustomLaunchArguments -notmatch "-enablerespec"){
			$Options = @{"1" = "-enablerespec"}
			$OptionsSubText = "enable"
			$CurrentState = "$($Script:Config.CustomLaunchArguments)"
		}
		Else {
			$Options = @{"1" = "PlaceholderValue Only :)"}
			$OptionsSubText = "disable"
			$CurrentState = "$($Script:Config.CustomLaunchArguments)"
		}
		$XMLChanged = OptionSubMenu -ConfigName "CustomLaunchArguments" -OptionsList $Options -Current $CurrentState `
		-Description "Choose to $OptionsSubText launching with the ability to respec an unlimited amount of times." `
		-OptionsText "Choose '$X[38;2;255;165;000;22m1$X[0m' to $OptionsSubText infinite respecs.`n"
	}
	If ($Option -eq "4"){ #resetofflinemaps  - SinglePlayerLaunch Options to add/remove from custom launch arguments
		$Pattern = "-resetofflinemaps" # Matches "-resetofflinemaps "
		if ($Script:Config.CustomLaunchArguments -notmatch "-resetofflinemaps"){
			$Options = @{"1" = "-resetofflinemaps"}
			$OptionsSubText = "enable"
			$CurrentState = "$($Script:Config.CustomLaunchArguments)"
		}
		Else {
			$Options = @{"1" = "PlaceholderValue Only :)"}
			$OptionsSubText = "disable"
			$CurrentState = "$($Script:Config.CustomLaunchArguments)"
		}
		$XMLChanged = OptionSubMenu -ConfigName "CustomLaunchArguments" -OptionsList $Options -Current $CurrentState `
		-Description "Choose to $OptionsSubText launching with the game where the maps are reset/rerolled for each new game you make (same that happens with online play)." `
		-OptionsText "Choose '$X[38;2;255;165;000;22m1$X[0m' to generate new maps on joining the game (same behaviour as online).`n"
	}
	If ($Option -eq "5"){ #Whitelist
		$Script:CharactersCSV = @(Import-Csv -Path "$Script:WorkingDirectory\characters.csv")
		$LongestCharNameLength = $Script:CharactersCSV | ForEach-Object { $_.CharacterName.Length } | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum #find out how many batches there are so table can be properly indented.
		Do {
			$HeaderIndent += " "
		} Until ($HeaderIndent.length -ge ($LongestCharNameLength - 14))
		if (($Script:CharactersCSV.count) -ge 10){$IDIndent = " "}
		if (($Script:CharactersCSV.count) -ge 100){$IDIndent = "  "}
		
		formatfunction -indent 1 -text "On this screen you can toggle which characters you want to have shown on the scripts main display screen.`nYou can also edit this directly (better for bulk editgs) by editing the characters.csv file."
		write-host "`n  $X[4m#$X[0m $IDIndent  $X[4mCharacter Name$X[0m  $HeaderIndent $X[4mShow/Hide$X[0m   |   $X[4m#$X[0m $IDIndent  $X[4mCharacter Name$X[0m  $HeaderIndent $X[4mShow/Hide$X[0m"
		$Counter = 0
		$DoCount = 0
		foreach ($Character in $Script:CharactersCSV){
			$Character | Add-Member -NotePropertyName TemporaryID -NotePropertyValue (++$Counter)
			$Indent = ""
			$CountIndent = ""
			$DoCount++
			Do {
				$Indent += " "
			} Until (($Indent.length + $Character.CharacterName.length) -ge $LongestCharNameLength + 1)
			if ($Counter -le 9)     {$CountIndent = "  "}
			elseif ($Counter -le 99){$CountIndent = " "}
			elseif ($Counter -gt 99){$CountIndent = ""}
			if ($DoCount -eq 1){
				write-host "  $Counter $CountIndent $($Character.CharacterName) $Indent " -nonewline
				if( $Character.ShowOnDisplayScreen -eq 'Show'){
					write-host "$X[38;2;000;255;000;22m$($Character.ShowOnDisplayScreen)$X[0m" -nonewline
				}
				Else {
					write-host "$X[38;2;255;000;000;22m$($Character.ShowOnDisplayScreen)$X[0m" -nonewline
				}
			}
			Else {
				write-host "        |   $Counter $CountIndent $($Character.CharacterName) $Indent " -nonewline
				if ($Character.ShowOnDisplayScreen -eq 'Show'){
					write-host "$X[38;2;000;255;000;22m$($Character.ShowOnDisplayScreen)$X[0m"
				}
				Else {
					write-host "$X[38;2;255;000;000;22m$($Character.ShowOnDisplayScreen)$X[0m"
				}
				$DoCount = 0
			}
		}
		if ($Counter % 2 -eq 1) {#if $counter finished on an odd number
			write-host "        |"
		}
		write-host
		if ($Script:CharactersCSV.count -gt 10){ #will only work for first 99 accounts. CBF adding logic for over 100 single player accounts lol.
			Write-Host "   Type the number of the account you want to toggle or '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline #user will need to press enter
			$ToggleCharacter = (ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27 -TwoDigitAcctSelection $True).tostring()
		}
		Else {
			Write-Host "   Press the number of the account you want to toggle: " -nonewline
			$ToggleCharacter = (ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27).tostring()
		}
		$CharacterName = ($Script:CharactersCSV | Where-Object {$_.TemporaryID -eq $ToggleCharacter}).CharacterName
		foreach ($Character in $Script:CharactersCSV) {
			$Character.PSObject.Properties.Remove("TemporaryID")
		}
		$Script:CharactersCSV = import-csv "$Script:WorkingDirectory\characters.csv"
		if ($CharacterName){
			$CharacterToToggle = $Script:CharactersCSV | Where-Object {$_.CharacterName -eq $CharacterName}
			if ($CharacterToToggle.ShowOnDisplayScreen -eq "Hide"){
				$CharacterToToggle.ShowOnDisplayScreen = "Show"
				write-host "    $CharacterName set to $X[38;2;000;255;000;22mshow$X[0m on display screen.`n"	
			}
			Else {
				$CharacterToToggle.ShowOnDisplayScreen = "Hide"
				write-host "    $CharacterName set to $X[38;2;255;000;000;22mhide$X[0m on display screen.`n"
			}
			start-sleep -milliseconds 2400
		}
		if ($Null -ne $CharacterToToggle){
			$Script:CharactersCSV | Export-Csv -Path "$Script:WorkingDirectory\characters.csv" -NoTypeInformation
		}

	}
	ElseIf ($Option -eq "6"){ #ManualSettingSwitcherEnabled
		If ($Script:Config.ManualSettingSwitcherEnabled -eq "False"){
			$Options = @{"1" = "True"}
			$OptionsSubText = "enable"
			$CurrentState = "Disabled"
		}
		Else {
			$Options = @{"1" = "False"}
			$OptionsSubText = "disable"
			$Script:AskForSettings = $False
			$CurrentState = "Enabled"
		}
		$XMLChanged = OptionSubMenu -ConfigName "ManualSettingSwitcherEnabled" -OptionsList $Options -Current $CurrentState `
		-Description "This enables you to manually choose which settings file the game should use launching another game instance.`nFor example if you want to choose to launch with potato graphics or good graphics.`nPlease see GitHub for instructions on how to set this up and how to edit settings." `
		-OptionsText "Choose '$X[38;2;255;165;000;22m1$X[0m' to $OptionsSubText`n"
	}
	ElseIf ($Option -eq "7"){ #DisableVideos
		If ($Script:Config.DisableVideos -eq "False"){
			$Options = @{"1" = "True"}
			$OptionsSubText = "disable videos" #less confusing than simply saying "enable" disablevideos
			$CurrentState = "Videos Enabled"
		}
		Else {
			$Options = @{"1" = "False"}
			$OptionsSubText = "enable videos" #less confusing than simply saying "disable" disablevideos
			$CurrentState = "Videos Disabled"
		}
		$XMLChanged = OptionSubMenu -ConfigName "DisableVideos" -OptionsList $Options -Current $CurrentState `
		-Description "This enables you to disable intro videos and videos in between each act." `
		-OptionsText "Choose '$X[38;2;255;165;000;22m1$X[0m' to $OptionsSubText`n"
	}
	ElseIf ($Option -eq "8"){#Backup
		If ($Script:Config.AutoBackup -eq "False"){
			$Options = @{"1" = "True";"2" = "PlaceholderValue Only :)";"3" = "PlaceholderValue Only :)"}
			$OptionsSubText = "enable"
			$CurrentState = "Disabled"
		}
		Else {
			$Options = @{"1" = "False";"2" = "PlaceholderValue Only :)";"3" = "PlaceholderValue Only :)"}
			$OptionsSubText = "disable"
			$CurrentState = "Enabled"
		}
		$XMLChanged = OptionSubMenu -ConfigName "AutoBackup" -OptionsList $Options -Current $CurrentState `
		-Description "This enables you to disable intro videos and videos in between each act." `
		-OptionsText "Choose '$X[38;2;255;165;000;22m1$X[0m' to $OptionsSubText`nChoose '$X[38;2;255;165;000;22m2$X[0m' to make a manual backup`nChoose '$X[38;2;255;165;000;22m3$X[0m' to make Cloud Backup Setup`n"		
	}
	ElseIf ($Option -eq "9" -and $null -ne $D2rDirectories){ #Swap Character Packs. Specifically swap .d2s files, other files can be used by online chars.
		$CurrentD2rSaveFiles = Get-ChildItem -Path "$CharacterSavePath" -Filter "*.d2s" -File
		$OptionsList = @{}
		For ($iterate = 0; $iterate -lt $NonEmptyDirectories.Count; $iterate ++) {
			$Optionkey = ($iterate + 1).ToString()  # Create keys as "1", "2", etc.
			$OptionsList[$Optionkey] = $NonEmptyDirectories[$iterate].Name
		}
		Foreach ($Option in $OptionsList.GetEnumerator() | Sort-Object Key){
			write-host "   Choose '$X[38;2;255;165;000;22m$($Option.key)$X[0m' to switch to '$($Option.Value)' character set."
		}
		do {
			Write-Host "`n   Enter " -nonewline;CommaSeparatedList -NoOr ($OptionsList.keys | sort-object); Write-Host " or '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
			if ($OptionsList.count -gt 10){
				$NewOptionValue = (ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27 -TwoDigitAcctSelection $True).tostring()
			}
			Else {
				$NewOptionValue = (ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27).tostring()
			}
			if ($NewOptionValue -eq "c"){
				return
			}
			if (!$OptionsList.ContainsKey($NewOptionValue)){
				write-host " Please enter a valid option." -foregroundcolor red
			}
		} until ($NewOptionValue -eq "c" -or $OptionsList.ContainsKey($NewOptionValue))
		do {
			if ($Null -ne (Get-Process | Where-Object {$_.processname -eq "D2r"})){
				write-host " D2r is currently open. Please close the game to swap profiles`n" -foregroundcolor red
				PressTheAnyKey
				write-host
			}
		} until ($Null -eq (Get-Process | Where-Object {$_.processname -eq "D2r"}))
		foreach ($file in $CurrentD2rSaveFiles){
			Move-Item -Path $file.FullName -Destination "$CharacterSavePath\$CurrentProfile" | out-null
			Write-Debug "Moved $($file.Name) to '$CharacterSavePath\$CurrentProfile'"
		}
		$ProfileToSwapD2rSaveFiles = Get-ChildItem -Path ("$CharacterSavePath\" + $OptionsList["$NewOptionValue"]) -Filter "*.d2s" -File
		foreach ($file in $ProfileToSwapD2rSaveFiles){
			Move-Item -Path $file.FullName -Destination "$CharacterSavePath" | out-null
			Write-Debug "Moved $($file.Name) to '$CharacterSavePath'"
		}
		Write-Host " Swapped character set to $($OptionsList["$NewOptionValue"])`n" -foregroundcolor green
		start-sleep -milliseconds 333
	}
	else {#go to main menu if no valid option was specified.
		return
	}
	if ($XMLChanged -eq $True){
		Write-Host "   Config Updated!" -foregroundcolor green
		ImportXML
		If ($Option -eq "7"){
			DisableVideos
		}
		start-sleep -milliseconds 2500
	}
}
Function Notifications {
	param (
		[bool] $Check
	)
if ($Check -eq $True -and $Script:LastNotificationCheck -lt (Get-Date).addminutes(-30).ToString('yyyy.MM.dd HH:mm:ss')){#check for notifications once every 30mins
		try {
			$URI = "https://raw.githubusercontent.com/shupershuff/D2rSPLoader/main/Notifications.txt"
			$Script:Notifications = Invoke-RestMethod -Uri $URI
			if ($Notifications.notification -ne ""){
				if ($Script:PrevNotification -ne $Notifications.notification){#if message has changed since last check
					$Script:PrevNotification = $Notifications.notification
					$Script:NotificationHasBeenChecked = $False
					if ((get-date).tostring('yyyy.MM.dd HH:mm:ss') -lt $Notifications.ExpiryDate -and (get-date).tostring('yyyy.MM.dd HH:mm:ss') -gt $Notifications.PublishDate){
						$Script:NotificationsAvailable = $True
					}
				}
			}
			Else {
				$Script:NotificationsAvailable = $False
			}
			$Script:LastNotificationCheck = (get-date).tostring('yyyy.MM.dd HH:mm:ss')
		}
		Catch {
			Write-Debug "  Couldn't check for notifications." # If this fails in production don't show any errors/warnings.
		}
	}
	ElseIf ($Check -eq $False){
		Write-Host
		formatfunction -text $Notifications.notification -indents 1
	}
	if ($Check -eq $True -and $Script:NotificationHasBeenChecked -eq $False -and $Script:NotificationsAvailable -eq $True){#only show message if user hasn't seen notification yet.
		Write-Host "     $X[38;2;255;165;000;48;2;1;1;1;4mNotification available. Press 'i' to go to info screen for details.$X[0m"
	}#%%%%%%%%%%%%%%%%%%%%
}
Function QuoteList {
$Script:QuoteList =
"Stay a while and listen..",
"My brothers will not have died in vain!",
"My brothers have escaped you...",
"Not even death can save you from me.",
"Good Day!",
"You have quite a treasure there in that Horadric Cube.",
"There's nothing the right potion can't cure.",
"Well, what the hell do you want? Oh, it's you. Uh, hi there.",
"Your souls shall fuel the Hellforge!",
"What do you need?",
"Your presence honors me.",
"I'll put that to good use.",
"Good to see you!",
"Looking for Baal?",
"All who oppose me, beware",
"Greetings",
"We live...AGAIN!",
"Ner. Ner! Nur. Roah. Hork, Hork.",
"Greetings, stranger. I'm not surprised to see your kind here.",
"There is a place of great evil in the wilderness.",
"East... Always into the east...",
"I shall make weapons from your bones",
"I am overburdened",
"This magic ring does me no good.",
"The siege has everything in short supply...except fools.",
"Beware, foul demons and beasts.",
"They'll never see me coming.",
"I will cleanse this wilderness.",
"I shall purge this land of the shadow.",
"I hear foul creatures about.",
"Ahh yes, ruins, the fate of all cities.",
"I'm never gunna give you up, never gunna let you down - Griswold, 1996.",
"I have no grief for him. Oblivion is his reward.",
"The catapults have been silenced.",
"The staff of kings, you astound me!",
"What's the matter, hero? Questioning your fortitude? I know we are.",
"This whole place is one big ale fog.",
"So, this is daylight... It's over-rated.",
"When - or if - I get to Lut Gholein, I'm going to find the largest bowl`nof Narlant weed and smoke 'til all earthly sense has left my body.",
"I've just about had my fill of the walking dead.",
"Oh I hate staining my hands with the blood of foul Sorcerers!",
"Damn it, I wish you people would just leave me alone!",
"Beware! Beyond lies mortal danger for the likes of you!",
"Beware! The evil is strong ahead.",
"Only the darkest Magics can turn the sun black.",
"You are too late! HAA HAA HAA",
"You now speak to Ormus. He was once a great mage, but now lives like a`nrat in a sinking vessel",
"I knew there was great potential in you, my friend. You've done a`nfantastic job.",
"Hi there. I'm Charsi, the Blacksmith here in camp. It's good to see some`nstrong adventurers around here.",
"Whatcha need?",
"Good day to you partner!",
"Moomoo, moo, moo. Moo, Moo Moo Moo Mooo.",
"Moo.",
"Moooooooooooooo",
"So cold and damp under the earth...",
"Good riddance, Blood Raven.",
"I shall meet death head-on.",
"The land here is dead and lifeless.",
"Let the gate be opened!",
"It is good to know that the sun shines again.",
"Maybe now the world will have peace.",
"Eternal suffering would be too brief for you, Diablo!",
"All that's left of proud Tristram are ghosts and ashes.",
"Ahh, the slow torture of caged starvation.",
"What a waste of undead flesh.",
"Good journey, Mephisto. Give my regards to the abyss.",
"My my, what a messy little demon!",
"Ah, the familiar scent of death.",
"What evil taints the light of the sun?",
"Light guide my way in this accursed place.",
"I shall honor Tal Rasha's sacrifice by destroying all the Prime Evils.",
"Oops...did I do that?",
"Death becomes you Andariel.",
"You dark mages are all alike, obsessed with power.",
"Planting the dead. How odd.",
"'Live, Laugh, Love' - Andariel, 1264.",
"Oh no, snakes. I hate snakes.",
"Who would have thought that such primitive beings could cause so much `ntrouble.",
"Hail to you champion",
"Help us! LET US OUT!",
"'I cannot carry anymore' - Me, carrying my teammates.",
"You're an even greater warrior than I expected...Sorry for `nunderestimating you.",
"Cut them down, warrior. All of them!",
"How can one kill what is already dead?",
"...That which does not kill you makes you stronger."
}
Function BannerLogo {
	$BannerLogo = @"

  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%#%%%%%%%%%%%%%%%%%%%%
  %%%%%%%#%%%%%%%/%%%%%%%%%%%#%%%%%$Script:MOO%%%%##%%%%%%%%%#/##%%%%%%%%%%%%%%%%%%
  %%#(%/(%%%//(/*#%%%%%%%%%%%###%%%%#%%%%%###%%%%%###(*####%##%%%#%%*%%%%%%
  %%%( **/*///%%%%%%%%%###%%#(######%%#########%####/*/*.,#((#%%%#*,/%%%%%%
  %%%#*/.,*/,,*/#/%%%/(*#%%#*(%%(*/%#####%###%%%#%%%(///*,**/*/*(.&,%%%%%%%
  %%%%%%// % ,***(./*/(///,/(*,./*,*####%#*/#####/,/(((/.*.,.. (@.(%%%%%%%%
  %%%%%%%#%* &%%#..,,,,**,.,,.,,**///.*(#(*.,.,,*,,,,.,*, .&%&&.*%%%%%%%%%%
  %%%%%%%%%%%#.@&&%&&&%%%%&&%%,.(((//,,/*,,*.%&%&&&&&&&&%%&@%,#%%%%%%%%%%%%
  %%%%%%%%%%%%%(.&&&&&%&&&%(.,,*,,,,,.,,,,.,.*,%&&&%&&%&&&@*##%%%%%%%%%%%%%
  %%%%%%%%%%%%%%# @@@&&&&&(  @@&&@&&&&&&&&&*..,./(&&&&&&&&*####%%%%%%%%%%%%
  %%%%%%%%%%%%%%# &@@&&&&&(*, @@@&.,,,,. %@@&&*.,(%&&&&&&&/%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%#.&@@&&&&&(*, @@@@,((#&&%#.&@@&&.*#&&@&&&&/#%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%%*&@@@&&&&#*, @@@@,*(#%&&%#,@@@&@,(%&&&&&&(%%%%%%%%%%%%%%%%
  %%%%%$Script:MOO%%%%%%%*&@&&&%&&(,. @@@@,(%%%%%%#/,@@@& *#&&@&&%(%%%%%%%$Script:MOO%%%%%%
  %%%%%%%%%%%%%%%*&&&@&%&&(,. @@@@,%&%%%%%%(.@@@@ /#&&&&&&(%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%%,&&&&&%%&(*, @@@@,&&&&&&&%//@@@@./%&&&&@&(%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%(*&&&&&&&%(,, @@@@,%&&&#(/*.@@@@&./%&&&&@&(%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%%,&&&&&&&%(,, @@@@,/##/(// @@&@@,/#&&&&&&&(%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%(,&&&&&&&%(,, @@@@.*,,..*@@@&&*./#&&%&&&&&(%#%%#%%%%%%%%%%%
  %%%%%%%%%%%#%%#.&&&&&%%#* @@@&&&@@@&&%&&&% */*%&&%#&&&&&/((#%%%%%%%%%%%%%
  %%%%%%%%(#//*/.&&&#%#%#.@&& ..,,****,,*//((/*#%%%####%%%#/#/#%%%%%%%%%%%%
  %%%%%##***.,**////*(//,&.*/***.*/%%#%/%#*.***/*/***//**/(((/.,*(//*/(##%%
  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
"@
	if ($Script:PGemActivated -eq $True -or $Script:CowKingActivated -eq $True ){
		Write-Host "  $X[38;2;255;165;0;22m$BannerLogo$X[0m"
	}
	Else {
		Write-Host $BannerLogo -foregroundcolor yellow
	}
}
Function KillHandle { #Thanks to sir-wilhelm for tidying this up.
	$handle64 = "$PSScriptRoot\handle\handle64.exe"
	$handle = & $handle64 -accepteula -a -p D2R.exe "Check For Other Instances" -nobanner | Out-String
	if ($handle -match "pid:\s+(?<d2pid>\d+)\s+type:\s+Event\s+(?<eventHandle>\w+):"){
		$d2pid = $matches["d2pid"]
		$eventHandle = $matches["eventHandle"]
		Write-Verbose "Closing handle: $eventHandle on pid: $d2pid"
		& $handle64 -c $eventHandle -p $d2pid -y #-nobanner
	}
}
Function CheckActiveCharacters {#Note: only works for accounts loaded by the script
	#check if there's any open instances and check the game title window for which account is being used.
	try {
		$ActiveSinglePlayerInstance = $Null
		$D2rRunning = $false
		$ActiveSinglePlayerInstance = New-Object -TypeName System.Collections.ArrayList
		$ActiveSinglePlayerInstance = (Get-Process | Where-Object {$_.processname -eq "D2r" -and $_.MainWindowTitle -match "Diablo II: Resurrected \(SP\)"} | Select-Object MainWindowTitle).mainwindowtitle.trim()
		$Script:D2rRunning = $true
		Write-Verbose "Running Instances."
	}
	catch {#if the above fails then there are no running D2r instances.
		$Script:D2rRunning = $false
		Write-Verbose "No Running Instances."
		$ActiveSinglePlayerInstance = ""
	}
	if ($Script:D2rRunning -eq $True){ 
		$CurrentTime = get-date
		#Game updates save file every 5 min even if idle, see if most recent save file has been modified in last 5min.
		$RecentFile = Get-ChildItem -Path "$CharacterSavePath" -File -Filter "*.d2s" | Where-Object {($CurrentTime - $_.LastWriteTime).TotalMinutes -le 5} | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1
		
		if ($RecentFile) {
		Write-Debug " Most recently updated file: $($RecentFile.Name)"
			$Script:ActiveCharacter = $RecentFile.Name -replace "\.d2s$"
		}
		else {
			Write-Debug "No .d2s files have been updated in the last 5 minutes."
		}
	}
	else {
		$Script:ActiveCharacter = $Null
	}
}
Function DisplayCharacters {
	Write-Host
	$LongestCharNameLength = ($Script:CharactersCSV | where-object {$_.ShowOnDisplayScreen -ne "Hide"}).CharacterName | ForEach-Object { $_.Length } | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum #find out how many batches there are so table can be properly indented.
	if ($LongestCharNameLength -ge 28){$LongestCharNameLength = 28} #we will limit long display names on the main screen to prevent odd things happening.
	while ($CharHeaderIndent.length -lt ($LongestCharNameLength -14)){#indent the header for 'Hours Played' based on how long the longest character name is
		$CharHeaderIndent = $CharHeaderIndent + " "
	}
	Write-Host ("    Character Name   " + $CharHeaderIndent + "Hours Played   ") #Header
	ForEach ($Character in ($Script:CharactersCSV | where-object {$_.ShowOnDisplayScreen -ne "Hide"}|  Sort-Object -Property CharacterName)){
		if ($Character.CharacterName.length -gt 28){ #later we ensure that strings longer than 28 chars are cut short so they don't disrupt the display.
			$Character.CharacterName = $Character.CharacterName.Substring(0, 28)
		}
		try {
			$CharPlayTime = (" " + (($time =([TimeSpan]::Parse($Character.TimeActive))).hours + ($time.days * 24)).tostring() + ":" + ("{0:D2}" -f $time.minutes) + "   ")  # Add hours + (days*24) to show total hours, then show ":" followed by minutes
		}
		catch {#if character hasn't been played yet (with script running).
			$CharPlayTime = "   0   "
			Write-Debug "Character not played yet."
		}
		if ($CharPlayTime.length -lt 14){#formatting. Depending on the amount of characters for this variable push it out until it's 13 chars long.
			while ($CharPlayTime.length -lt 14){
				$CharPlayTime = " " + $CharPlayTime
			}
		}
		$CharIndent = ""
		if ($LongestCharNameLength -lt 14){#If longest character labels are shorter than 14 characters (header length), set this variable to the minimum (14) so indenting works properly.
			$LongestCharNameLength = 14
		}
		while (($Character.CharacterName.length + $CharIndent.length) -le $LongestCharNameLength){#keep adding indents until character name plus the indents matches the longest character name. Keeps table nice and neat.
			$CharIndent = $CharIndent + " "
		}
		if ($Character.CharacterName -eq $Script:ActiveCharacter){ #if character is currently active
			Write-Host ("    " + $Character.CharacterName + $CharIndent + "   " + $CharPlayTime + "(Character Active)") -foregroundcolor yellow
		}
		else {#if character isn't currently active
			Write-Host ("    " + $Character.CharacterName + $CharIndent + "   " + $CharPlayTime) -foregroundcolor green
		}
	}
}
Function GetCharacters {
	CheckForModSavePath -CheckOnly
	$D2rCharacters = Get-ChildItem -Path "$CharacterSavePath" -Filter "*.d2s" -File
	$Script:CharactersCSV = @(Import-Csv -Path "$Script:WorkingDirectory\characters.csv")
	ForEach ($D2rCharacter in $D2rCharacters){
		if (($D2rCharacter.name -replace '\.d2s$') -notin $CharactersCSV.charactername){
			$Char = $D2rCharacter.name -replace '\.d2s$'
			write-debug " Character $($D2rCharacter.name -replace '\.d2s$') is missing, will add to characters.csv."
			$NewCharacter = [PSCustomObject]@{
				CharacterName		= "$($D2rCharacter.name -replace '\.d2s$')"
				ShowOnDisplayScreen	= "Show"
				TimeActive			= ""
			}
			$Script:CharactersCSV += $NewCharacter
		}
	}
	if ($Null -ne $NewCharacter){
		$Script:CharactersCSV | Export-Csv -Path "$Script:WorkingDirectory\characters.csv" -NoTypeInformation
	}
}
Function Menu {
	Clear-Host
	if ($Script:ScriptHasBeenRun -ne $true){
		Write-Host ("  You have quite a treasure there in that Horadric SP Launcher v" + $Currentversion)
	}
	Notifications -check $True
	BannerLogo
	QuoteRoll
	ChooseAccount
	Processing
	Menu
}
Function ChooseAccount {
	do {
		if ($Script:MainMenuOption -eq "i"){
			Inventory #show stats
			$Script:MainMenuOption = "r"
		}
		if ($Script:MainMenuOption -eq "o"){ #options menu
			Options
			$Script:MainMenuOption = "r"
		}
		if ($Script:MainMenuOption -eq "s"){
			if ($Script:AskForSettings -eq $True){
				Write-Host "  Manual Setting Switcher Disabled." -foregroundcolor Green
				$Script:AskForSettings = $False
			}
			else {
				Write-Host "  Manual Setting Switcher Enabled." -foregroundcolor Green
				$Script:AskForSettings = $True
			}
			Start-Sleep -milliseconds 1550
			$Script:MainMenuOption = "r"
		}
		if ($Script:MainMenuOption -eq "g"){#silly thing to replicate in game chat gem.
			$Script:CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
			if ($Script:GemActivated -ne $True){
				$GibberingGemstone = get-random -minimum 0 -maximum  4095
				if ($GibberingGemstone -eq 69 -or $GibberingGemstone -eq 420){#nice
					Write-Host "  Perfect Gem Activated" -ForegroundColor magenta
					Write-Host "`n     OMG!" -foregroundcolor green
					$Script:PGemActivated = $True
					([int]$Script:CurrentStats.PerfectGems) ++
					SetQualityRolls
					Start-Sleep -milliseconds 4567
				}
				else {
					if ($GibberingGemstone -in 16..32){
						CowKingKilled
						$SkipCSVExport = $True
					}
					else {
						Write-Host "  Gem Activated" -ForegroundColor magenta
						([int]$Script:CurrentStats.Gems) ++
					}
				}
				$Script:GemActivated = $True
				SetQualityRolls
				if ($SkipCSVExport -ne $True){
					try {
						$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv with Total Time played.
					}
					Catch {
						Write-Host "  Couldn't update stats.csv" -foregroundcolor yellow
					}
				}
			}
			Else {
				Write-Host "  Gem Deactivated" -ForegroundColor magenta
				$Script:GemActivated = $False
				SetQualityRolls
			}
			Start-Sleep -milliseconds 850
			$Script:MainMenuOption = "r"
		}
		if ($Script:MainMenuOption -eq "r"){#refresh
			Clear-Host
			Notifications -check $True
			BannerLogo
			QuoteRoll
		}
		if ($Script:Config.AutoBackup -eq $True){
			$CurrentTime = Get-Date
			if ($CurrentTime.Minute -eq 0 -or $CurrentTime.Minute -eq 30) {
				if ($currenttime.addminutes(-1) -ge $lastbackup) { #if a backup hasn't just happened. (prevents two backups happening in a minute).
					$null = LocalBackup #Run a backup if the time is exactly on a 30 minute mark.
					$Script:LastBackup = $CurrentTime
				}
			}
		}
		GetCharacters
		CheckActiveCharacters
		DisplayCharacters
		$OpenD2rSPLoaderInstances = Get-WmiObject -Class Win32_Process | Where-Object { $_.name -eq "powershell.exe" -and $_.commandline -match $Script:ScriptFileName} | Select-Object name,processid,creationdate | Sort-Object creationdate -descending
		if ($OpenD2rSPLoaderInstances.length -gt 1){#If there's more than 1 D2rSPloader.ps1 script open, close until there's only 1 open to prevent the time played accumulating too quickly.
			ForEach ($Process in $OpenD2rSPLoaderInstances[1..($OpenD2rSPLoaderInstances.count -1)]){
				Stop-Process -id $Process.processid -force #Closes oldest running D2rSPLoader script
			}
		}
		if ($Script:ActiveCharacter){#if there is an active character, add to total script time
			#Add time against the character that's being played
			$Script:CharactersCSV = import-csv "$Script:WorkingDirectory\characters.csv"
			$AdditionalTimeSpan = New-TimeSpan -Start $Script:StartTime -End (Get-Date) #work out elapsed time to add to characters.csv
			$CharacterToUpdate = $Script:CharactersCSV | Where-Object {$_.CharacterName -eq $Script:ActiveCharacter}
			if ($CharacterToUpdate){
				try {#get current time from csv and add to it
					$CharacterToUpdate.TimeActive = [TimeSpan]::Parse($CharacterToUpdate.TimeActive) + $AdditionalTimeSpan
				}
				Catch {#if CSV hasn't been populated with a time yet.
					$CharacterToUpdate.TimeActive = $AdditionalTimeSpan
				}
			}
			try {
				$Script:CharactersCSV | Export-Csv -Path "$Script:WorkingDirectory\characters.csv" -NoTypeInformation #update characters.csv with the new time played.
			}
			Catch {
				$WriteAcctCSVError = $True
			}
			$Script:SessionTimer = $Script:SessionTimer + $AdditionalTimeSpan #track current session time but only if a game is running
			if ($WriteAcctCSVError -eq $true){
				Write-Host "`n  Couldn't update characters.csv with playtime info." -ForegroundColor Red
				Write-Host "  It's likely locked for editing, please ensure you close this file." -ForegroundColor Red
				start-sleep -milliseconds 1500
				$WriteAcctCSVError = $False
			}
			#Add Time to Total Script Time only if there's an open game.
			$Script:CurrentStats = import-csv "$Script:WorkingDirectory\Stats.csv"
			try {
				$AdditionalTimeSpan = New-TimeSpan -Start $Script:StartTime -End (Get-Date)
				try {#get current time from csv and add to it
					$Script:CurrentStats.TotalGameTime = [TimeSpan]::Parse($CurrentStats.TotalGameTime) + $AdditionalTimeSpan
				}
				Catch {#if CSV hasn't been populated with a time yet.
					$Script:CurrentStats.TotalGameTime = $AdditionalTimeSpan
				}
				$CurrentStats | Export-Csv -Path "$Script:WorkingDirectory\Stats.csv" -NoTypeInformation #update Stats.csv with Total Time played.
			}
			Catch {
				Write-Host "`n  Couldn't update Stats.csv with playtime info." -ForegroundColor Red
				Write-Host "  It's likely locked for editing, please ensure you close this file." -ForegroundColor Red
				start-sleep -milliseconds 1500
			}
		}
		$Script:StartTime = Get-Date #restart timer for session time and character time.
		do {
			Write-Host
			if ($Script:D2rRunning -ne $True){
				Write-Host "  Press '$X[38;2;255;165;000;22mEnter$X[0m' or '$X[38;2;255;165;000;22mp$X[0m' to launch the game (in offline mode)."
				Write-Host "  Alternatively choose from the following menu options:"
			}
			else {
				Write-Host "  Choose from the following menu options:"
			}			
			if ($Script:Config.ManualSettingSwitcherEnabled -eq $true){
				$ManualSettingSwitcherOption = "s"
				Write-Host "  '$X[38;2;255;165;000;22mr$X[0m' to Refresh, $X[38;2;255;165;000;22mo$X[0m' for config options, '$X[38;2;255;165;000;22mi$X[0m' for info"
				Write-Host "  '$X[38;2;255;165;000;22ms$X[0m' to toggle the Manual Setting Switcher, or '$X[38;2;255;165;000;22mx$X[0m' to $X[38;2;255;000;000;22mExit$X[0m: "-nonewline
			}
			Else {
				$ManualSettingSwitcherOption = $null
				Write-Host "  '$X[38;2;255;165;000;22mr$X[0m' to Refresh, $X[38;2;255;165;000;22mo$X[0m' for config options, '$X[38;2;255;165;000;22mi$X[0m' for info or '$X[38;2;255;165;000;22mx$X[0m' to $X[38;2;255;000;000;22mExit$X[0m: " -nonewline
			}
			$Script:MainMenuOption = ReadKeyTimeout "" $MenuRefreshRate "r" -AdditionalAllowedKeys 13 #$MenuRefreshRate represents the refresh rate of the menu in seconds (30). if no button is pressed, send "r" for refresh.
			if ($script:key.virtualkeycode -eq "13"){
				if ($Null -eq (Get-Process | Where-Object {$_.processname -eq "D2r" -and $_.MainWindowTitle -match "Diablo II: Resurrected \(SP\)"})){
					$Script:MainMenuOption = "p"
				}
				else {
					$Script:MainMenuOption = "r"
				}
			}
			if ($Script:MainMenuOption -notin ("x" + "r" + "g" + "i" + "o" + "p" + $ManualSettingSwitcherOption).ToCharArray()){
				Write-Host " Invalid Input. Please enter one of the options above." -foregroundcolor red
				$Script:MainMenuOption = $Null
			}
		} until ($Null -ne $Script:MainMenuOption)
		if ($Null -ne $Script:MainMenuOption){
			if ($Script:MainMenuOption -eq "x"){
				Write-Host "`n Good day to you partner :)" -foregroundcolor yellow
				Start-Sleep -milliseconds 486
				Exit
			}
		}
		$Script:RunOnce = $True
	} until ($Script:MainMenuOption -ne "r" -and $Script:MainMenuOption -ne "g" -and $Script:MainMenuOption -ne "s" -and $Script:MainMenuOption -ne "i" -and $Script:MainMenuOption -ne "o")
}
Function Processing {
	#Open diablo with parameters
		# IE, this is essentially just opening D2r like you would with a shortcut target of "C:\Program Files (x86)\Battle.net\Games\Diablo II Resurrected\D2R.exe" -username <yourusername -password <yourPW> -address <SERVERaddress>
	$CustomLaunchArguments = ($Config.CustomLaunchArguments).replace("`"","").replace("'","") #clean up arguments in case they contain quotes (for folks that have used excel to edit characters.csv).
	$arguments = (" -address xxx" + " " + $CustomLaunchArguments).tostring() #force offline mode by giving it a garbage region to connect to.
	if ($Config.ForceWindowedMode -eq $true){#starting with forced window mode sucks, but someone asked for it.
		$arguments = $arguments + " -w"
	}
	#Switch Settings file to load D2r from.
	if ($Script:AskForSettings -eq $True){#steps go through if user has toggled on the manual setting switcher ('s' in the menu).
		CheckForModSavePath -settings
		$SettingsJSON = ($SettingsProfilePath + "Settings.json")
		$files = Get-ChildItem -Path $SettingsProfilePath -Filter "settings.*.json"
		$Counter = 1
		$SettingsDefaultOptionArray = New-Object -TypeName System.Collections.ArrayList #Add in an option for the default settings file (if it exists, if the auto switcher has never been used it won't appear.
		$SettingsDefaultOption = New-Object -TypeName psobject
		$SettingsDefaultOption | Add-Member -MemberType NoteProperty -Name "ID" -Value $Counter
		$SettingsDefaultOption | Add-Member -MemberType NoteProperty -Name "Name" -Value ("Default - settings.json")
		$SettingsDefaultOption | Add-Member -MemberType NoteProperty -Name "FileName" -Value ("settings.json")
		[VOID]$SettingsDefaultOptionArray.Add($SettingsDefaultOption)
		$SettingsFileOptions = New-Object -TypeName System.Collections.ArrayList
		ForEach ($file in $files){
			 $SettingsFileOption = New-Object -TypeName psobject
			 $Counter = $Counter + 1
			 $Name = $file.Name -replace '^settings\.|\.json$' #remove 'settings.' and '.json'. The text in between the two periods is the name.
			 $SettingsFileOption | Add-Member -MemberType NoteProperty -Name "ID" -Value $Counter
			 $SettingsFileOption | Add-Member -MemberType NoteProperty -Name "Name" -Value $Name
			 $SettingsFileOption | Add-Member -MemberType NoteProperty -Name "FileName" -Value $file.Name
			 [VOID]$SettingsFileOptions.Add($SettingsFileOption)
		}
		if ($Null -ne $SettingsFileOptions){# If settings files are found, IE the end user has set them up prior to running script.
			$SettingsFileOptions = $SettingsDefaultOptionArray + $SettingsFileOptions
			Write-Host "  Settings options you can choose from are:"
			ForEach ($Option in $SettingsFileOptions){
				Write-Host ("   " + $Option.ID + ". " + $Option.name) -foregroundcolor green
			}
			do {
				Write-Host "  Choose the settings file you like to load from: " -nonewline
				ForEach ($Value in $SettingsFileOptions.ID){ #write out each account option, comma separated but show each option in orange writing. Essentially output overly complicated fancy display options :)
					if ($Value -ne $SettingsFileOptions.ID[-1]){
						Write-Host "$X[38;2;255;165;000;22m$Value$X[0m" -nonewline
						if ($Value -ne $SettingsFileOptions.ID[-2]){Write-Host ", " -nonewline}
					}
					else {
						Write-Host " or $X[38;2;255;165;000;22m$Value$X[0m"
					}
				}
				if ($Null -eq $ManualSettingSwitcher){#if not launched from parameters
					Write-Host "  Or Press '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
					$SettingsCancelOption = "c"
				}
				$SettingsChoice = ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27 #$MenuRefreshRate represents the refresh rate of the menu in seconds (30). If no button is pressed, send "c" for cancel.
				Write-Host
				if ($SettingsChoice.tostring() -notin $SettingsFileOptions.id + $SettingsCancelOption + "Esc"){
					Write-Host "  Invalid Input. Please enter one of the options above." -foregroundcolor red
					$SettingsChoice = ""
				}
			} until ($SettingsChoice.tostring() -in $SettingsFileOptions.id + $SettingsCancelOption + "Esc")
			if ($SettingsChoice -ne "c" -and $SettingsChoice -ne "Esc" -and $SettingsChoice -ne "1"){
				$SettingsChoice 
				pause
				$SettingsToLoadFrom = $SettingsFileOptions | where-object {$_.id -eq $SettingsChoice.tostring()}
				try {
					Copy-item ($SettingsProfilePath + $SettingsToLoadFrom.FileName) -Destination $SettingsJSON #-ErrorAction Stop #overwrite settings.json with settings<Name>.json (<Name> being the name of the config user selects). This means any changes to settings in settings.json will be lost the next time an account is loaded by the script.
					Write-Host (" Custom game settings (" + $SettingsToLoadFrom.Name + ") being used.`n") -foregroundcolor green
					Start-Sleep -milliseconds 100
				}
				catch {
					FormatFunction -Text "Couldn't overwrite settings.json for some reason. Make sure you don't have the file open!" -IsError
					PressTheAnyKey
				}
			}
		}
		Else {# if no custom settings files are found, IE user hasn't set them up yet.
			Write-Host "`n  No Custom Settings files have been saved yet. Loading default settings." -foregroundcolor Yellow
			Write-Host "  See README for setup instructions.`n" -foregroundcolor Yellow
			PressTheAnyKey
		}
	}
	if ($SettingsChoice -ne "c" -and $SettingsChoice -ne "Esc"){
		if ($Script:StartWithDisableVideosMod -eq $True){
			$arguments += " -mod DisableVideos -txt"
		}
		#Start Game
		KillHandle | out-null
		$process = Start-Process "$Gamepath\D2R.exe" -ArgumentList "$arguments --instanceSinglePlayer" -PassThru 
		Start-Sleep -milliseconds 1500 #give D2r a bit of a chance to start up before trying to kill handle
		#Close the 'Check for other instances' handle
		Write-Host " Attempting to close `"Check for other instances`" handle..."
		$Output = KillHandle | out-string #run KillHandle function.
		if (($Output.contains("DiabloII Check For Other Instances")) -eq $true){
			$handlekilled = $true
			Write-Host " `"Check for Other Instances`" Handle closed." -foregroundcolor green
		}
		else {
			Write-Host " `"Check for Other Instances`" Handle was NOT closed." -foregroundcolor red
			Write-Host " Who even knows what happened. I sure don't." -foregroundcolor red
			FormatFunction -text " If you are seeing this error and are running the script for the first time`n" -IsError
			PressTheAnyKey
		}
		if ($handlekilled -ne $True){
			Write-Host " Couldn't find any handles to kill." -foregroundcolor red
			Write-Host " Game may not have launched as expected." -foregroundcolor red
			PressTheAnyKey
		}
		#Rename the Diablo Game window for easier identification of which character the game is.
		$rename = ("Diablo II: Resurrected (SP)")
		$Command = ('"'+ $WorkingDirectory + '\SetText\SetTextv2.exe" /PID ' + $process.id + ' "' + $rename + '"')
		try {
			cmd.exe /c $Command
			write-debug $Command #debug
			Write-debug " Window Renamed." #debug
			Start-Sleep -milliseconds 250
		}
		catch {
			Write-Host " Couldn't rename window :(" -foregroundcolor red
			PressTheAnyKey
		}
		if ($Config.GrailAppExecutablePath -ne ""){
			$GrailProcessName = (split-path $Config.GrailAppExecutablePath -leaf).trim(".exe") #check if app is already running, if so we will skip
			if ($null -eq (Get-Process | Where-Object {$_.processname -eq $GrailProcessName})){	
				Start-Job -ScriptBlock {
					param($exePath)
					Start-Process -FilePath $exePath #-WindowStyle Hidden
				} -ArgumentList $Config.GrailAppExecutablePath | Out-Null
			}
		}
		if ($Config.RunTimerAppExecutablePath -ne ""){
			$TimerProcessName = (split-path $Config.RunTimerAppExecutablePath -leaf).trim(".exe") #check if app is already running, if so we will skip
			if ($null -eq (Get-Process | Where-Object {$_.processname -eq $TimerProcessName})){	
				Start-Job -ScriptBlock {
					param($exePath)
					Start-Process -FilePath $exePath #-WindowStyle Hidden
				} -ArgumentList $Config.RunTimerAppExecutablePath | Out-Null
			}
		}
		Write-Host "`nGood luck hero..." -foregroundcolor magenta
		Start-Sleep -milliseconds 1000
		$Script:ScriptHasBeenRun = $true
		Remove-Job *
	}
}
RemoveMaximiseButton
InitialiseCurrentStats
CheckForUpdates
ImportXML
ValidationAndSetup
DisableVideos
ImportCSV
Clear-Host
QuoteList
SetQualityRolls
Menu

#For Diablo II: Resurrected
#Dedicated to my cat Toby.

#Text Colors
#Color Hex	RGB Value	Description
#FFA500		255 165 000	Crafted items
#4169E1		065 105 225	Magic items
#FFFF00		255 255 000	Rare items
#00FF00		000 255 000	Set items
#A59263		165 146 099	Unique items
