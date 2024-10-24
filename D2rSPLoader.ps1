<#
Author: Shupershuff
Usage:
Happy for you to make any modifications to this script for your own needs providing:
- Any variants of this script are never sold.
- Any variants of this script published online should always be open source.
Purpose:
	Script is mainly orientated around tracking character playtime and total game time for single player.
	Script will import account details from CSV.
Instructions: See GitHub readme https://github.com/shupershuff/D2rSPLoader


1.0 to do list
remove auto setting switcher
Add skip intro and skip video options, and notify user if they're using customcommands
add save game checks and calculate active player by savegame with most recent update. need to confirm D2r saving pattern.
add options for seed, nosave. enable respec, playersX, -resetofflinemaps.
Couldn't write :) in release notes without it adding a new line, some minor issue with formatfunction regex
Fix whatever I broke or poorly implemented in the last update :)
#>

$CurrentVersion = "0.1" #single player edit, adjusted shortcut.
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
$Script:StartTime = Get-Date #Used for elapsed time. Is reset when script refreshes.
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
Function ReadKey([string]$message=$Null,[bool]$NoOutput,[bool]$AllowAllKeys){#used to receive user input
	$key = $Null
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
					$key = $key_
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
	$key = $Null
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
					$key = $key_
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
						Copy-Item -Path $Script:WorkingDirectory\Stats.backup.csv -Destination $Script:WorkingDirectory\Stats.csv
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
				if ($ShouldUpdate -eq "y" -or $ShouldUpdate -eq "yes"){#if user wants to update script, download .zip of latest release, extract to temporary folder and replace old D2Loader.ps1 with new D2Loader.ps1
					Write-Host "`n Updating... :)" -foregroundcolor green
					try {
						New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") -ErrorAction stop | Out-Null #create temporary folder to download zip to and extract
					}
					Catch {#if folder already exists for whatever reason.
						Remove-Item -Path ($Script:WorkingDirectory + "\UpdateTemp\") -Recurse -Force
						New-Item -ItemType Directory -Path ($Script:WorkingDirectory + "\UpdateTemp\") | Out-Null #create temporary folder to download zip to and extract
					}
					$ZipURL = $ReleaseInfo.zipball_url #get zip download URL
					$ZipPath = ($WorkingDirectory + "\UpdateTemp\D2Loader_" + $ReleaseInfo.tag_name + "_temp.zip")
					Invoke-WebRequest -Uri $ZipURL -OutFile $ZipPath
					if ($Null -ne $releaseinfo.assets.browser_download_url){#Check If I didn't forget to make a version.zip file and if so download it. This is purely so I can get an idea of how many people are using the script or how many people have updated. I have to do it this way as downloading the source zip file doesn't count as a download in github and won't be tracked.
						Invoke-WebRequest -Uri $releaseinfo.assets.browser_download_url -OutFile $null | out-null #identify the latest file only.
					}
					$ExtractPath = ($Script:WorkingDirectory + "\UpdateTemp\")
					Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
					$FolderPath = Get-ChildItem -Path $ExtractPath -Directory -Filter "shupershuff*" | Select-Object -ExpandProperty FullName
					Copy-Item -Path ($FolderPath + "\D2Loader.ps1") -Destination ($Script:WorkingDirectory + "\" + $Script:ScriptFileName) #using $Script:ScriptFileName allows the user to rename the file if they want
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
			$ZipPath = ($WorkingDirectory + "\UpdateTemp\D2Loader_" + $ReleaseInfo.tag_name + "_temp.zip")
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
		$Script:Config = ([xml](Get-Content "$Script:WorkingDirectory\Config.xml" -ErrorAction Stop)).D2loaderconfig
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
	#Perform some validation on config.xml. Helps avoid errors for people who may be on older versions of the script and are updating. Will look to remove all of this in a future update.
	$Script:Config = ([xml](Get-Content "$Script:WorkingDirectory\Config.xml" -ErrorAction Stop)).D2loaderconfig #import config.xml again for any updates made by the above.
	#check if there's any missing config.xml options, if so user has out of date config file.
	$AvailableConfigs = #add to this if adding features.
	"GamePath",
	"CustomLaunchArguments",
	"ShortcutCustomIconPath"
	$BooleanConfigs =
	"ManualSettingSwitcherEnabled",
	"RememberWindowLocations",
	"CreateDesktopShortcut",
	"ForceWindowedMode",
	"SettingSwitcherEnabled"
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
Function ImportCSV { #Import Account CSV
	do {
		try {
			$Script:AccountOptionsCSV = import-csv "$Script:WorkingDirectory\characters.csv" #import all accounts from csv
		}
		Catch {
			FormatFunction -text "`ncharacters.csv does not exist. Make sure you create this and populate with accounts first." -IsError
			PressTheAnyKeyToExit
		}
		if ($Null -ne $Script:AccountOptionsCSV){
			#check characters.csv has been updated and doesn't contain the example account.
			if ($Script:AccountOptionsCSV -match "yourcharactername"){
				Write-Host "`n You haven't setup characters.csv with your characters." -foregroundcolor red
				Write-Host " Add your character details to the CSV file and run the script again :)`n" -foregroundcolor red
				PressTheAnyKeyToExit
			}
			if ($Null -ne ($AccountOptionsCSV | Where-Object {$_.id -eq ""})){
				$Script:AccountOptionsCSV = $Script:AccountOptionsCSV | Where-Object {$_.id -ne ""} # To account for user error, remove any empty lines from characters.csv
			}
			ForEach ($Account in $AccountOptionsCSV){
				if ($Account.CharacterName -eq ""){ # if user doesn't specify a friendly name, use id. Prevents display issues later on.
					$Account.CharacterName = ("Account " + $Account.id)
					$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\characters.csv" -NoTypeInformation
				}
			}
			$DuplicateIDs = $AccountOptionsCSV | Where-Object {$_.id -ne ""} | Group-Object -Property ID | Where-Object { $_.Count -gt 1 }
			if ($duplicateIDs.Count -gt 0){
				$duplicateIDs = ($DuplicateIDs.name | out-string).replace("`r`n",", ").trim(", ") #outputs more meaningful error.
				Write-Host "`n characters.csv has duplicate IDs: $duplicateIDs" -foregroundcolor red
				FormatFunction -Text "Please adjust characters.csv so that the ID numbers against each account are unique.`n" -IsError
				PressTheAnyKeyToExit
			}
			if (-not ($Script:AccountOptionsCSV | Get-Member -Name "TimeActive" -MemberType NoteProperty -ErrorAction SilentlyContinue)){#For update 1.8.0. If TimeActive column doesn't exist, add it
				# Column does not exist, so add it to the CSV data
				$Script:AccountOptionsCSV | ForEach-Object {
					$_ | Add-Member -NotePropertyName "TimeActive" -NotePropertyValue $Null
				}
				# Export the updated CSV data back to the file
				$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\characters.csv" -NoTypeInformation
				Write-Host " Added TimeActive column to characters.csv." -foregroundcolor Green
				PressTheAnyKey
			}
			$AccountCSVImportSuccess = $True
		}
		else {#Error out and exit if there's a problem with the csv.
			if ($AccountCSVRecoveryAttempt -lt 1){
				try {
					Write-Host " Issue with characters.csv. Attempting Autorecovery from backup..." -foregroundcolor red
					Copy-Item -Path $Script:WorkingDirectory\characters.backup.csv -Destination $Script:WorkingDirectory\characters.csv
					Write-Host " Autorecovery successful!" -foregroundcolor Green
					$AccountCSVRecoveryAttempt ++
					PressTheAnyKey
				}
				Catch {
					$AccountCSVImportSuccess = $False
				}
			}
			Else {
				$AccountCSVRecoveryAttempt = 2
			}
			if ($AccountCSVImportSuccess -eq $False -or $AccountCSVRecoveryAttempt -eq 2){
				Write-Host "`n There's an issue with characters.csv." -foregroundcolor red
				Write-Host " Please ensure that this is filled out correctly and rerun the script." -foregroundcolor red
				Write-Host " Alternatively, rebuild CSV from scratch or restore from characters.backup.csv`n" -foregroundcolor red
				PressTheAnyKeyToExit
			}
		}
	} until ($AccountCSVImportSuccess -eq $True)
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
	Copy-Item -Path ($Script:WorkingDirectory + "\characters.csv") -Destination ($Script:WorkingDirectory + "\characters.backup.csv")
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
Function LoadWindowClass { #Used to get window locations and place them in the same screen locations at launch. Code courtesy of Sir-Wilhelm and Microsoft.
	try {
		[void][Window]
	}
	catch {
		Add-Type @"
		using System;
		using System.Runtime.InteropServices;
		public class Window {
			[DllImport("user32.dll")]
			[return: MarshalAs(UnmanagedType.Bool)]
			public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect); //get window coordinates
			[DllImport("user32.dll")]
			[return: MarshalAs(UnmanagedType.Bool)]
			public extern static bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw); //set window coordinates
			[DllImport("user32.dll")]
			[return: MarshalAs(UnmanagedType.Bool)]
			public static extern bool SetForegroundWindow(IntPtr hWnd); //Bring window to front
            [DllImport("user32.dll")]
			[return: MarshalAs(UnmanagedType.Bool)]
			public static extern bool ShowWindow(IntPtr handle, int state); //used in this script to restore minimized window (state 9)
			
			// Add SetWindowPos
            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
		}
		public struct RECT {
			public int Left;        // x position of upper-left corner
			public int Top;         // y position of upper-left corner
			public int Right;       // x position of lower-right corner
			public int Bottom;      // y position of lower-right corner
		}
"@	}
}
Function SaveWindowLocations {# Get Window Location coordinates and save to characters.csv
	LoadWindowClass
	FormatFunction -indents 2 -text "Saving locations of each open account so that they the windows launch in the same place next time. Assumes you've configured the game to launch in windowed mode."
	CheckActiveAccounts
	#If Feature is enabled, add 'WindowXCoordinates' and 'WindowYCoordinates' columns to characters.csv with empty values.
	if (-not ($Script:AccountOptionsCSV | Get-Member -Name "WindowXCoordinates" -MemberType NoteProperty -ErrorAction SilentlyContinue) -or -not ($Script:AccountOptionsCSV | Get-Member -Name "WindowYCoordinates" -MemberType NoteProperty -ErrorAction SilentlyContinue) -or -not ($Script:AccountOptionsCSV | Get-Member -Name "WindowWidth" -MemberType NoteProperty -ErrorAction SilentlyContinue) -or -not ($Script:AccountOptionsCSV | Get-Member -Name "WindowHeight" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
		# Column does not exist, so add it to the CSV data
		if (-not ($Script:AccountOptionsCSV | Get-Member -Name "WindowXCoordinates" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
			$Script:AccountOptionsCSV | ForEach-Object {$_ | Add-Member -NotePropertyName "WindowXCoordinates" -NotePropertyValue $Null}
		}
		if (-not ($Script:AccountOptionsCSV | Get-Member -Name "WindowYCoordinates" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
			$Script:AccountOptionsCSV | ForEach-Object {$_ | Add-Member -NotePropertyName "WindowYCoordinates" -NotePropertyValue $Null}
		}
		if (-not ($Script:AccountOptionsCSV | Get-Member -Name "WindowHeight" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
			$Script:AccountOptionsCSV | ForEach-Object {$_ | Add-Member -NotePropertyName "WindowHeight" -NotePropertyValue $Null}
		}
		if (-not ($Script:AccountOptionsCSV | Get-Member -Name "WindowWidth" -MemberType NoteProperty -ErrorAction SilentlyContinue)){
			$Script:AccountOptionsCSV | ForEach-Object {$_ | Add-Member -NotePropertyName "WindowWidth" -NotePropertyValue $Null}
		}
		# Export the updated CSV data back to the file
		$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\characters.csv" -NoTypeInformation
	}
	if ($null -eq $Script:ActiveAccountsList){
		FormatFunction -text "`nThere are no open accounts to save coordinates from.`nTo save Window positions, you need to launch one or more instances first.`n" -indents 2 -IsError
		PressTheAnyKey
		Return $False
	}
	$NewCSV = ForEach ($Account in $Script:AccountOptionsCSV){
		if ($account.id -in $Script:ActiveAccountsList.id){
			$process = Get-Process -Id ($Script:ActiveAccountsList | where-object {$_.id -eq $account.id}).ProcessID
			$handle = $process.MainWindowHandle
			Write-Verbose "$($process.ProcessName) `(Id=$($process.Id), Handle=$handle`, Path=$($process.Path))"
			$rectangle = New-Object RECT
			[Window]::GetWindowRect($handle, [ref]$rectangle) | Out-Null
			FormatFunction -indents 2 -text "`nSaved Coordinates for account $($account.id) ($($account.CharacterName))" -IsSuccess
			Write-Host "     X Position = $($rectangle.Left)" -Foregroundcolor Green
			Write-Host "     Y Position = $($rectangle.Top)" -Foregroundcolor Green
			write-Host "     Width = $($rectangle.Right - $rectangle.Left)" -Foregroundcolor Green
			write-Host "     Height = $($rectangle.Bottom - $rectangle.Top)" -Foregroundcolor Green
			$Account.WindowXCoordinates = $rectangle.Left
			$Account.WindowYCoordinates = $rectangle.Top
			$Account.WindowWidth = $rectangle.Right - $rectangle.Left
			$Account.WindowHeight = $rectangle.Bottom - $rectangle.Top
			$Account
		}
		Else {#Leave as is.
			$Account
			write-Verbose "Account $($account.id) ($($account.CharacterName)) isn't running."
		}
	}
	$NewCSV | Export-CSV "$Script:WorkingDirectory\characters.csv" -NoTypeInformation
	Write-Host "`n   Updated CSV with window positions." -foregroundcolor green
	start-sleep -milliseconds 2500
}
Function SetWindowLocations {# Move windows to preferred location/layout
	param(
		[int]$Id,
		[int]$X,
		[int]$Y,
		[int]$Width,
		[int]$Height
	)
	LoadWindowClass
	$handle = (Get-Process -Id $Id).MainWindowHandle
    # Constants for SetWindowPos
    $HWND_TOPMOST = [IntPtr]::Zero # Change this to [IntPtr]::Zero to avoid topmost if you don't want it to be
	#$SWP_NOMOVE = 0x0002
	#$SWP_NOSIZE = 0x0001
	$SWP_SHOWWINDOW = 0x0040
	$SWP_NOREDRAW = 0x0008

    # Restore the window (if minimized)
    [Window]::ShowWindow($handle, 9)
    Start-Sleep -Milliseconds 10

    # Move the window and set its position
   # [Window]::SetWindowPos($handle, $HWND_TOPMOST, $X, $Y, $Width, $Height, $SWP_SHOWWINDOW)
	[Window]::SetWindowPos($handle, $HWND_TOPMOST, $X, $Y, $Width, $Height, $SWP_SHOWWINDOW -bor $SWP_NOREDRAW)
    Start-Sleep -Milliseconds 10

    # Optionally, bring it to the foreground
    [Window]::SetForegroundWindow($handle)
}
Function Options {
	ImportXML
	Clear-Host
	Write-Host "`n This screen allows you to change script config options."
	Write-Host " Note that you can also change these settings (and more) in config.xml."
	Write-Host " Options you can change/toggle below:`n"
	$OptionList = "1","2","3","4","5","6","7","8"
	$XML = Get-Content "$Script:WorkingDirectory\Config.xml"
	Write-Host "`n  $X[38;2;255;165;000;22m2$X[0m - $X[4mSettingSwitcherEnabled$X[0m (Currently $X[38;2;255;165;000;22m$(if($Script:Config.SettingSwitcherEnabled -eq 'True'){'Enabled'}else{'Disabled'})$X[0m)"
	Write-Host "  $X[38;2;255;165;000;22m3$X[0m - $X[4mManualSettingSwitcherEnabled$X[0m (Currently $X[38;2;255;165;000;22m$(if($Script:Config.ManualSettingSwitcherEnabled -eq 'True'){'Enabled'}else{'Disabled'})$X[0m)"
	Write-Host "  $X[38;2;255;165;000;22m4$X[0m - $X[4mRememberWindowLocations$X[0m (Currently $X[38;2;255;165;000;22m$(if($Script:Config.RememberWindowLocations -eq 'True'){'Enabled'}else{'Disabled'})$X[0m)"

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
		Write-Host;Write-Host $OptionsText
		do {
			if ($OptionInteger -eq $True){
				Write-Host "   Enter a number between $X[38;2;255;165;000;22m1$X[0m and $X[38;2;255;165;000;22m99$X[0m or '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
				$AcceptableOptions = 1..99 # Allow user to enter 1 to 99
				$NewOptionValue = (ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27 -TwoDigitAcctSelection $True).tostring()
				$NewValue = $NewOptionValue
			}
			Else {
				Write-Host "   Enter " -nonewline;CommaSeparatedList -NoOr ($OptionsList.keys | sort-object); Write-Host " or '$X[38;2;255;165;000;22mc$X[0m' to cancel: " -nonewline
				$AcceptableOptions = $OptionsList.keys
				$NewOptionValue = (ReadKeyTimeout "" $MenuRefreshRate "c" -AdditionalAllowedKeys 27).tostring()
				$NewValue = $($OptionsList[$NewOptionValue])
			}
			if ($NewOptionValue -notin $AcceptableOptions + "c" + "Esc"){
				Write-Host "   Invalid Input. Please enter one of the options above.`n" -foregroundcolor red
			}
		} until ($NewOptionValue -in $AcceptableOptions + "c" + "Esc")
		if ($NewOptionValue -in $AcceptableOptions){
			if ($NewOptionValue -ne "s" -and $NewOptionValue -ne "r" -and $NewOptionValue -ne "a"){
				try {
					$Pattern = "(<$ConfigName>)([^<]*)(</$ConfigName>)"
					$ReplaceString = '{0}{1}{2}' -f '${1}', $NewValue, '${3}'
					$NewXML = [regex]::Replace($Xml, $Pattern, $ReplaceString)
					$NewXML | Set-Content -Path "$Script:WorkingDirectory\Config.xml"
					return $True
				}
				Catch {
					Write-Host "`n  Was unable to update config :(" -foregroundcolor red
					start-sleep -milliseconds 2500
					return $False
				}
			}
			ElseIf ($NewOptionValue -eq "s"){
				$PositionsRecorded = SaveWindowLocations
				if ($PositionsRecorded -ne $False){#redundant?
					return $False
				}
			}
			ElseIf ($NewOptionValue -eq "r"){
				LoadWindowClass
				CheckActiveAccounts
				if ($null -eq $Script:ActiveAccountsList){
					FormatFunction -text "`nThere are no open games.`nTo reset window positions, you need to launch one or more instances first.`n" -indents 2 -IsWarning
					PressTheAnyKey
					Return $False
				}
				ForEach ($Account in $Script:AccountOptionsCSV){
					if ($account.id -in $Script:ActiveAccountsList.id){
						if ($account.WindowXCoordinates -ne "" -and $account.WindowYCoordinates -ne ""){
							SetWindowLocations -X $Account.WindowXCoordinates -Y $Account.WindowYCoordinates -Width $Account.WindowWidth -height $Account.WindowHeight -Id ($Script:ActiveAccountsList | where-object {$_.id -eq $account.id}).ProcessID | out-null
						}
					}
				}
				FormatFunction -indents 2 -text "Moved game windows back to their saved screen coordinates and reset window sizes.`n" -IsSuccess
				start-sleep -milliseconds 1500
				Return $False
			}
			ElseIf ($NewOptionValue -eq "a"){
				LoadWindowClass
				CheckActiveAccounts
				if ($null -eq $Script:ActiveAccountsList){
					FormatFunction -text "`nThere are no open games.`nTo set window positions to alternative layout, you need to launch one or more instances first.`n" -indents 2 -IsWarning
					PressTheAnyKey
					Return $False
				}
				ForEach ($Account in $Script:AltWindowLayoutCoords){
					if ($account.id -in $Script:ActiveAccountsList.id){
						if ($account.WindowXCoordinates -ne "" -and $account.WindowYCoordinates -ne ""){
							SetWindowLocations -X $Account.WindowXCoordinates -Y $Account.WindowYCoordinates -Width $Account.WindowWidth -height $Account.WindowHeight -Id ($Script:ActiveAccountsList | where-object {$_.id -eq $account.id}).ProcessID | out-null
						}
					}
				}
				FormatFunction -indents 2 -text "Moved game windows to alternative coordinates and window sizes.`n" -IsSuccess
				start-sleep -milliseconds 1500
				Return $False
			}
		}
		else {
			Return $False
		}
	}
	If ($Option -eq "2"){ #SettingSwitcherEnabled
		If ($Script:Config.SettingSwitcherEnabled -eq "False"){
			$Options = @{"1" = "True"}
			$OptionsSubText = "enable"
			$CurrentState = "Disabled"
		}
		Else {
			$Options = @{"1" = "False"}
			$OptionsSubText = "disable"
			$CurrentState = "Enabled"
		}
		$XMLChanged = OptionSubMenu -ConfigName "SettingSwitcherEnabled" -OptionsList $Options -Current $CurrentState `
		-Description "This enables the script to automatically switch which settings file to use when launching the game based on the account you're launching.`nA very cool feature!`nPlease see GitHub for instructions on setting this up/editing settings." `
		-OptionsText "    Choose '$X[38;2;255;165;000;22m1$X[0m' to $OptionsSubText`n"
	}
	ElseIf ($Option -eq "3"){ #ManualSettingSwitcherEnabled
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
		-OptionsText "    Choose '$X[38;2;255;165;000;22m1$X[0m' to $OptionsSubText`n"
	}
	ElseIf ($Option -eq "4"){ #RememberWindowLocations
		If ($Script:Config.RememberWindowLocations -eq "False"){
			$Options = @{"1" = "True"}
			$OptionsSubText = "enable"
			$DescriptionSubText = "`nOnce enabled, return to this menu and choose the '$X[38;2;255;165;000;22ms$X[0m' option to save coordinates of any open game instances."
			$CurrentState = "Disabled"
		}
		Else {
			$Options = @{"1" = "False";"S" = "PlaceholderValue Only :)"} # SaveWindowLocations function used if user chooses "S"
			$OptionsSubText = "disable"
			$OptionsSubTextAgain = "    Choose '$X[38;2;255;165;000;22ms$X[0m' to save current window locations and sizes.`n"
			if ($Script:AccountOptionsCSV | Get-Member -Name "WindowXCoordinates" -MemberType NoteProperty -ErrorAction SilentlyContinue){
				$OptionsSubTextAgain += "    Choose '$X[38;2;255;165;000;22mr$X[0m' to reset window locations and sizes.`n"
				$Options += @{"R" = "PlaceholderValue Only :)"}
				if (Test-Path -Path "$Script:WorkingDirectory\AltLayout.csv"){
					$Script:AltWindowLayoutCoords = import-csv "$Script:WorkingDirectory\AltLayout.csv"
					$OptionsSubTextAgain += "    Choose '$X[38;2;255;165;000;22ma$X[0m' to set window locations to alternative layout.`n"
					$Options += @{"A" = "PlaceholderValue Only :)"}
				}
			}
			$DescriptionSubText = "`nChoosing the '$X[38;2;255;165;000;22ms$X[0m' option will save coordinates (and window sizes) of any open game instances.`nChoosing the '$X[38;2;255;165;000;22mr$X[0m' option will move your windows back to their default placements."
			$CurrentState = "Enabled"
		}
		$XMLChanged = OptionSubMenu -ConfigName "RememberWindowLocations" -OptionsList $Options -Current $CurrentState `
		-Description "For those that have configured the game to launch in windowed mode, this setting is used to make the script move the window locations at launch, so that you never have to rearrange your windows when launching accounts.$DescriptionSubText" `
		-OptionsText "    Choose '$X[38;2;255;165;000;22m1$X[0m' to $OptionsSubText`n$OptionsSubTextAgain"
	}
	else {#go to main menu if no valid option was specified.
		return
	}
	if ($XMLChanged -eq $True){
		Write-Host "   Config Updated!" -foregroundcolor green
		ImportXML
		If ($Option -eq "4" -and $Script:Config.RememberWindowLocations -eq $True -and -not ($Script:AccountOptionsCSV | Get-Member -Name "WindowXCoordinates" -MemberType NoteProperty -ErrorAction SilentlyContinue)){#if this is the first time it's been enabled display a setup message
			Formatfunction -indents 2 -IsWarning -Text "`nYou've enabled RememberWindowsLocations but you still need to set it up. To set this up you need to perform the following steps:"
			FormatFunction -indents 3 -iswarning -SubsequentLineIndents 3 -text "`n1. Open all of your D2r account instances.`n2. Move the window for each game instance to your preferred layout and size."
			FormatFunction -indents 3 -iswarning -SubsequentLineIndents 3 -text "3. Come back to this options menu and go into the 'RememberWindowLocations' setting.`n4. Once in this menu, choose the option 's' to save coordinates of any open game instances."
			FormatFunction -indents 2 -iswarning -text  "`n`nNow when you open these accounts they will open in this screen location each time :)`n"
			PressTheAnyKey
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
Function CheckActiveAccounts {#Note: only works for accounts loaded by the script
	#check if there's any open instances and check the game title window for which account is being used.
	try {
		$Script:ActiveIDs = $Null
		$D2rRunning = $false
		$Script:ActiveIDs = New-Object -TypeName System.Collections.ArrayList
		$Script:ActiveIDs = (Get-Process | Where-Object {$_.processname -eq "D2r" -and $_.MainWindowTitle -match "- Diablo II: Resurrected \(SP\)"} | Select-Object MainWindowTitle).mainwindowtitle.substring(0,2).trim() #find all diablo 2 game windows and pull the account ID from the title
		$Script:D2rRunning = $true
		Write-Verbose "Running Instances."
	}
	catch {#if the above fails then there are no running D2r instances.
		$Script:D2rRunning = $false
		Write-Verbose "No Running Instances."
		$Script:ActiveIDs = ""
	}
	if ($Script:D2rRunning -eq $True){
		$Script:ActiveAccountsList = New-Object -TypeName System.Collections.ArrayList
		ForEach ($ActiveID in $ActiveIDs){#Build list of active accounts that we can omit from being selected later
			$ActiveAccountDetails = $Script:AccountOptionsCSV | where-object {$_.id -eq $ActiveID}
			$ActiveAccount = New-Object -TypeName psobject
			$ActiveAccount | Add-Member -MemberType NoteProperty -Name ID -Value $ActiveAccountDetails.ID
			$ActiveAccount | Add-Member -MemberType NoteProperty -Name AccountName -Value $ActiveAccountDetails.CharacterName
			$InstanceProcessID = (Get-Process | Where-Object {$_.processname -eq "D2r" -and ($_.MainWindowTitle -match "$($ActiveAccountDetails.ID) - $($ActiveAccountDetails.CharacterName)" -and $_.MainWindowTitle -match "- Diablo II: Resurrected \(SP\)")} | Select-Object ID).id
			write-verbose "  ProcessID for $($ActiveAccountDetails.ID) - $($ActiveAccountDetails.CharacterName) is $InstanceProcessID"
			$ActiveAccount | Add-Member -MemberType NoteProperty -Name ProcessID -Value $InstanceProcessID
			[VOID]$Script:ActiveAccountsList.Add($ActiveAccount)
		}
	}
	else {
		$Script:ActiveAccountsList = $Null
	}
}
Function DisplayActiveAccounts {
	Write-Host
	if ($Script:ActiveAccountsList.id -ne ""){
		$PlayTimeHeader = "Hours Played   "
		Write-Host ("  ID   " + $PlayTimeHeader + "Character Name") #Header
	}
	else {
		Write-Host "  ID   Character Name"
	}
	ForEach ($AccountOption in ($Script:AccountOptionsCSV | Sort-Object -Property @{ #Try sort by number first (needed for 2 digit ID's), then sort by character.
		Expression = {
			$intValue = [int]::TryParse($_.ID, [ref]$null) # Try to convert the value to an integer
			if ($intValue){# If it's not null then it's a number, so return it as an integer for sorting.
				[int]$_.ID
			}
			else {# If it's not a number, return a character and sort that.
				[char]$_.ID
			}
		}
	}))
	{
		if ($AccountOption.CharacterName.length -gt 28){ #later we ensure that strings longer than 28 chars are cut short so they don't disrupt the display.
			$AccountOption.CharacterName = $AccountOption.CharacterName.Substring(0, 28)
		}
		if ($AccountOption.ID.length -ge 2){#keep table formatting looking lovely if some crazy user has 10+ accounts.
			$IDIndent = ""
		}
		else {
			$IDIndent = " "
		}
		try {
			$AcctPlayTime = (" " + (($time =([TimeSpan]::Parse($AccountOption.TimeActive))).hours + ($time.days * 24)).tostring() + ":" + ("{0:D2}" -f $time.minutes) + "   ")  # Add hours + (days*24) to show total hours, then show ":" followed by minutes
		}
		catch {#if account hasn't been opened yet.
			$AcctPlayTime = "   0   "
			Write-Debug "Account not opened yet."
		}
		if ($AcctPlayTime.length -lt 15){#formatting. Depending on the amount of characters for this variable push it out until it's 15 chars long.
			while ($AcctPlayTime.length -lt 15){
				$AcctPlayTime = " " + $AcctPlayTime
			}
		}
		if ($AccountOption.id -in $Script:ActiveAccountsList.id){ #if account is currently active
			Write-Host ("  " + $IDIndent + $AccountOption.ID + "   " + $AcctPlayTime  + $AccountOption.CharacterName + " - Account Active.") -foregroundcolor yellow
		}
		else {#if account isn't currently active
			Write-Host ("  " + $IDIndent + $AccountOption.ID + "   " + $AcctPlayTime + $AccountOption.CharacterName) -foregroundcolor green
		}
	}
}
Function Menu {
	Clear-Host
	if ($Script:ScriptHasBeenRun -ne $true){
		Write-Host ("  You have quite a treasure there in that Horadric multibox script v" + $Currentversion)
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
		if ($Script:AccountID -eq "i"){
			Inventory #show stats
			$Script:AccountID = "r"
		}
		if ($Script:AccountID -eq "o"){ #options menu
			Options
			$Script:AccountID = "r"
		}
		if ($Script:AccountID -eq "s"){
			if ($Script:AskForSettings -eq $True){
				Write-Host "  Manual Setting Switcher Disabled." -foregroundcolor Green
				$Script:AskForSettings = $False
			}
			else {
				Write-Host "  Manual Setting Switcher Enabled." -foregroundcolor Green
				$Script:AskForSettings = $True
			}
			Start-Sleep -milliseconds 1550
			$Script:AccountID = "r"
		}
		if ($Script:AccountID -eq "g"){#silly thing to replicate in game chat gem.
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
			$Script:AccountID = "r"
		}
		if ($Script:AccountID -eq "r"){#refresh
			Clear-Host
			Notifications -check $True
			BannerLogo
			QuoteRoll
		}
		CheckActiveAccounts
		DisplayActiveAccounts
		$OpenD2LoaderInstances = Get-WmiObject -Class Win32_Process | Where-Object { $_.name -eq "powershell.exe" -and $_.commandline -match $Script:ScriptFileName} | Select-Object name,processid,creationdate | Sort-Object creationdate -descending
		if ($OpenD2LoaderInstances.length -gt 1){#If there's more than 1 D2loader.ps1 script open, close until there's only 1 open to prevent the time played accumulating too quickly.
			ForEach ($Process in $OpenD2LoaderInstances[1..($OpenD2LoaderInstances.count -1)]){
				Stop-Process -id $Process.processid -force #Closes oldest running d2loader script
			}
		}
		if ($Script:ActiveAccountsList.id.length -ne 0){#if there are active accounts open add to total script time
			#Add time for each account that's open
			$Script:AccountOptionsCSV = import-csv "$Script:WorkingDirectory\characters.csv"
			$AdditionalTimeSpan = New-TimeSpan -Start $Script:StartTime -End (Get-Date) #work out elapsed time to add to characters.csv
			ForEach ($AccountID in $Script:ActiveAccountsList.id |Sort-Object){ #$Script:ActiveAccountsList.id
				$AccountToUpdate = $Script:AccountOptionsCSV | Where-Object {$_.ID -eq $accountID}
				if ($AccountToUpdate){
					try {#get current time from csv and add to it
						$AccountToUpdate.TimeActive = [TimeSpan]::Parse($AccountToUpdate.TimeActive) + $AdditionalTimeSpan
					}
					Catch {#if CSV hasn't been populated with a time yet.
						$AccountToUpdate.TimeActive = $AdditionalTimeSpan
					}
				}
				try {
					$Script:AccountOptionsCSV | Export-Csv -Path "$Script:WorkingDirectory\characters.csv" -NoTypeInformation #update characters.csv with the new time played.
				}
				Catch {
					$WriteAcctCSVError = $True
				}
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
		$Script:StartTime = Get-Date #restart timer for session time and account time.
		$Script:AcceptableValues = New-Object -TypeName System.Collections.ArrayList
		$Script:TwoDigitIDsUsed = $False
		ForEach ($AccountOption in $Script:AccountOptionsCSV){
			if ($AccountOption.id -notin $Script:ActiveAccountsList.id){
				$Script:AcceptableValues = $AcceptableValues + ($AccountOption.id) #+ "x"
				if ($AccountOption.id.length -eq 2){
					$Script:TwoDigitIDsUsed = $True
				}
			}
		}
		$accountoptions = ($Script:AcceptableValues -join  ", ").trim()
		if ($Script:MovedWindowLocations -ge 1){ # Tidy up any old jobs created for moving windows.
			Get-Job | Where-Object { $Script:JobIDs -contains $_.Id -and $_.state -ne "Running"} | Remove-Job -Force
			$Script:JobIDs = @()
			$Script:MovedWindowLocations = 0
		}
		do {
			Write-Host
			if ($accountoptions.length -gt 0){#if there are unopened account options available
				if ($accountoptions.length -le 24){ #if so many accounts are available to be used that it's too long and impractical to display all the individual options.
					Write-Host ("  Select which account to sign into: " + "$X[38;2;255;165;000;22m$accountoptions$X[0m")
					Write-Host "  Alternatively choose from the following menu options:"
				}
				Else {
					Write-Host " Enter the ID# of the account you want to sign into."
					Write-Host " Alternatively choose from the following menu options:"
				}
			}
			else {#if there aren't any available options, IE all accounts are open
				Write-Host " All Accounts are currently open!" -foregroundcolor yellow
			}
			#Write-Host "  '$X[38;2;255;165;000;22mr$X[0m' to Refresh"
			if ($Script:Config.ManualSettingSwitcherEnabled -eq $true){
				$ManualSettingSwitcherOption = "s"
				Write-Host "  '$X[38;2;255;165;000;22mr$X[0m' to Refresh, $X[38;2;255;165;000;22mo$X[0m' for config options, '$X[38;2;255;165;000;22ms$X[0m' to toggle the Manual Setting Switcher, "
				Write-Host "  '$X[38;2;255;165;000;22mi$X[0m' for info or '$X[38;2;255;165;000;22mx$X[0m' to $X[38;2;255;000;000;22mExit$X[0m: "-nonewline
			}
			Else {
				$ManualSettingSwitcherOption = $null
				Write-Host "  '$X[38;2;255;165;000;22mr$X[0m' to Refresh, $X[38;2;255;165;000;22mo$X[0m' for config options, '$X[38;2;255;165;000;22mi$X[0m' for info or '$X[38;2;255;165;000;22mx$X[0m' to $X[38;2;255;000;000;22mExit$X[0m: " -nonewline
			}
			if ($Script:TwoDigitIDsUsed -eq $True){
				$Script:AccountID = ReadKeyTimeout "" $MenuRefreshRate "r" -TwoDigitAcctSelection $True #$MenuRefreshRate represents the refresh rate of the menu in seconds (30). if no button is pressed, send "r" for refresh.
			}
			else {
				$Script:AccountID = ReadKeyTimeout "" $MenuRefreshRate "r" #$MenuRefreshRate represents the refresh rate of the menu in seconds (30). if no button is pressed, send "r" for refresh.
			}
			if ($Script:AccountID -notin ($Script:AcceptableValues + "x" + "r" + "g" + "i" + "o" + $ManualSettingSwitcherOption) -and $Null -ne $Script:AccountID){
				Write-Host " Invalid Input. Please enter one of the options above." -foregroundcolor red
				$Script:AccountID = $Null
			}
		} until ($Null -ne $Script:AccountID)
		if ($Null -ne $Script:AccountID){
			if ($Script:AccountID -eq "x"){
				Write-Host "`n Good day to you partner :)" -foregroundcolor yellow
				Start-Sleep -milliseconds 486
				Exit
			}
			$Script:AccountChoice = $Script:AccountOptionsCSV | where-object {$_.id -eq $Script:AccountID} #filter out to only include the account we selected.
		}
		$Script:RunOnce = $True
	} until ($Script:AccountID -ne "r" -and $Script:AccountID -ne "g" -and $Script:AccountID -ne "s" -and $Script:AccountID -ne "i" -and $Script:AccountID -ne "o")
}
Function Processing {
	$Script:AccountFriendlyName = $Script:AccountChoice.CharacterName.tostring()
	#Open diablo with parameters
		# IE, this is essentially just opening D2r like you would with a shortcut target of "C:\Program Files (x86)\Battle.net\Games\Diablo II Resurrected\D2R.exe" -username <yourusername -password <yourPW> -address <SERVERaddress>
	$CustomLaunchArguments = ($Config.CustomLaunchArguments).replace("`"","").replace("'","") #clean up arguments in case they contain quotes (for folks that have used excel to edit characters.csv).
	$arguments = (" -address xxx" + " " + $CustomLaunchArguments).tostring() #force offline mode by giving it a garbage region to connect to.
	if ($Config.ForceWindowedMode -eq $true){#starting with forced window mode sucks, but someone asked for it.
		$arguments = $arguments + " -w"
	}
	#Switch Settings file to load D2r from.
	if ($Config.SettingSwitcherEnabled -eq $True -and $Script:AskForSettings -ne $True){#if user has enabled the auto settings switcher.
		$SettingsProfilePath = ("C:\Users\" + $Env:UserName + "\Saved Games\Diablo II Resurrected\")
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
							FormatFunction -Text "Using standard settings path. Couldn't find Modinfo.json in '$($Config.GamePath)\Mods\$ModName\$ModName.mpq'" -IsWarning
							start-sleep -milliseconds 1500
						}
					}
					If ($Null -eq $Modinfo){
						Write-Verbose " No Custom Save Path Specified for this mod."
					}
					ElseIf ($Modinfo -ne "../"){
						$SettingsProfilePath += "mods\$Modinfo\"
						if (-not (Test-Path $SettingsProfilePath)){
							Write-Host " Mod Save Folder doesn't exist yet. Creating folder..."
							New-Item -ItemType Directory -Path $SettingsProfilePath -ErrorAction stop | Out-Null
							Write-Host " Created folder: $SettingsProfilePath" -ForegroundColor Green
						}
						Write-Host " Mod: $ModName detected. Using custom path for settings.json." -ForegroundColor Green
						Write-Verbose " $SettingsProfilePath"
					}
					Else {
						Write-Verbose " Mod used but save path is standard."
					}
				}
				Catch {
					Write-Verbose " Mod used but custom save path not specified."
				}
			}
			else {
				Write-Host " Couldn't detect Mod name. Standard path to be used for settings.json." -ForegroundColor Red
			}
		}
		$SettingsJSON = ($SettingsProfilePath + "Settings.json")
		if ((Test-Path -Path ($SettingsProfilePath + "Settings.json")) -eq $true){ #check if settings.json does exist in the savegame path (if it doesn't, this indicates first time launch or use of a new single player mod).
			ForEach ($id in $Script:AccountOptionsCSV){#create a copy of settings.json file per account so user doesn't have to do it themselves
				if ((Test-Path -Path ($SettingsProfilePath + "Settings" + $id.id +".json")) -ne $true){#if somehow settings<ID>.json doesn't exist yet make one from the current settings.json file.
					try {
						Copy-Item $SettingsJSON ($SettingsProfilePath + "Settings"+ $id.id + ".json") -ErrorAction Stop
					}
					catch {
						FormatFunction -Text "`nCouldn't find settings.json in $SettingsProfilePath" -IsError
						if ($Config.CustomLaunchArguments -match "-mod"){
							Break
						}
						Else {
							Write-Host " Start the game normally (via Bnet client) & this file will be rebuilt." -foregroundcolor red
						}
						Write-Host
						PressTheAnyKeyToExit
					}
				}
			}
			try {
				Copy-item ($SettingsProfilePath + "settings"+ $Script:AccountID + ".json") $SettingsJSON -ErrorAction Stop #overwrite settings.json with settings<ID>.json (<ID> being the account ID). This means any changes to settings in settings.json will be lost the next time an account is loaded by the script.
				$CurrentLabel = ($Script:AccountOptionsCSV | where-object {$_.id -eq $Script:AccountID}).CharacterName
				formatfunction -text ("Custom game settings (settings" + $Script:AccountID + ".json) being used for " + $CurrentLabel) -success
				Start-Sleep -milliseconds 133
			}
			catch {
				FormatFunction -Text "Couldn't overwrite settings.json for some reason. Make sure you don't have the file open!" -IsError
				PressTheAnyKey
			}
		}
	}
	if ($Script:AskForSettings -eq $True){#steps go through if user has toggled on the manual setting switcher ('s' in the menu).
		$SettingsProfilePath = ("C:\Users\" + $Env:UserName + "\Saved Games\Diablo II Resurrected\")
		$SettingsJSON = ($SettingsProfilePath + "Settings.json")
		$files = Get-ChildItem -Path $SettingsProfilePath -Filter "settings.*.json"
		$Counter = 1
		if ((Test-Path -Path ($SettingsProfilePath+ "Settings" + $id.id +".json")) -ne $true){#if somehow settings<ID>.json doesn't exist yet make one from the current settings.json file.
			try {
				Copy-Item $SettingsJSON ($SettingsProfilePath + "Settings"+ $id.id + ".json") -ErrorAction Stop
			}
			catch {
				Write-Host "`n Couldn't find settings.json in $SettingsProfilePath" -foregroundcolor red
				Write-Host " Please start the game normally (via Bnet client) & this file will be rebuilt." -foregroundcolor red
				PressTheAnyKeyToExit
			}
		}
		$SettingsDefaultOptionArray = New-Object -TypeName System.Collections.ArrayList #Add in an option for the default settings file (if it exists, if the auto switcher has never been used it won't appear.
		$SettingsDefaultOption = New-Object -TypeName psobject
		$SettingsDefaultOption | Add-Member -MemberType NoteProperty -Name "ID" -Value $Counter
		$SettingsDefaultOption | Add-Member -MemberType NoteProperty -Name "Name" -Value ("Default - settings"+ $Script:AccountID + ".json")
		$SettingsDefaultOption | Add-Member -MemberType NoteProperty -Name "FileName" -Value ("settings"+ $Script:AccountID + ".json")
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
				if ($SettingsChoice -eq ""){
					$SettingsChoice = 1
				}
				Write-Host
				if ($SettingsChoice.tostring() -notin $SettingsFileOptions.id + $SettingsCancelOption + "Esc"){
					Write-Host "  Invalid Input. Please enter one of the options above." -foregroundcolor red
					$SettingsChoice = ""
				}
			} until ($SettingsChoice.tostring() -in $SettingsFileOptions.id + $SettingsCancelOption + "Esc")
			if ($SettingsChoice -ne "c" -and $SettingsChoice -ne "Esc"){
				$SettingsToLoadFrom = $SettingsFileOptions | where-object {$_.id -eq $SettingsChoice.tostring()}
				try {
					Copy-item ($SettingsProfilePath + $SettingsToLoadFrom.FileName) -Destination $SettingsJSON #-ErrorAction Stop #overwrite settings.json with settings<Name>.json (<Name> being the name of the config user selects). This means any changes to settings in settings.json will be lost the next time an account is loaded by the script.
					$CurrentLabel = ($Script:AccountOptionsCSV | where-object {$_.id -eq $Script:AccountID}).CharacterName
					Write-Host (" Custom game settings (" + $SettingsToLoadFrom.Name + ") being used for " + $CurrentLabel + "`n") -foregroundcolor green
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
		#Start Game
		KillHandle | out-null
		$process = Start-Process "$Gamepath\D2R.exe" -ArgumentList "$arguments" -PassThru 
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
		$rename = ($Script:AccountID + " - " + $Script:AccountFriendlyName + " - Diablo II: Resurrected (SP)")
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
		If ($Script:Config.RememberWindowLocations -eq $True){ #If user has enabled the feature to automatically move game Windows to preferred screen locations.
			if ($Script:AccountChoice.WindowXCoordinates -ne "" -and $Script:AccountChoice.WindowYCoordinates -ne "" -and $Null -ne $Script:AccountChoice.WindowXCoordinates -and $Null -ne $Script:AccountChoice.WindowYCoordinates -and $Script:AccountChoice.WindowWidth -ne "" -and $Script:AccountChoice.WindowHeight -ne "" -and $Null -ne $Script:AccountChoice.WindowWidth -and $Null -ne $Script:AccountChoice.WindowHeight){ #Check if the account has had coordinates saved yet.
				$GetLoadWindowClassFunc = $(Get-Command LoadWindowClass).Definition
				$GetSetWindowLocationsFunc = $(Get-Command SetWindowLocations).Definition
				$JobID = (Start-Job -ScriptBlock { # Run this in a background job so we don't have to wait for it to complete
					start-sleep -milliseconds 2024 # We need to wait for about 2 seconds for game to load as if we move it too early, the game itself will reposition the window. Absolute minimum is 420 milliseconds (funnily enough). Delay may need to be a bit higher for people with wooden computers.
					Invoke-Expression "function LoadWindowClass {$using:GetLoadWindowClassFunc}"
					Invoke-Expression "function SetWindowLocations {$using:GetSetWindowLocationsFunc}"
					SetWindowLocations -x $Using:AccountChoice.WindowXCoordinates -y $Using:AccountChoice.WindowYCoordinates -Width $Using:AccountChoice.WindowWidth -height $Using:AccountChoice.WindowHeight -Id $Using:process.id
				}).id
				$Script:MovedWindowLocations ++
				$Script:JobIDs += $JobID
			}
			Else { #Show a warning if user has RememberWindowLocations but hasn't configured it for this account yet.
				FormatFunction -iswarning -text "`n'RememberWindowLocations' config is enabled but can't move game window to preferred location as coordinates need to be defined for the account first.`n`nTo setup follow the quick steps below:"
				FormatFunction -iswarning -indents 1 -SubsequentLineIndents 3 -text "1. Open all of your D2r account instances.`n2. Move the window for each game instance to your preferred layout."
				FormatFunction -iswarning -indents 1 -SubsequentLineIndents 3 -text "3. Go to the options menu in the script and go into the 'RememberWindowLocations' setting.`n4. Once in this menu, choose the option 's' to save coordinates of any open game instances."
				FormatFunction -iswarning -text  "`nNow when you open these accounts they will open in this screen location each time :)`n"
				PressTheAnyKey
			}
		}
		if ($Script:MovedWindowLocations -ge 1){
			FormatFunction -IsSuccess -text "Moved game window to preferred location."
			Start-Sleep -milliseconds 750
		}
		Write-Host "`nGood luck hero..." -foregroundcolor magenta
		Start-Sleep -milliseconds 1000
		$Script:ScriptHasBeenRun = $true
	}
}
InitialiseCurrentStats
#CheckForUpdates
ImportXML
ValidationAndSetup
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
