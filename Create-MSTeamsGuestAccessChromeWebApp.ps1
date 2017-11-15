
# Create a Chrome "app" for Each of your Microsoft Teams Guest Access Teams
# Use entirely at your own risk
# Tom Arbuthnot - TomTalks.uk


# Check if we can find Chrome Installed
IF ((Test-Path "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe") -eq $True)
    {


$ShortCutName = Read-Host -Prompt "Enter ShortCut Name For your Microsoft Teams Guest Web App e.g. 'Customer X Team'"

Write-Host ""

Write-Host "Enter the Tenant ID of the Guest Team, need help finding this? Go here for instructions http://tom.qa/TeamsGuestTenantID"
$tenantID = Read-Host -Prompt 'Enter the Tenant ID of the Guest Team:'

Write-Host ""

$ChromeProfileLocation = Read-Host -Prompt "Enter a location to save Chrome Profiles for these web apps, e.g. 'C:\MSTeamsProfiles'"

Write-Host ""

# Add backslash to the path if user misses it:
IF ($ChromeProfileLocation[-1] -ne "\")
    {
    $ChromeProfileLocation = $ChromeProfileLocation + "\"
    }

$msg = "Do you want the $ShortCutName Guest Team to automatically launch at Windows Startup? ( Y or N )"
do {
    $response = Read-Host -Prompt $msg
    if ($response -eq 'y') {
    }
} until ($response -eq 'n' -or $response -eq 'y')

Write-Host "Please Wait, chrome profile and desktop shortcut being created, Please wait 10 seconds"

Write-Host ""

# Create profile

$profilelocation = ("$ChromeProfileLocation" + "$tenantID")

$chromeexe = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"

# Start Chrome mimimised to create a fresh profile
$id = Start-Process -FilePath  "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe" -ArgumentList "--user-data-dir=`"$profilelocation`" -first-run" -PassThru -WindowStyle Minimized

Start-Sleep -Seconds 5

# Kill Chrome process
Stop-Process $id


# Create Desktop Shortcut

  # Build the Target String, has to be build in sections or the pareser gets confused
        $Arguements = $null
      

        [string]$Arguements = "--user-data-dir=`"$profilelocation`" --app=`"https://teams.microsoft.com/_?tenantId=$tenantID`""
        
        $ws = New-Object -com WScript.Shell
        $Dt = $ws.SpecialFolders.Item("Desktop")
        $Scp = Join-Path -Path $Dt -ChildPath "$ShortCutName.lnk"
        # $sc.IconLocation = "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe,0"
        $Sc = $ws.CreateShortcut($Scp)
        $Sc.TargetPath = "$chromeexe"
        $sc.Arguments = "$Arguements"
        $Sc.Description = "$ShortCutName"
        $Sc.Save()


Write-Host "Desktop Shortcut $ShortCutName Created"

Write-Host ""

Write-Host "Please sign in using your credentials"

# Start Chrome in app mode
Start-Process -FilePath  "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe" -ArgumentList $Arguements -WindowStyle Normal

$Desktop = [Environment]::GetFolderPath("Desktop")
If ($response -eq 'y')
    {
    Copy-Item $Desktop\$ShortCutName.lnk -Destination "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\"
    }

    }
    else
    {
    Write-Host ""
    Write-Host "Chrome.exe not found at "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe" - Please install Chrome or Edit Path in Script"
    Write-Host ""
    }