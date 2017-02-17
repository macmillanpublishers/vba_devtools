# Must run AS ADMINISTRATOR 

# Installs vba_devtools macros and assorted settings
# Must run as administrator. Will exit if you aren't.

# PARAMS: value for "installType" in config file. Must be one of the "location" options
param([String]$installType)

# functions
function ConvertTo-Json20([object] $item){
    add-type -assembly system.web.extensions
    $ps_js=new-object system.web.script.serialization.javascriptSerializer
    return $ps_js.Serialize($item)
}

function ConvertFrom-Json20([object] $item){ 
    add-type -assembly system.web.extensions
    $ps_js=new-object system.web.script.serialization.javascriptSerializer

    #The comma operator is the array construction operator in PowerShell
    return ,$ps_js.DeserializeObject($item)
}

# Check if passed installType variable, if not exit
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning "You do not have Administrator rights to run this script.`nPlease re-run this script as an Administrator."
    Break
}

# SET UP VARIABLES
# Word startup dir 
$startupDir = "$Env:AppData\Microsoft\Word\STARTUP"
# the current directory (of this script)
$PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
# name of the repo
$repoName = Split-Path $PSScriptRoot -Leaf
# config file name
$configName = ($repoName + "_config.json")
# path to config in repo
$configRepoFullName = ($PSScriptRoot + "\" + $configName)
# path to config in startup
$configStartupFullName = ($startupDir + "\" + $configName)
# template file name
$templateName = ($repoName + ".dotm")


# Read config file
Write-Host "Checking config file..."
$configString = Get-Content "$configRepoFullName"
$config = ConvertFrom-Json20 $configString

# Check if we got an argument
IF([string]::IsNullOrEmpty($installType)) {            
    Write-Warning "Must include installType parameter"
           
} else {            
    # Check that the value exists in the config file
    if ($config.Get_Item("location").ContainsKey("$installType") -eq $false) {
        Write-Warning "That install type doesn't exist in the config file."
        Break    
    }          
}

Write-Host "Updating config file..."
# Add installType to local config file
$config["installType"] = $installType

# Convert back to a string and write to startup
$content = ConvertTo-Json20 $config
[IO.File]::WriteAllLines($configStartupFullName, $content)

# Copy devtools template to startup too
Copy-Item "$PSScriptRoot\$templateName" "$startupDir\$templateName"

Write-Host "Setting environment variables..."
# Set required environment variables
[Environment]::SetEnvironmentVariable("WordStartup", "$startupDir", "User")
[Environment]::SetEnvironmentVariable("VbaDebug", "true", "User")
# this so we can run our custom git commands
[Environment]::SetEnvironmentVariable("Path", $env:Path + ";" + "$PSScriptRoot", [System.EnvironmentVariableTarget]::User )

Write-Host ""
Write-Host "Install complete!"