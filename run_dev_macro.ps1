# VARIABLES
# Do NOT use named parameters, so we can accept any number for macros.
# First argument should always be the macro to be run as a string, in form Module.Procedure
# and the project/template with that macro should be in Word STARTUP.

# Second argument should always be full WINDOWS path to dir or file macro is running on.
# Any subsequent args will be passed to VBA as is.

$macroName = $args[0]
$workingDir = $args[1]  # just getting for msg below, will pass to macro too, though

# New array with macroName arg removed, to pass to macro
$macroArgs = $args[1..$args.Length]

Write-Host "Running macro $macroName on $workingDir"

# GET ACTIVE WORD OBJECT IF ALREADY RUNNING, ELSE CREATE NEW
# Vba needs to close template if already open, but a new process won't have those docs,
# So it won't have access to close them.
$ProcessActive = Get-Process winword -ErrorAction SilentlyContinue
if($ProcessActive -ne $null) {
    $wordOpen = $true
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
} else {
    $wordopen = $false
    $word = new-object -comobject word.application
}

# RUN THA MACRO
# try / catch so we don't have to switch between 'em
try 
{ 
    $word.run($macroName, $macroArgs)
}
catch [System.Management.Automation.MethodException]
{
    $word.run($macroName, [ref]$macroArgs) 
}
			
# KILL WORD PROCESS IF IT WASN'T RUNNING TO START
if($wordOpen -eq $false) {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
}

# Exit otherwise powershell.exe keeps running forever
Exit