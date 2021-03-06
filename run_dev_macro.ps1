param([string]$macroName)
# PARAMETERS
# Always need $macroName, the macro to be run as a string, in form "Module.Procedure".
# The project/template with that macro must be in Word STARTUP dir.

# Any arguments for *the VBA macro* should be passed unnamed, and we'll pass them to the 
# macro with the $args[] array (which only contains unnamed parameters).

# GET ACTIVE WORD OBJECT IF ALREADY RUNNING, ELSE CREATE NEW
# Vba needs to close template if already open, but a new process won't have those docs,
# so it won't have access to close them.
$ProcessActive = Get-Process winword -ErrorAction SilentlyContinue
if($ProcessActive -ne $null) {
    $wordOpen = $true
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
} else {
    $wordOpen = $false
    $word = new-object -comobject word.application
}

# BUILD STRING OF ARGS TO PASS TO MACRO
# Because we have an unknown number of arguments for the VBA macro (and we might need to
# add [ref] to all but the macro name), we need to build those arguments as a string and
# use Invoke-Expression to run the macro.

# make sure we've got a macro to run, and start the word.run args list with it.
if (!$macroName) {
    Write-Warning "You must pass a macro name."
} 


# if args were passed for the VBA macro, append them to the string

$var = if ($args.Count -gt 0) {
    # Test if we need to use [ref] or not. Must have this testing macro in STARTUP
    try 
    { 
        $word.run("VbaDev.TestRef", "StringVariable")
        $refStr = ""
    }
    catch [System.Management.Automation.MethodException]
    {
        $word.run("VbaDev.TestRef", [ref]"StringVariable")
        $refStr = "[ref]"
    }
    
    # String passed to Invoke-Expression needs to be exactly what you'd input 
    # to run independently, so need to add literal quotes around string args
    $delim = "`", $refStr`""
    
    # Build string of args FOR THE MACRO, to append to macro name in run command
    $macroArgs = "$delim" + "$($args -join $delim)"

    # could maybe do arglist here too and pass single string to function below?
} else {
    $macroArgs = ""
}

$runArgs = "`"$macroName" + "$macroArgs" + "`""

# Build final string to pass to Invoke-Expression. Must use single quotes and $() for
# $word object so it *doesn't* expand, and double quotes for $macroArgs variable so
# it *does* expand.
$doThis = '$($word).' + "run($runArgs)"

# RUN THA MACRO!
# Store output and error in separate variables, so we can send them to separate streams explicitly
Invoke-Expression "$doThis" -OutVariable macroOutput -ErrorVariable errorOutput

# KILL WORD PROCESS IF IT WASN'T RUNNING TO START
$var = if($wordOpen -eq $false) {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
}

# Exit otherwise powershell.exe keeps running forever
Exit