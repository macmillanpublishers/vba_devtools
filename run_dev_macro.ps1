# PARAMETERS
# Do NOT use named parameters, so we can accept any number for macros.
# First argument should always be the macro to be run as a string, in form Module.Procedure
# and the project/template with that macro should be in Word STARTUP.
# Any subsequent args will be passed to VBA as is.

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

# RUN THA MACRO!
# Using Switch statement to pass exact number of args passed to this script. Technically you can pass up
# to 30 args to a macro, plus the macro name, but that seems excessive for now.

# try / catch so we don't have to switch between 'em
try 
{ 
    switch ($args.Length) {
        1 {$outValue = $word.run($args[0])}
        2 {$outValue = $word.run($args[0], $args[1])}
        3 {$outValue = $word.run($args[0], $args[1], $args[2])}
        4 {$outValue = $word.run($args[0], $args[1], $args[2], $args[3])}
        5 {$outValue = $word.run($args[0], $args[1], $args[2], $args[3], $args[4])}
        default {$outValue = "No macro. " + $args.Length + " args passed."}
    }
}
catch [System.Management.Automation.MethodException]
{
    switch ($args.Length) {
        1 {$outValue = $word.run($args[0])}
        2 {$outValue = $word.run($args[0], [ref]$args[1])}
        3 {$outValue = $word.run($args[0], [ref]$args[1], [ref]$args[2])}
        4 {$outValue = $word.run($args[0], [ref]$args[1], [ref]$args[2], [ref]$args[3])}
        5 {$outValue = $word.run($args[0], [ref]$args[1], [ref]$args[2], [ref]$args[3], [ref]$args[4])}
        default {$outValue = "No macro. " + $args.Length + " args passed."}
    }
}

# KILL WORD PROCESS IF IT WASN'T RUNNING TO START
if($wordOpen -eq $false) {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
}

return $outValue

# Exit otherwise powershell.exe keeps running forever
Exit