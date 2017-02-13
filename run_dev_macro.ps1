# VARIABLES
param([string]$macroName, [string]$cmd, [string]$repoDir)

$fixedRepoDir=$repoDir -replace '/','\'
Write-Host "Running macro $macroName on $fixedRepoDir"

##### Notes on running macros from server:
## the Template with the Macro needs to be in the Word Start folder on the server
## the macro project name should not be specified, but include module name: 'module.macro'

# GET ACTIVE WORD OBJECT IF ALREADY RUNNING, ELSE CREATE NEW
# Vba needs to close template if already open, but a new process won't have those docs
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
    $word.run($macroName, $fixedRepoDir, $cmd)
}
catch [System.Management.Automation.MethodException]
{
    $word.run($macroName, [ref]$fixedRepoDir, [ref]$cmd) 
}
			
# KILL WORD PROCESS IF IT WASN'T RUNNING TO START
if($wordOpen -eq $false) {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
}