# VARIABLES
param([string]$repoDir, [string]$macroName)
$fixedRepoDir=$repoDir -replace '/','\'

##### Notes on running macros from server:
## the Template with the Macro needs to be in the Word Start folder on the server
## the macro project name should not be specified, but include module name: 'module.macro'

# GET ACTIVE WORD OBJECT IF ALREADY RUNNING, ELSE CREATE NEW
# Vba needs to close template if already open, but a new process won't have those docs
$ProcessActive = Get-Process winword -ErrorAction SilentlyContinue
if($ProcessActive -ne $null) {
    $wordOpen = $true
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
} else {
    $wordopen = $false
    $word = new-object -comobject word.application
}

# RUN THA MACRO
#this one for running via batch (deploy) script
$word.run($macroName, [ref]$fixedRepoDir)

#this one for calling direct from cmd line
#	$word.run($macroName, $fixedRepoDir) 				

# KILL WORD PROCESS IF IT WASN'T RUNNING TO START
if($wordOpen -eq $false) {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
}