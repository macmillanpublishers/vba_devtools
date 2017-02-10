Attribute VB_Name = "dev"
Option Explicit
Option Base 1

' ====== USE ==================================================================
' For help using git with VBA development.
' See docs about devtools installation and command-line scripts.

' ===== DEPENDENCIES ==========================================================
' VBA code modules that should be tracked in current repo must be added to git
' as git submodules.

' PC only, because this code requires the MS RegEx lib as a reference, which
' isn't available on Mac. Also saving templates on Mac causes all kinds of
' nonsense so you shouldn't be doing primary dev there anyway.

' ====== WARNING ==============================================================
' advice from http://www.cpearson.com/excel/vbe.aspx :
' "Many VBA-based computer viruses propagate themselves by creating and/or modifying
' VBA code. Therefore, many virus scanners may automatically and without warning or
' confirmation delete modules that reference the VBProject object, causing a permanent
' and irretrievable loss of code. Consult the documentation for your anti-virus
' software for details."
'
' So be sure to export and commit often!

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'      PUBLIC PROCEDURES
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



' ===== VbaStatus =============================================================
' Designed to be called from an external script that will pass working dir. If
' this all goes according to plan, you can run this instead of git status, and
' all local files/code will be exported to repo.

' ASSUMPTIONS
' You're working on code from file in local directory

' RETURNS: Boolean
' True = exporting and copying was successful
' False = not successful

' TODO:
' * Return more detailed info?

Public Function VbaStatus(WorkingDir As String) As Boolean
  Dim objStatusRepo As Repository
  Set objStatusRepo = Factory.CreateRepository(Path:=WorkingDir)

' Export code files to repo
  objStatusRepo.UpdateCodeInRepo

' This isn't doing anything, but maybe we'll need to add error handling later
  VbaStatus = True
  
End Function

' ===== VbaCheckout ===========================================================
' Macro doc files is in two places ("repo" and "local"), this copies one to
' overwrite the other. Designed to be called by external script. Right now need
' to use AFTER git checkout and git pull, but might be able to incorporate later.

' PARAMS
' WorkingDir[String]: script must pass working directory

' RETURNS
' True = successful
' False = unsuccessful

Public Function VbaCheckout(WorkingDir As String) As Boolean
  Dim objCheckoutRepo As Repository
  Set objCheckoutRepo = Factory.CreateRepository(WorkingDir)
  
' Checkout means we changed file in repo, need to copy TO local
  objCheckoutRepo.CopyRepoDocToLocal

' maybe useful later
  VbaCheckout = True
  
End Function


' ===== VbaMerge ==============================================================
' Call from a script that passes working directory. Right now you still need to
' run git merge first, then run this macro after.

' PARAMS
' WorkingDir[String]: current working directory (should be repo)

' RETURNS
' True = successful
' False = not

' TODO
' incorporate actual git merge command

Public Function VbaMerge(WorkingDir As String) As Boolean
  Dim objMergeRepo As Repository
  Set objMergeRepo = Factory.CreateRepository(WorkingDir)

' Merge might change files in repo, so we'll reimport them into the docs
  objMergeRepo.UpdateCodeInDoc
  
  VbaMerge = True
  
End Function



