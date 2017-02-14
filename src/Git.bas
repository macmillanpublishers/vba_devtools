Attribute VB_Name = "Git"
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


' ===== VbaSync ===============================================================
' Opens LOCAL copy of the Word template file for this repo, and opens VB editor.

' ASSUMPTIONS
' You're working on code from file in local directory

' PARAMS
' WorkingDir[String]: full Windows path to working dir, no trailing separator

' RETURNS: String
' Error number and message if any. Errors roll up the stack until they hit an
' On Error statement, so unhandled errors *should* end up here.

' TODO
' Add a "help" command that returns a string with each available command and
' what it does. Eventually could even store commands in a JSON that we read
' into a dictionary to loop through. Oh, could also use Application.Run to
' run a macro from a string, and then we don't even have to edit this code
' to add new commands, as long as all the basic things we want to do are
' publically available functions.

Public Function VbaSync(WorkingDir As String, Cmd As String) As String
   On Error GoTo Sync_Error

' Create Repository object for WordingDir
  Dim objRepo As Repository
  Set objRepo = Factory.CreateRepository(Path:=WorkingDir)

' Run macros based on the Cmd sent to the script
  Select Case Cmd
    Case "status"
      objRepo.UpdateCodeInRepo
      
    Case "checkout"
      objRepo.CopyRepoDocToLocal
  
    Case "merge"
      objRepo.UpdateCodeInDoc
      
    Case Else
      VbaSync = Cmd & " is not available in vba_devtools, please try again"
      Exit Function
  End Select

Sync_Error:
  VbaSync = Vba_Error
End Function


' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'     PRIVATE PROCEDURES
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ===== Vba_Error =============================================================
' Parse error object to return error number to calling script.

' PARAMS
' NONE: Only one Err object at a time so will just access current error properties

' RETURNS: String
' If error, returns number and description
' If no error, returns success message

' TODO
' Might have to add Err.Clear for handled errors, not sure if it persists

Private Function Vba_Error() As String
  If Err.Number = 0 Then
    Vba_Error = "SUCCESS: macro completed without error"
  Else
    Vba_Error = "VBA error " & Err.Number & ": " & Err.Description
  End If
End Function
