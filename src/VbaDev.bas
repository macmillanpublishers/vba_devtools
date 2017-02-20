Attribute VB_Name = "VbaDev"
Option Explicit

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'     PUBLIC PROCEDURES
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ===== RunValidator ==========================================================
' Run Validator.Launch macro solo, without whole validator toolchain. Copies
' test file to a tmp directory (in same dir as orig file) and runs macro on that
' so we don't override the orig file. Also gets log file set up.

' ASSUMPTIONS
' You have the bookmaker.dotm file saved in STARTUP.
' This is called from powershell run_dev_macro script.

' PARAMS
' ArrayParam: Pass as an array so we can use same powershell. Array only needs
' 1 element, the full windows path to original document file.

' RETURNS
' Error messages

' TODO
' Move some of the tmp file stuff to Word-template and implement for users.
Sub testit()
  Dim a_testfile(0 To 0) As String
  a_testfile(0) = ActiveDocument.FullName
  Dim stroutput As String
  stroutput = RunValidator(a_testfile)
  Debug.Print stroutput
End Sub

Public Function RunValidator(ArrayParam As Variant) As String
  On Error GoTo RunValidatorError
' Get parameter into variable
  Dim strOrigFullName As String
  strOrigFullName = ArrayParam(0)
  
' Save and close, if it was open, record origin state
  Dim blnOrigWasOpen As Boolean
  blnOrigWasOpen = Utils.DocSaveClose(strOrigFullName)
  
' Create paths to tmp dir, log file, etc.
  Dim strOrigPath As String
  Dim strNormName As String
  Dim strNormNameOnly As String
  Dim strTmpNameTail As String
  
  strOrigPath = Utils.GetPath(strOrigFullName)
  strNormName = NormalizeFileName(Utils.GetFileName(strOrigFullName))
  strNormNameOnly = Utils.GetFileNameOnly(strNormName)
  strTmpNameTail = "_tmp_" & strNormName
  
  Dim strTmpPath As String
  Dim strLogFullName As String

  strTmpPath = strOrigPath & Application.PathSeparator & "MacroTmp_" & strNormNameOnly
  strLogFullName = strTmpPath & Application.PathSeparator & "LOG_" & strNormNameOnly & ".txt"

  Dim strTmpName As String
  Dim strTmpFullName As String

  If Utils.IsItThere(strTmpPath) = False Then
    MkDir strTmpPath
    strTmpName = "00" & strTmpNameTail
  Else
    Dim A As Long
    For A = 0 To 99
      strTmpName = VBA.Format(A, "00") & strTmpNameTail
      If Utils.IsItThere(strTmpName) = False Then
        Exit For
      End If
    Next A
  End If
  
  strTmpFullName = strTmpPath & Application.PathSeparator & strTmpName
  FileCopy strOrigFullName, strTmpFullName
  
' Validator needs a file called book_info.json in same dir as file.
  Call CheckBookInfo(strOrigPath, strTmpPath)
  
  Call Validator.Launch(strTmpFullName, strLogFullName)
  
  If blnOrigWasOpen = True Then
    Documents.Open (strOrigFullName)
  End If

RunValidatorError:
  If Err.Number <> 0 Then
    RunValidator = Err.Number & ": " & Err.Description
  Else
    RunValidator = "SUCCESS: macro run without errors"
  End If
End Function


' ===== NormalizeFileName =====================================================
' Remove everything other than letters, numbers, underscores and hyphens from
' file name. Keeps dot separating name and extension, but remove any other dots

' NOT in Utils.bas because it requires RegEx lib to compile, and we can't be
' sure other projects will have that (and Mac doesn't have it available anyway)

' ASSUMPTIONS
' Have Regular Expressions library referenced in project

' PARAMS
' FileName: String file name, not including path

' RETURNS
' String file name with that stuff removed.

' TODO
' Figure out a way to do it w/o RegEx so can use in Utils/Mac


Private Function NormalizeFileName(OrigFileName As String) As String
  Dim strOrigExt As String
  strOrigExt = Utils.GetFileExtension(File:=OrigFileName)
  
' Prefix extension with dot here, because if it didn't have one,
' we don't want to add it again later
  If strOrigExt <> vbNullString Then
    strOrigExt = "." & strOrigExt
  End If
  
  Dim strNormName As String
  strNormName = Utils.GetFileNameOnly(File:=OrigFileName)
  
  Dim objNormalizeRegEx As RegExp
  Set objNormalizeRegEx = New RegExp
  
  objNormalizeRegEx.Global = True
  objNormalizeRegEx.Pattern = "[^a-zA-Z0-9_-]"
  
  strNormName = objNormalizeRegEx.Replace(strNormName, vbNullString)
  NormalizeFileName = strNormName & strOrigExt
End Function


' ===== CheckBookInfo =========================================================
' Validator macro needs a file named book_info.json in same dir as file. This
' checks if one is in the original location and, if it is, copies it to the
' tmp dir. If there isn't one, it creates a file with dummy data.

' PARAMS
' SourceDir: Folder to check for original file
' TmpDir: Folder to copy original file to

' TODO
' Put together some kind of standard dummy data, commit to repo, copy to
' Startup, and just copy that file? discuss w/ everyone to see if makes sense

Private Sub CheckBookInfo(SourceDir As String, TmpDir As String)
  Dim strInfoFileName As String
  Dim strOrigInfoFullName As String
  Dim strTmpInfoFullName As String
  
  strInfoFileName = "book_info.json"
  strOrigInfoFullName = SourceDir & Application.PathSeparator & strInfoFileName
  strTmpInfoFullName = TmpDir & Application.PathSeparator & strInfoFileName
  
  If Utils.IsItThere(strOrigInfoFullName) = True Then
    FileCopy strOrigInfoFullName, strTmpInfoFullName
  Else
    Dim dictInfoJson As Dictionary
    Set dictInfoJson = New Dictionary
    
    dictInfoJson.Add "production_editor", "Eric Meyer"
    dictInfoJson.Add "production_manager", "Eric Gladstone"
    dictInfoJson.Add "work_id", "86877"
    dictInfoJson.Add "isbn", "9781250087058"
    dictInfoJson.Add "title", "The Netanyahu Years"
    dictInfoJson.Add "author", "Ben Caspit translated by Ora Cummings"
    dictInfoJson.Add "product_type", "Book"
    dictInfoJson.Add "imprint", "Thomas Dunne Books"
    
    Utils.WriteJson strTmpInfoFullName, dictInfoJson
  End If
End Sub
