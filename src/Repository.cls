VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Repository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----- USE -------------------------------------------------------------------
' Creates Repository object based on directory in .RepoDir property.
' Easier to access info about the repo and do various dev things

' ----- DEPENDENCIES ----------------------------------------------------------
' Reference to Microsoft VBScript Regular Expressions 5.5 (Windows only!)
' vba_devtools setup stuff:
' * Set WordStartup env variable
' * Copy vba_devtools.dotm and vba_devtools_config.json to Startup
' * Set "installType" key in local vba_devtools_config.json

' Imported (from git submodules):
' vba_utilities/Utils.bas
' VBA-JSON/JsonConverter.bas
' VBA-Dictionary/Dictionary.cls

' Assumes the following repo organization:
' * 1 VBA project/template per repo
' * .dotm template file saved in root of repo
' * template file and VBA project match repo name
' * VBA modules (.bas, .cls, .frm) exported to /src dir in repo
' * .dotm is run from other local dir (b/c can't checkin/checkout/merge
'   an open file, so don't run from repo directly)
' * local .dotm path stored in config/config.json
' * VBA modules not tracked in this repo added as git submodules
' * ALL .bas and .cls files in submodule repos are imported
' * New submodule VBA modules must be exported manually first time

' *****************************************************************************
'           CLASS VARIABLES
' *****************************************************************************
' To store property variables between calls. Only reference in procedures for
' that same property. All other code should reference property directly.

Private prop_RepoPath As String
Private prop_RepoName As String
Private prop_ConfigName As String
Private prop_ConfigFullName As String
Private prop_ConfigData As Dictionary
Private prop_MacroDocName As String
Private prop_RepoDocFullName As String
Private prop_InstallType As String
Private prop_LocalDocPath As String
Private prop_LocalDocFullPath As String
Private prop_SrcPath As String
Private prop_SubmoduleExists As Boolean

Private prop_SubmodulePaths As Collection

Private prop_SubmoduleFiles As Dictionary

' JUST VARIABLE TO STORE, NOT A PROPERTY
Private g_blnRepoInit As Boolean

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'     EVENTS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ===== Initialize ============================================================
' Sets the g_blnRepoInit variable False so we know it's OK to run Init later.

Private Sub Class_Initialize()
  g_blnRepoInit = False
End Sub


' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'     PUBILC PROPERTIES
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'       PRIVATE PROPERTIES
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ===== RepoPath ==============================================================
' Full path to local repo. Only Let from Init Method!

Private Property Let RepoPath(v_strRepoPath As String)
  prop_RepoPath = v_strRepoPath
End Property

Private Property Get RepoPath() As String
  RepoPath = prop_RepoPath
End Property

' ===== RepoName ==============================================================
' Name of repo. Read only because pulled from RepoPath property.

Private Property Get RepoName() As String
  If prop_RepoName = vbNullString Then
    prop_RepoName = Right(Me.RepoPath, Len(Me.RepoPath) - InStrRev(Me.RepoPath, _
      Application.PathSeparator))
  End If
  RepoName = prop_RepoName
End Property

' ===== ConfigName ============================================================
' File name of config file in repo.

Private Property Get ConfigName() As String
  If prop_ConfigName = vbNullString Then
    prop_ConfigName = Me.RepoName & "_config.json"
  End If
  ConfigName = prop_ConfigName
End Property

' ===== ConfigFullName ========================================================
' Full path (including file name and extension) of config file in repo.

Private Property Get ConfigFullName() As String
  If prop_ConfigFullName = vbNullString Then
    prop_ConfigFullName = Me.RepoPath & Application.PathSeparator & Me.ConfigName
  End If
  ConfigFullName = prop_ConfigFullName
End Property

' ===== ConfigData ============================================================
' Dictionary object created from config JSON in repo.

' TODO: probably need a class just for this at some point.
Private Property Get ConfigData() As Dictionary
  If prop_ConfigData Is Nothing Then
    Set prop_ConfigData = Utils.ReadJson(JsonPath:=Me.ConfigFullPath)
  End If
  Set ConfigData = prop_ConfigData
End Property

' ===== MacroDocName ==========================================================
' File name of the macro-enabled template/doc file. Read only, so file name must
' match repo name. Extension stored in config file.

Private Property Get MacroDocName() As String
  If prop_MacroDocName = vbNullString Then
    prop_MacroDocName = Me.RepoName & "." & Me.ConfigData("extension")
  End If
  MacroDocName = prop_MacroDocName
End Property

' ===== RepoDocFullName =======================================================
' Path to the template file in the repo, including file name and ext. Read only.

Private Property Get RepoDocFullName() As String
  If prop_RepoDocFullName = vbNullString Then
    prop_RepoDocPath = Me.RepoPath & Application.PathSeparator & Me.MacroDocName
  End If
  RepoDocFullName = prop_RepoDocPath
End Property

' ===== InstallType ===========================================================
' Returns the InstallType from the vba_devtools config file. Requires WordStartup
' is saved as an environment variable.

' TODO: create enum for return options.

Private Property Get InstallType() As String
  If prop_InstallType = vbNullString Then
  ' Build path to devtools config file (always in startup)
    Dim strDevConfigFullName As String
    strDevConfigFullName = Environ("WordStartup") & Application.PathSeparator _
      & "vba_devtools_config.json"
      
  ' Read JSON config into Dictionary
    Dim dictDevConfig As Dictionary
    Set dictDevConfig = Utils.ReadJson(JsonPath:=strDevConfigFullName)
    prop_InstallType = dictDevConfig("installType")
  End If
  InstallType = prop_InstallType
End Property

' ===== LocalDocPath ==========================================================
' Path to local doc w/o file name. NOTE: assumes you've set the WordStartup env.

Private Property Get LocalDocPath() As String
  If prop_LocalDocPath = vbNullString Then
  ' Access "location" value (which should = an env. variable)
    Dim strLocationEnv As String
    strLocationEnv = Me.ConfigData("location")(Me.InstallType)
  
  ' Get full path to that dir from env variables
    prop_LocalDocPath = Environ(strLocationEnv)
  End If
  LocalDocPath = prop_LocalDocPath
End Property

' ===== LocalDocFullPath ======================================================
' Path & file name to local macro template/doc/whatever, per repo Config file.

Private Property Get LocalDocFullPath() As String
  If prop_LocalDocFullPath = vbNullString Then
    prop_LocalDocFullPath = Me.LocalDocPath & Application.PathSeparator & _
      Me.MacroDocName
  End If
  LocalDocFullPath = prop_LocalDocFullPath
End Property

' ===== SrcPath ===============================================================
' Full path to directory VBA modules are exported to. Must be "src". Read only

Private Property Get SrcPath() As String
  If prop_SrcPath = vbNullString Then
    prop_SrcPath = Me.RepoPath & Application.Path & "src"
  End If
  ScrPath = prop_SrcPath
End Property

' ===== SubmoduleExists =======================================================
' Does repo contain git submodules. No global prop_SubmoduleExists because I
' can't figure how to test if we've set it yet (Boolean defaults to False).

Private Property Get SubmoduleExists() As Boolean
  SubmoduleExists = Utils.IsItThere(Me.RepoPath & Application.PathSeparator _
    & ".gitmodules")
End Property

' ===== SubmodulePaths ========================================================
' Collection of paths to submodule dirs, if any.

Private Property Get SubmodulePaths() As Collection
  If prop_SubmodulePaths Is Nothing Then
    Dim strGitmodulesText As String
    Dim objRegEx As RegExp
    Dim c_Matches As MatchCollection
    Dim c_strSubmodulePaths As Collection
    
    If Me.SubmoduleExists = True Then
    ' Read .gitmodules file
      strGitmodulesText = Utils.ReadTextFile(Me.RepoPath & _
        Application.PathSeparator & ".gitmodules", False)
  '    Debug.Print strGitmodulesText
      
    ' Pull relative paths out with regex (Matches full "path" lines)
      Set objRegEx = New RegExp
      With objRegEx
        .Global = True  ' True = find all matches
        .MultiLine = True  ' True = start/end matches line by line
        .Pattern = "^\tpath = (.*)$"
        Set c_Matches = .Execute(strGitmodulesText)
      End With
    End If
  
  ' Parse regex Match objects to pull out just returned values (parens in regex)
  ' and build full path, then add those to a Collection to return
    Dim objMatch As Match
    Dim strPath As String
    Set prop_SubmodulePaths = New Collection
    
    If c_Matches.Count > 0 Then
      For Each objMatch In c_Matches
        strPath = Me.RepoPath & Application.PathSeparator & _
          objMatch.SubMatches.Item(0)
  '      Debug.Print strPath
        prop_SubmodulePaths.Add strPath
      Next objMatch
    End If
  End If
  Set SubmodulePaths = prop_SubmodulePaths
End Property

' ===== SubmoduleFiles ========================================================
' Dictionary of VBA module files that are currently in git submodules. Key is
' module name, value is the full path to that file, including file name & ext.

Private Property Get SubmoduleFiles() As Dictionary
  If prop_SubmoduleFiles Is Nothing Then
    Set prop_SubmoduleFiles = New Dictionary
    
    If Me.SubmoduleExists = True Then
    ' Loop through files in each git submodule dir, add VBA code files to dict.
    ' Just looking in root now, could do recursive later if needed.
      Dim varPath As Variant
      Dim strFile As String
      Dim lngCounter As Long
      For Each varPath In Me.SubmodulePaths
        strFile = Dir(varPath)
        Do Until strFile = vbNullString Or lngCounter = 50
          prop_SubmoduleFiles.Add strFile, varPath & Application.PathSeparator _
            & strFile
          strFile = Dir()
        Loop
      Next varPath
    Else
    ' What can we return if there aren't any files?
    End If

  End If
  Set SubomduleFiles = prop_SubmoduleFiles
End Property


' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'       PUBLIC METHODS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ===== Init ==================================================================
' Constructor for the class. Called from Factory.CreateRepository. Will create
' new object and also set default properties based on arguments. Will only run
' once per object because of g_blnRepoInit variable.

Public Sub Init(p_RepoDir As String)
  Me.RepoPath = p_RepoDir
End Sub

' ===== UpdateRepo ============================================================
' Update files in repo with code from local macro doc, also copy macro doc to repo
' Will leave local macro doc in original open/closed state.

Public Function UpdateRepo() As Boolean
' Make sure local doc is open
  Dim blnUpdateDocOpen As Boolean
  blnUpdateDocOpen = Me.DocOpenSave(Path:=Me.LocalDocFullName)

' Export code modules from local doc to repo
  Me.UpdateCodeFiles

' If doc was closed before, close it now.
  If blnUpdateDocOpen = False Then
    Documents(Me.LocalDocFullName).Close SaveChanges:=wdSaveChanges
  End If
  
' Copy doc file from local to repo
  Me.SyncDocs p_RepoIsDestination:=True
  
End Function

' ===== SyncDocs ==============================================================
' Macro doc files is in two places ("repo" and "local"), this copies one to
' overwrite the other. Will reopen files that were open to start.

' PARAMS
' RepoIsDestination[Boolean]: True = copy file from local -> repo, False = copy
'   file from repo -> local.

Public Sub SyncDocs(p_RepoIsDestination As Boolean)
  Dim strOriginFile As String
  Dim strDestinationFile As String
  
  If p_RepoIsDestination = True Then
    strOriginFile = Me.LocalDocFullName
    strDestinationFile = Me.RepoDocFullName
  Else
    strOriginFile = Me.RepoDocFullName
    strDestinationFile = Me.LocalDocFullName
  End If
  Me.DocFileCopy strOriginFile, strDestinationFile
End Sub


' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'       PRIVATE METHODS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ===== DocFileCopy ===========================================================
' Copies a doc from Origin to Destination. Makes sure files are closed before
' copying, and reopens docs if they were open to start. DestinationFile will
' be overwritten.

' PARAMS
' OriginFile[String]: file to copy from
' DestinationFile[String]: file to copy to

Private Sub DocFileCopy(OriginFile As String, DestinationFile As String)
  Dim blnOriginOpened As Boolean
  Dim blnDestinationOpened As Boolean
  blnOriginOpened = Me.DocSaveClose(OriginFile)
  blnDestinationOpened = Me.DocSaveClose(DestinationFile)
  
  VBA.FileCopy Source:=OriginFile, Destination:=DestinationFile
  
  If blnOriginOpened = True Then
    Documents.Open (OriginFile)
  End If
  
  If blnDestinationOpened = True Then
    Documents.Open (DestinationFile)
  End If
End Sub


' ===== DocSaveClose ==========================================================
' Makes sure document is saved and closed, including various validation things.

' PARAMS
' Path[String]: Full path to file w/ filename and ext

' RETURNS
' True = Doc was originally open
' False = Doc was originally closed

Private Function DocSaveClose(Path As String) As Boolean
  Dim blnFileOpen As Boolean
  blnFileOpen = Utils.IsOpen(Path)
  
  If blnFileOpen = True Then
    Documents(Path).Close SaveChanges:=wdSaveChanges
  End If
  
  DocSaveClose = blnFileOpen

End Function


' ===== DocOpenSave ===========================================================
' Makes sure document is opened and saved.

' PARAMS
' Path[String]: Full path to file with filename and ext

' RETURNS
' True = Doc was originally open
' False = Doc was originally closed

Private Function DocOpenSave(Path As String) As Boolean
  Dim blnFileOpen As Boolean
  blnFileOpen = Utils.IsOpen(Path)
  
  If blnFileOpen = False Then
    Documents.Open Path
  End If
  
  DocOpenSave = blnFileOpen

End Function










' ===== UpdateDocFile =========================================================
' Imports all VBA components in repo to Macro document in repo.

' ASSUMPTIONS
' Repo file has already been validated and opened.

' RETURNS
' True = success
' False = not

Private Function UpdateDocFile() As Boolean

End Function



' ===== UpdateCodeFiles ======================================================
' Exports all VBA components (i.e., .bas, .cls, and .frm files) from the local
' doc/template to this repo. Handles all the submodule stuff.
' Repo must be set up as described at top.

' ASSUMPTIONS:
' Local doc file has already been validated and opened.

' RETURNS: Boolean
' True = it worked
' False = it didn't work, bah

' TODO:
' Maybe have it return a more descriptive string we can return for user?

Private Function UpdateCodeFiles() As Boolean
    
    ' Cycle through each open document
    Dim oDoc As Document
    Dim strExtension As String
    Dim oProject As VBIDE.VBProject
    Dim oModule As VBIDE.VBComponent
    Dim strTemplateModules As String
    Dim strDependencies As String
    Dim strDepFiles As String
    Dim strEachFile As String
    Dim strRepoPath As String
    Dim openTemplates As Collection
    Set openTemplates = New Collection
    
    For Each oDoc In Documents
        Debug.Print oDoc.Name
        ' Separate the name and the extension of the document
        strExtension = Right(oDoc.Name, Len(oDoc.Name) - _
            (InStrRev(oDoc.Name, ".") - 1))
        
        ' We just want to work with .dotm and .docm (others can't have macros)
        If strExtension = ".dotm" Or strExtension = ".docm" Then
            ' later need to close > copy > open these files, but if we loop
            ' thru Documents collection, "open" will add file back try again.
            ' So create Collection to loop once:
            openTemplates.Add oDoc
            
            ' get FULL path to this template in its repo
            strRepoPath = GetRepoPath(oDoc)
    
            If oDoc.Name = "genUtils.dotm" Then
            ' Modules that need to be imported into templates but that we do
            ' not want to track. We don't want to export these, so let's get
            ' then into a string check against later.
                
                strDependencies = strRepoPath & Application.PathSeparator & _
                    "dependencies"
                Debug.Print strDependencies
                ' Dir() w/ arguments returns first file name that matches
                ' !!!! When switch to submodules, will need to change this to
                ' search subdirectories.
                strEachFile = Dir(strDependencies & Application.PathSeparator & _
                     "*.*", vbNormal)
                Do While Len(strEachFile) > 0
                    Debug.Print strEachFile
                    strDepFiles = strDepFiles & strEachFile & vbNewLine
'                    Debug.Print strDepFiles
                    ' Dir() again w/o arguments returns the NEXT file that matches orig arguments
                    ' if nothing else matches, returns empty string
                    strEachFile = Dir
                Loop
            Else
                strDepFiles = vbNullString
            End If

            ' Make sure we're referencing the correct project
            Set oProject = oDoc.VBProject
        
            strTemplateModules = strRepoPath & Application.PathSeparator
            
            ' Cycle through each module
            For Each oModule In oProject.VBComponents
                ' Skip modules in dependencies directory
                If InStr(strDepFiles, oModule.Name) = 0 Then
                    ' Don't export forms, they are always wonky. Will have to
                    ' manage manually
                    If oModule.Type <> vbext_ct_MSForm Then
                        Call ExportVBComponent(VBComp:=oModule, _
                            FolderName:=strTemplateModules)
                    End If
                End If
            Next
        End If
    Next oDoc

End Function




' ===== SaveToRepo ============================================================
' Save the component passed to this repo. Checks for files that should go into
' git submodules.


' ===== ClearThisDocument =====================================================
' The ThisDocument class can't be deleted (because it represents our actual doc
' object, so this function simple deletes the code that is in it.


' ===== BuildThisDocument =====================================================
' Can't delete ThisDocument class, thus can't import a saved one. So this adds
' code from a saved ThisDocument.cls file line by line. If you don't use Clear
' first, will append (and probably mess something up).
