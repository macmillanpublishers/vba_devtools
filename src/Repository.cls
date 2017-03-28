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
Private prop_LocalDocFullName As String
Private prop_SrcPath As String
Private prop_SubmoduleExists As Boolean
Private prop_SubmodulePaths As Collection
Private prop_SubmoduleFiles As Dictionary
Private prop_ExtFromModuleType As Dictionary
Private prop_ModuleTypeFromExt As Dictionary
Private prop_AllModuleFiles As Dictionary

' JUST VARIABLE TO STORE, NOT A PROPERTY
Private m_blnRepoInit As Boolean

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'     EVENTS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ===== Initialize ============================================================
' Sets the g_blnRepoInit variable False so we know it's OK to run Init later.

Private Sub Class_Initialize()
  m_blnRepoInit = False
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
    prop_RepoName = Right(RepoPath, Len(RepoPath) - InStrRev(RepoPath, _
      Application.PathSeparator))
  End If
  RepoName = prop_RepoName
End Property

' ===== ConfigName ============================================================
' File name of config file in repo.

Private Property Get ConfigName() As String
  If prop_ConfigName = vbNullString Then
    prop_ConfigName = RepoName & "_config.json"
  End If
  ConfigName = prop_ConfigName
End Property

' ===== ConfigFullName ========================================================
' Full path (including file name and extension) of config file in repo.

Private Property Get ConfigFullName() As String
  If prop_ConfigFullName = vbNullString Then
    prop_ConfigFullName = RepoPath & Application.PathSeparator & ConfigName
  End If
  ConfigFullName = prop_ConfigFullName
End Property

' ===== ConfigData ============================================================
' Dictionary object created from config JSON in repo.

' TODO: probably need a class just for this at some point.
Private Property Get ConfigData() As Dictionary
  If prop_ConfigData Is Nothing Then
    Set prop_ConfigData = Utils.ReadJson(JsonPath:=ConfigFullName)
  End If
  Set ConfigData = prop_ConfigData
End Property

' ===== MacroDocName ==========================================================
' File name of the macro-enabled template/doc file. Read only, so file name must
' match repo name. Extension stored in config file.

Private Property Get MacroDocName() As String
  If prop_MacroDocName = vbNullString Then
    prop_MacroDocName = RepoName & "." & ConfigData("extension")
  End If
  MacroDocName = prop_MacroDocName
End Property

' ===== RepoDocFullName =======================================================
' Path to the template file in the repo, including file name and ext. Read only.

Private Property Get RepoDocFullName() As String
  If prop_RepoDocFullName = vbNullString Then
    prop_RepoDocFullName = RepoPath & Application.PathSeparator & MacroDocName
  End If
  RepoDocFullName = prop_RepoDocFullName
End Property

' ===== InstallType ===========================================================
' Returns the InstallType from the vba_devtools config file. Requires WordStartup
' is saved as an environment variable.

' TODO: make more generic for all files to read config.

Private Property Get InstallType() As String
  If prop_InstallType = vbNullString Then
  ' Build path to devtools config file (always in startup)
    Dim strDevConfigFullName As String
    strDevConfigFullName = Environ("WordStartup") & Application.PathSeparator _
      & "vba_devtools_config.json"
      
  ' Read JSON config into Dictionary
    Dim dictDevConfig As Dictionary
    Set dictDevConfig = Utils.ReadJson(JsonPath:=strDevConfigFullName)
    prop_InstallType = dictDevConfig("localInstall")("current")
  End If
  InstallType = prop_InstallType
End Property

' ===== LocalDocPath ==========================================================
' Path to local doc w/o file name. NOTE: assumes you've set the WordStartup env.

Private Property Get LocalDocPath() As String
  If prop_LocalDocPath = vbNullString Then
  ' Access "location" value (which should = an env. variable)
    Dim strLocationEnv As String
    strLocationEnv = ConfigData("localInstall")(InstallType)("environmentVar")
  
  ' Get full path to that dir from env variables
    prop_LocalDocPath = Environ(strLocationEnv)
    
    If prop_LocalDocPath = vbNullString Then
      Stop
    End If
  End If
  LocalDocPath = prop_LocalDocPath
End Property

' ===== LocalDocFullName ======================================================
' Path & file name to local macro template/doc/whatever, per repo Config file.

Private Property Get LocalDocFullName() As String
  If prop_LocalDocFullName = vbNullString Then
    prop_LocalDocFullName = LocalDocPath & Application.PathSeparator & _
      MacroDocName
  End If
  LocalDocFullName = prop_LocalDocFullName
End Property

' ===== SrcPath ===============================================================
' Full path to directory VBA modules are exported to. Must be "src". Read only

Private Property Get SrcPath() As String
  If prop_SrcPath = vbNullString Then
    prop_SrcPath = RepoPath & Application.PathSeparator & "src"
  End If
  SrcPath = prop_SrcPath
End Property

' ===== SubmoduleExists =======================================================
' Does repo contain git submodules. No global prop_SubmoduleExists because I
' can't figure how to test if we've set it yet (Boolean defaults to False).

Private Property Get SubmoduleExists() As Boolean
  SubmoduleExists = Utils.IsItThere(RepoPath & Application.PathSeparator _
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
    
    If SubmoduleExists = True Then
    ' Read .gitmodules file
      strGitmodulesText = Utils.ReadTextFile(RepoPath & _
        Application.PathSeparator & ".gitmodules", False)
  '    Debug.Print strGitmodulesText
      
    ' Pull relative paths out with regex (Matches full "path" lines)
      Set objRegEx = New RegExp
      With objRegEx
        .Global = True  ' True = find all matches
        .MultiLine = True  ' True = start/end matches line by line
        .Pattern = "^\tpath = (.*)[\r\n]+$"
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
        strPath = RepoPath & Application.PathSeparator & _
          objMatch.SubMatches.Item(0)
'        Debug.Print "|" & strPath & "|"
        strPath = Trim(strPath)
        prop_SubmodulePaths.Add strPath
      Next objMatch
    End If
  End If
  Set SubmodulePaths = prop_SubmodulePaths
End Property

' ===== SubmoduleFiles ========================================================
' Dictionary of VBA module names that are currently saved in git submodules.
' Dictionary key is file name w/o file ext (should match module name), and value
' is the full path to that file, including file name & ext.

Private Property Get SubmoduleFiles() As Dictionary
  If prop_SubmoduleFiles Is Nothing Then
    Set prop_SubmoduleFiles = New Dictionary
    
    If SubmoduleExists = True Then
    ' .SubmodulePaths is a collection, need Variant to do For Each
      Dim varPath As Variant
      Dim dictFromOneDir As Dictionary
      For Each varPath In SubmodulePaths
      ' Set = New to clear dict from previous loop
        Set dictFromOneDir = New Dictionary
        Set dictFromOneDir = GetModulesInDir(CStr(varPath))
        Set prop_SubmoduleFiles = Utils.MergeDictionary(prop_SubmoduleFiles, _
          dictFromOneDir)
      Next varPath
    End If
  End If
  Set SubmoduleFiles = prop_SubmoduleFiles
End Property


' ===== AllModuleFiles ========================================================
' Dictionary in which keys are file names WITHOUT extension (ie module name),
' values are full path to that file. SAME STRUCTURE AS SubmoduleFiles, so we're
' really just adding the files in Src to that Dictionary.

Private Property Get AllModuleFiles() As Dictionary
  If prop_AllModuleFiles Is Nothing Then
  ' It's just Submodule files + files in Src, so add together
    Set prop_AllModuleFiles = GetModulesInDir(SrcPath)
    Set prop_AllModuleFiles = Utils.MergeDictionary(prop_AllModuleFiles, _
      SubmoduleFiles)
  End If
  
  Set AllModuleFiles = prop_AllModuleFiles
End Property

    
' ===== ExtFromModuleType =====================================================
' VBComponent.Name just returns name in code window, not file name. Have to
' generate file extension based on the .Type property to save it. This builds
' a Dictionary object to grab that from.

' Should be able to just use as a Dictionary: strExt = .ExtFromModuleType(VbComp.Type)

' RETURNS
' Dictionary with keys = vbext_ComponentType and values = file ext w/o dot

Private Property Get ExtFromModuleType() As Dictionary
  If prop_ExtFromModuleType Is Nothing Then
    Set prop_ExtFromModuleType = New Dictionary
    With prop_ExtFromModuleType
      .Add vbext_ct_ClassModule, "cls"
      .Add vbext_ct_Document, "cls"
      .Add vbext_ct_StdModule, "bas"
' Don't export userforms, .frx gets tracked every time
'      .Add vbext_ct_MSForm, "frm"
    End With
  End If
  Set ExtFromModuleType = prop_ExtFromModuleType
End Property

' ===== ModuleTypeFromExt =======================================================
' Give it a valid VBA module file extension (w/o dot) and it returns the Type.

' RETURNS
' Dictionary with keys = file ext w/o dot and values = vbext_ComponentType

Private Property Get ModuleTypeFromExt() As Dictionary
  If prop_ModuleTypeFromExt Is Nothing Then
    Set prop_ModuleTypeFromExt = New Dictionary
    With prop_ModuleTypeFromExt
      .Add "cls", vbext_ct_ClassModule
      .Add "doccls", vbext_ct_Document
      .Add "bas", vbext_ct_StdModule
' Removing userforms, because .frx binary file is tracked as modified every time
'      .Add "frm", vbext_ct_MSForm
    End With
  End If
  Set ModuleTypeFromExt = prop_ModuleTypeFromExt
End Property




' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'       PUBLIC METHODS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ===== Init ==================================================================
' Constructor for the class. Called from Factory.CreateRepository. Will create
' new object and also set default properties based on arguments. Will only run
' once per object because of blnRepoInit variable.

Public Sub Init(RepoDir As String)
  If m_blnRepoInit = False Then
    m_blnRepoInit = True
    RepoPath = RepoDir
  End If
End Sub

' ===== UpdateCodeInRepo ======================================================
' Update files in repo with code from local macro doc, also copy macro doc to repo
' Will leave local macro doc in original open/closed state.

Public Sub UpdateCodeInRepo()
' Make sure local doc is open
  Dim blnUpdateDocOpen As Boolean
  blnUpdateDocOpen = Utils.DocOpenSave(Path:=LocalDocFullName)

' Export code modules from local doc to repo
  ExportCodeFiles

' If doc was closed before, close it now
' (because CopyLocalToRepo returns file to original state)
  If blnUpdateDocOpen = False Then
    Documents(LocalDocFullName).Close SaveChanges:=wdSaveChanges
  End If
  
' Copy doc file from local to repo
  CopyLocalDocToRepo
End Sub

' ===== UpdateCodeInDoc =====================================================
' Replace code in Macro document in REPO with code files in repo. Then copy the
' updated repo macro doc to the local directory.

Public Sub UpdateCodeInDoc()
' Make sure repo doc is open (and record state so we can return later)
  Dim blnRepoDocOpen As Boolean
  blnRepoDocOpen = Utils.DocOpenSave(Path:=RepoDocFullName)

' Import code from repo to doc
  ImportCodeToDoc
  
' Return file to original save state
  If blnRepoDocOpen = False Then
    Documents(RepoDocFullName).Close SaveChanges:=wdSaveChanges
  End If

' Copy doc from repo to local dir.
  CopyRepoDocToLocal

End Sub


' ===== CopyLocalToRepo =======================================================
' Copies the local macro doc/template to the repo. Calls DocFileCopy, which
' validates the file and closes before copying, and re-opens if was open before

Public Sub CopyLocalDocToRepo()
  Utils.DocFileCopy LocalDocFullName, RepoDocFullName
End Sub


' ===== CopyRepoToLocal =======================================================
' Same as above but reversed.

Public Sub CopyRepoDocToLocal()
  Utils.DocFileCopy RepoDocFullName, LocalDocFullName
End Sub



' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'       PRIVATE METHODS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ===== GetModulesInDir =======================================================
' We've done this a couple of times, so we'll make it a more general thing.
' Pass in a directory, and it will create a dictionary of files that are VBA
' modules.

' ASSUMPTIONS
' Only checks main directory, no subdirectoires (though we could certainly do
' that at some point if we need to)

' PARAMS
' TargetDir[String]: The directory we're looking in, no training slash

' RETURNS
' Dictionary with keys = file name without extension, value = full path to file

Private Function GetModulesInDir(TargetDir As String) As Dictionary
  Set GetModulesInDir = New Dictionary

  Dim strFileAndExt As String
  Dim strFileNameOnly As String
  Dim lngCounter As Long
  Dim strExt As String

  strFileAndExt = Dir(TargetDir & Application.PathSeparator)
  Do Until strFileAndExt = vbNullString Or lngCounter = 50
  ' To prevent infinite loop
    lngCounter = lngCounter + 1
  
  ' Get file extension to verify it's a valid VBA code module
    strExt = Utils.GetFileExtension(strFileAndExt)
    
  ' Valid code file extensions are in ModuleTypeFromExt dictionary
    If ModuleTypeFromExt.Exists(strExt) = True Then
    ' Module.Name has no extension, so save that as key
      strFileNameOnly = Utils.GetFileNameOnly(strFileAndExt)
      
    ' Use .Item() = value not .Add , even though I don't think you can have two
    ' different types of modules w/ same name.
      GetModulesInDir.Item(strFileNameOnly) = TargetDir & Application.PathSeparator _
        & strFileAndExt
    End If
  
  ' Get next file in dir for loop
    strFileAndExt = Dir()
  Loop
End Function


' ===== ClearThisDocument =====================================================
' The ThisDocument class can't be deleted, so we have to delete current code
' line by line.

' ASSUMPTIONS
' Target document is already open.

' PARAMS
' ThisDocModule[VBComponent]: The ThisDocument class component

Private Sub ClearThisDocument(ThisDocModule As VBIDE.VBComponent)
  Dim colComponents As VBIDE.VBComponents
  Dim objTempModule As VBIDE.VBComponent

' We access the Components collection a few times, so set a variable here
  Set colComponents = ThisDocModule.Collection
  
' Code for the REAL ThisDocument class
  With ThisDocModule.CodeModule
  ' Delete current code line by line, if any
    If .CountOfLines > 0 Then
      .DeleteLines 1, .CountOfLines
    End If
  End With
  
End Sub

' ===== ImportThisDocument =====================================================
' The ThisDocument class can't be imported b/c there is already a class with that
' name that can't be deleted, so we'll add new code line by line.

' ASSUMPTIONS
' Target document is already open.
' We know the file Repo/src/ThisDocument.cls already exists.
' ThisDocument module has already been cleared of code

' PARAMS
' ThisDocModule[VBComponent]: The ThisDocument module

Private Sub ImportThisDocument(ModulesCollection As VBIDE.VBComponents)
' We access the Components collection a few times, so set a variable here
  Dim objThisDocModule As VBIDE.VBComponent
  Set objThisDocModule = ModulesCollection("ThisDocument")
  
' Exported text doc has attributes that do not get printed to the code module
' so we can't just read the text file and paste it in. Instead we'll import
' the module in a TEMP module...
  Dim objTempModule As VBIDE.VBComponent
  Set objTempModule = ModulesCollection.Import(SrcPath & Application.PathSeparator _
    & "ThisDocument.cls")
  
' ... copy code from temp module
  Dim strNewCode As String
  strNewCode = objTempModule.CodeModule.lines(1, objTempModule.CodeModule.CountOfLines)
 
' Insert copied code into original ThisDocument code
' If there is already code, append
  With objThisDocModule.CodeModule
    .InsertLines Line:=.CountOfLines + 1, String:=strNewCode
  End With

' And cleanup! delete temp module
  ModulesCollection.Remove objTempModule
  
End Sub


' ===== ImportCodeToDoc =======================================================
' Imports all VBA components in repo to Macro document in repo. If the module
' is already in the doc, it is overwritten by the version in the repo. If a new
' module is in the repo, it is added to doc. If a module is in the doc but not
' in the repo, it stays in the doc.

' ASSUMPTIONS
' Repo doc file has already been validated and opened.

' TODO (maybe)
' Maybe if module is in doc but not in repo, we export it to repo?
' OR do we rebuild the doc file from scratch and delete the old one?

Private Sub ImportCodeToDoc()
' Collection of all modules for the doc in the repo
  Dim colVbComponents As VBIDE.VBComponents
  Set colVbComponents = Documents(RepoDocFullName).VBProject.VBComponents

' We can't import a code file if there is already a module with the same name
' in the document. So we'll loop through the modules already in the doc and
' remove any that match code files that we want to import. Doing delete and
' import and two separate steps because we want to keep modules that aren't
' in repo.

' Loop through all modules in project, delete any that are already in repo
  Dim objVbModule As VBIDE.VBComponent
  For Each objVbModule In colVbComponents
  ' We need to treat ThisDocument class separately b/c can't delete
    If objVbModule.Name = "ThisDocument" Then
      ClearThisDocument ThisDocModule:=objVbModule
    Else
      colVbComponents.Remove objVbModule
    End If
  Next objVbModule
  
' Now we can import files!
  Dim varFileName As Variant
  For Each varFileName In AllModuleFiles.Keys
    If CStr(varFileName) = "ThisDocument" Then
      
      ImportThisDocument colVbComponents
    Else
      colVbComponents.Import AllModuleFiles(CStr(varFileName))
    End If
  Next varFileName
  
' Save!
  Documents(RepoDocFullName).Save

End Sub


' ===== ExportCodeFiles =======================================================
' Exports VBA components (i.e., .bas and .cls files) from the local doc/template
' to this repo. Handles all the submodule stuff. Repo must be set up as described
' at top.

' ASSUMPTIONS:
' Local doc file has already been validated and opened.
' Does NOT export userforms (.frm) b/c the .frx binary file is always tracked as
' modified even when nothing has changed. Will have to export those manually for now

Private Sub ExportCodeFiles()
  Dim objVbProj As VBIDE.VBProject
  Dim objVbModule As VBIDE.VBComponent
  Dim strModuleRepoPath As String
  Dim strFileExt As String
  
  Set objVbProj = Documents(LocalDocFullName).VBProject
  
' Loop through VBA modules in the template project
  For Each objVbModule In objVbProj.VBComponents
    If SubmoduleFiles.Exists(objVbModule.Name) = True Then
      strModuleRepoPath = SubmoduleFiles(objVbModule.Name)
    Else
      strFileExt = ExtFromModuleType(objVbModule.Type)
      If strFileExt <> vbNullString Then
        strModuleRepoPath = SrcPath & Application.PathSeparator & objVbModule.Name _
          & "." & strFileExt
      Else
        strModuleRepoPath = vbNullString
      End If
    End If
    
  ' If module is a userform we've set save path to null, so don't try to export
    If strModuleRepoPath <> vbNullString Then
      ExportVbModule VBComp:=objVbModule, FullPath:=strModuleRepoPath
    End If
  Next objVbModule
End Sub

' ===== ExportVbModule ========================================================
' Export VBA module passed to the FILE passed to it.

Private Sub ExportVbModule(VBComp As VBIDE.VBComponent, FullPath As String)

' ThisDocument module always there by default but doesn't always have code
  If VBComp.Name <> "ThisDocument" Or VBComp.CodeModule.CountOfLines <> 0 Then
  ' Delete file in repo if it exists
    If Utils.IsItThere(Path:=FullPath) = True Then
      Utils.KillAll Path:=FullPath
    End If
  ' Export the module
    VBComp.Export FileName:=FullPath
  End If
    
End Sub
