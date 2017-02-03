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
' Creates Repository object based on directory passed to it at instantiation.
' Easier to access info about the repo.

' ----- DEPENDENCIES ----------------------------------------------------------
' Assumes the following repo organization:
' * 1 VBA project per repo
' * .dotm template file saved in root of repo
' * template file and VBA project match repo name
' * VBA modules (.bas, .cls, .frm) exported to /src dir in repo
' * .dotm saved locally (not in repo) to run (b/c can't checkin/checkout/merge
'   an open file)
' * local .dotm path stored in config
' * VBA modules not tracked in this repo added as git submodules
' * ALL .bas and .cls files in submodule repos are imported
' * New submodule VBA modules must be exported manually first time

' *****************************************************************************
'           CLASS VARIABLES
' *****************************************************************************
Private p_RepoName As String
Private p_RepoDir As String
Private p_TemplateFile As String
Private p_TemplatePath As String
Private p_SrcDir As String
Private p_SubmoduleExists As Boolean
Private p_SubmodulePaths As Collection
Private p_SubmoduleNames As Collection


' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'     PUBILC PROPERTIES
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ===== RepoName ==============================================================
' Name of repo. Read only because pulled from RepoDir.

Public Property Get RepoName() As String
  p_RepoName = Right(Me.RepoDir, Len(Me.RepoDir) - InStrRev(Me.RepoDir, _
    Application.PathSeparator))
  RepoName = p_RepoName
End Property


' ===== RepoDir ===============================================================
' Full path to local repo

Public Property Let RepoDir(v_strRepoDir As String)
  p_RepoDir = v_strRepoDir
End Property

Public Property Get RepoDir() As String
  RepoDir = p_RepoDir
End Property


' ===== TemplateFile ==========================================================
' Just the file name of the template file in the repo. Read only. Template must
' match repo name.

Public Property Get TemplateFile() As String
  p_TemplateFile = Me.RepoName & ".dotm"
  TemplateFile = p_TemplateFile
End Property


' ===== TemplatePath ==========================================================
' Full path to the template file in the repo. Read only. Template must
' match repo name.

Public Property Get TemplatePath() As String
  p_TemplatePath = Me.RepoDir & Application.PathSeparator & Me.TemplateFile
  TemplatePath = p_TemplatePath
End Property


' ===== SrcDir ================================================================
' Full path to directory VBA modules are exported to. Must be "src". Read only

Public Property Get SrcDir() As String
  p_SrcDir = Me.RepoDir & Application.Path & "src"
  ScrDir = p_SrcDir
End Property


' ===== SubmoduleExists =======================================================
' Boolean, does repo contain submodules

Public Property Get SubmoduleExists() As Boolean
  p_SubmoduleExists = Utils.IsItThere(Me.RepoDir & Application.PathSeparator & _
    ".gitmodules")
  SubmoduleExists = p_SubmoduleExists
End Property
