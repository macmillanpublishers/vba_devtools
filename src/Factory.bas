Attribute VB_Name = "Factory"
Option Explicit

' ===== USE ===================================================================
' Functions here create instances of objects that require arguments (since
' Class_Initialize events can't have params). For each class, there is a public
' function here called CreateClassname that accepts arguments. That function
' instantiates a new object and calls the Init method of that function.

' Each class then needs to have (1) a private global variable to store whether
' the factory function has initialized the object already, and (2) a method
' named Init that will set the default properties.

' =============================================================================
' =============================================================================


' ===== CreateRepository ======================================================
' Creates and sets default properties.

Public Function CreateRepository(Path As String) As Repository
  Set CreateRepository = New Repository
  CreateRepository.Init RepoDir:=Path
End Function

