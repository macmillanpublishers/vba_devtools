VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' let's test some things

Sub GitVbaStatus()
  Dim strWorkingDir As String
  strWorkingDir = "C:\Users\erica.warren\Documents\GitHub\Word-template"
  Dim blnDidItWork As Boolean
  blnDidItWork = Dev.VbaStatus(strWorkingDir)
  Debug.Print "DID IT WORK???? : " & blnDidItWork
End Sub
