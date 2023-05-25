Option Explicit
On Error Resume Next
Dim intDaysOlderThan, objShellApp, colNamedArguments, objFso, objFolder, objFile, strFileExtension, strFolderPath, strFullPath
Set colNamedArguments = WScript.Arguments.Named
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
Set objShellApp = WScript.CreateObject("Shell.Application")

If IsEmpty(colNamedArguments.Item("folderpath")) Then
  WScript.Echo "/folderpath is a required parameter"
  WScript.Quit 1
Else
  strFolderPath = colNamedArguments.Item("folderpath")
End If

' Check the specified folder path exists
If Not objFso.FolderExists(strFolderPath) Then
  WScript.Echo "The specified folder does not exist"
  WScript.Quit 1
End If

If IsEmpty(colNamedArguments.Item("daysolderthan")) Then
  intDaysOlderThan = 30
Else
  intDaysOlderThan = CInt(colNamedArguments.Item("daysolderthan"))
End If

If IsEmpty(colNamedArguments.Item("filextension")) Then
  strFileExtension = "html"
Else
  strFileExtension = colNamedArguments.Item("filextension")
End If

Call DeleteFiles(strFolderPath, intDaysOlderThan)

Sub DeleteFiles(path, days)
  Set objFolder = objFso.GetFolder(path)
  For Each objFile In objFolder.Files
    If objFile.DateCreated < (Now() - days) and (StrComp(CStr(objFso.GetExtensionName(objFile)), strFileExtension, 1) = 0) Then
      strFullPath = objFso.GetAbsolutePathName(objFile)
      ' This recycles rather than deletes - use 'objFile.Delete(True)' to delete
      objShellApp.Namespace(0).ParseName(strFullPath).InvokeVerb("delete")
    End If
  Next
End Sub
