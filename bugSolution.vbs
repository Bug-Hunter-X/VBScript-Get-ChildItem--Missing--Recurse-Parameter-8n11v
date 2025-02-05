Function GetChildItemsRecursive(strFolder)
  Dim fso, folder, file, arrFiles
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set folder = fso.GetFolder(strFolder)
  Set arrFiles = CreateObject("Scripting.Dictionary")

  For Each file In folder.Files
    arrFiles.Add file.Path, file.Path
  Next

  For Each subfolder In folder.SubFolders
    Call GetChildItemsRecursive(subfolder.Path)
  Next

  ' Process the files found.
  For Each strFilePath In arrFiles.Items
    WScript.Echo strFilePath
  Next

End Function

' Example usage
Call GetChildItemsRecursive("C:\Temp") ' Replace with your path