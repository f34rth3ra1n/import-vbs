Sub Import (library)
	Set FileSysObj = CreateObject("Scripting.FileSystemObject")
	Set FileObj = FileSysObj.OpenTextFile(library, 1)
	ExecuteGlobal FileObj.ReadAll
	FileObj.Close
	Set FileSysObj = Nothing
	Set FileObj = Nothing
End Sub
