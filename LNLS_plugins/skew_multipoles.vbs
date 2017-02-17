Import("multipoles.vbs")

Const ParametersFilename = "skew_multipoles_parameters.txt"
Const FieldComponent = "B x"
Const MultipolesLabel = "skew"

Call FieldMultipoles( ParametersFilename, FieldComponent, Multipoleslabel )

Sub Import(strFile)

	Set objFs = CreateObject("Scripting.FileSystemObject")
	Set wshShell = CreateObject("Wscript.Shell")
	strFile = WshShell.ExpandEnvironmentStrings(strFile)
	strFile = objFs.GetAbsolutePathName(strFile)
	Set objFile = objFs.OpenTextFile(strFile)
	strCode = objFile.ReadAll
	objFile.Close
	ExecuteGlobal strCode

End Sub
