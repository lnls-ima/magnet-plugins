Import("auxiliary_functions.vbs")

Const CommentChar = "#"

Sub FieldMultipoles(ParametersFilename, FieldComponent, Multipoleslabel)

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If hasDocument() Then
		Set Doc = getDocument()
	Else
		MsgBox("The application does not have a document open.")
		Exit Sub
	End If

	FilePath = getDocumentPath(Doc)
	If isNull(FilePath) Then Exit Sub End If

	DocumentName = getDocumentName(Doc)
	If isNull(DocumentName) Then Exit Sub End If

	Parameters = GetParametersFromFile( objFSO.BuildPath(FilePath, ParametersFilename), CommentChar)
	If not isNull(Parameters) Then
		If Ubound(Parameters) = 2 Then
			XRange_FromFile				= Parameters(0)
			Z_FromFile						= Parameters(1)
			FittingOrder_FromFile = Parameters(2)
		End If
	End If

	nproblem = GetProblemNumber(Doc, "Field Multipoles")
	If isNull(nproblem) Then Exit Sub End If

	Range = GetVariableRange("X", "Field Multipoles", "-10, 10, 101", XRange_FromFile)
	If isNull(Range) Then Exit Sub End If
	xmin    = Range(0)
	xmax    = Range(1)
	xpoints = Range(2)
	xstep   = Range(3)

	z = GetVariableValue("Z", "Field Multipoles", "0", Z_FromFile)
	If isNull(z) Then Exit Sub End If

	FittingOrder = GetVariableValue("Fitting Order", "Field Multipoles", "15", FittingOrder_FromFile)
	If isNull(FittingOrder) Then Exit Sub End If
	If (FittingOrder < 1) Then
		MsgBox("Invalid fitting order.")
		Exit Sub
	End If

	Dim nproblems
	If isNumeric(nproblem) Then
		ReDim nproblems(0)
		nproblems(0) = nproblem
		Filename = ProcessDocumentName(DocumentName) + "_" + Multipoleslabel + "_multipoles_pid" + CStr(nproblem) + ".txt"
	Else
		nproblems = nproblem
		Filename = ProcessDocumentName(DocumentName) + "_" + Multipoleslabel + "_multipoles.txt"
	End If

	Dim ncoeffs
	Dim coeffs
	ReDim coeffs(FittingOrder-1, Ubound(nproblems))

	For i=0 to Ubound(nproblems)
		ncoeffs = GetProblemMultipoles(Doc, nproblems(i), xmin, xpoints, xstep, z, FittingOrder, FieldComponent)
		For j=0 to FittingOrder-1
			coeffs(j, i) = ncoeffs(j)
		Next
	Next

	Call WriteMultipolesFile( objFSO.BuildPath(FilePath, Filename), nproblems, coeffs)

	Call OpenMultipolesDialog( Filename, nproblems, coeffs, Multipoleslabel)

End Sub


Function GetProblemMultipoles(Doc, nproblem, xmin, xpoints, xstep, z, FittingOrder, FieldComponent)

	Set Mesh = Doc.getSolution().getMesh( nproblem )
	Set Field = Doc.getSolution().getSystemField (Mesh, FieldComponent)

	Dim ncoeffs
	Dim x_mm
	Dim XVector
	ReDim XVector(xpoints)

	Dim FieldValue
	Dim FieldVector
	ReDim FieldVector(xpoints)

	For i = 0 to xpoints-1
		x_mm = xmin + xstep*i
		Call Field.getFieldAtPoint( x_mm, 0, z, FieldValue)
		XVector(i) = x_mm/1000
		FieldVector(i) = FieldValue(0)
	Next

	ncoeffs =  PolynomialFitting(XVector, FieldVector, FittingOrder)
	GetProblemMultipoles = ncoeffs

End Function


Sub WriteMultipolesFile(FullFilename, nproblems, coeffs)

	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile( FullFilename, True)

	Dim TempStr
	TempStr = "n" & vbTab
	For i = 0 To Ubound(nproblems)
		TempStr = TempStr & "Bn (PID " & nproblems(i) & ")" & vbTab
	Next
	objFile.Write TempStr & vbCrLf

	For i = 0 To Ubound(coeffs, 1)
		TempStr = CStr(i) & vbTab
		For j = 0 To Ubound(coeffs, 2)
			TempStr = TempStr & CStr(coeffs(i, j)) & vbTab
		Next
		objFile.Write TempStr & vbCrLf
	Next
	objFile.Close

End Sub


Sub OpenMultipolesDialog(Filename, nproblems, coeffs, Multipoleslabel)

	If Ubound(nproblems) = 0 Then

		Dim c(5)
		For i = 0 To 4
			c(i) = nan
		Next

		If (Ubound(coeffs) < 5) Then
			For i = 0 To Ubound(coeffs, 1)
				c(i) = coeffs(i, 0)
			Next
		Else
			For i = 0 To 4
				c(i) = coeffs(i, 0)
			Next
		End If

		MsgBox("Problem " & nproblems(0) & _
		vbCrLf & vbCrLf & Multipoleslabel & " dipole:      " & c(0) & _
	  vbCrLf & vbCrLf & Multipoleslabel & " quadrupole:  " & c(1) & _
	  vbCrLf & vbCrLf & Multipoleslabel & " sextupole:   " & c(2) & _
	  vbCrLf & vbCrLf & Multipoleslabel & " octupole:    " & c(3) & _
	  vbCrLf & vbCrLf & Multipoleslabel & " decapole:    " & c(4) & _
		vbCrLf & vbCrLf & "Multipoles table save in file: " & vbCrLf & Filename)

	Else
		MsgBox("Multipoles table save in file: " & vbCrLf & Filename)
	End If

End Sub


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
