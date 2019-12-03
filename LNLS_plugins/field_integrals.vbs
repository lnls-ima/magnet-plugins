Import("auxiliary_functions.vbs")

Call FieldIntegrals()

Sub FieldIntegrals()

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

	Dim BoxTitle
	BoxTitle = "Field integrals"

	nproblem = GetProblemNumber(Doc, BoxTitle)
	If isNull(nproblem) Then Exit Sub End If

	DefaultName = ProcessDocumentName(DocumentName)

	Dim nproblems
	If isNumeric(nproblem) Then
		ReDim nproblems(0)
		nproblems(0) = nproblem
		DefaultName = DefaultName + "_field_integrals_pid" + CStr(nproblem) + ".txt"
	Else
		nproblems = nproblem
		DefaultName = DefaultName + "_field_integrals.txt"
	End If

	InitialPos = GetVector("Initial Position [x, y, z] (mm)", BoxTitle, "0 0 0", EmptyVar)
	If isNull(InitialPos) Then Exit Sub End If
	xi = InitialPos(0)
	yi = InitialPos(1)
	zi = InitialPos(2)

	FinalPos = GetVector("Final Position [x, y, z] (mm)", BoxTitle, "0 0 10", EmptyVar)
	If isNull(FinalPos) Then Exit Sub End If
	xf = FinalPos(0)
	yf = FinalPos(1)
	zf = FinalPos(2)

	StepPos = GetVariableValue("Step (mm)", BoxTitle, "0.01", EmptyVar)
	If isNull(StepPos) Then Exit Sub End If

	Filename = GetVariableString("Enter the file name:", BoxTitle, DefaultName, EmptyVar)
	If isNull(Filename) Then Exit Sub End If

	FullFilename = objFSO.BuildPath(FilePath, Filename)

	Set objFile = objFSO.CreateTextFile(FullFilename, True)

	Bxname = "B x"
	Byname = "B y"
	Bzname = "B z"

	dist = ((xf - xi)^2 + (yf - yi)^2 + (zf - zi)^2)^(1/2)
	npts = dist/StepPos + 1

	a = (xf - xi)/npts
	b = (yf - yi)/npts
	c = (zf - zi)/npts

	Dim Distance, BxArray, ByArray, BzArray
	ReDim Distance(npts)
	ReDim BxArray(npts)
	ReDim ByArray(npts)
	ReDim BzArray(npts)

	objFile.Write "PID" & vbTab & "IBx [T.m]" & vbTab & "IBy [T.m]" & vbTab & "IBz [T.m]" & vbCrlf

	text = "PID" & vbTab & "IBx [T.m]" & vbTab & vbTab & "IBy [T.m]" & vbTab & vbTab & "IBz [T.m]"

	ndigits = 4

	For i=0 to Ubound(nproblems)
	 	np = nproblems(i)

		Set Mesh = Doc.getSolution.getMesh( np )
		Set Field1 = Doc.getSolution.getSystemField (Mesh, Bxname)
		Set Field2 = Doc.getSolution.getSystemField (Mesh, Byname)
		Set Field3 = Doc.getSolution.getSystemField (Mesh, Bzname)

		For j=0 to npts
			x = xi + a*j
			y = yi + b*j
			z = zi + c*j

			Call Field1.getFieldAtPoint (x, y, z, bx)
			Call Field2.getFieldAtPoint (x, y, z, by)
			Call Field3.getFieldAtPoint (x, y, z, bz)

			Distance(j) = (((x - xi)^2 + (y - yi)^2 + (z - zi)^2)^(1/2))/1000
			BxArray(j) = bx(0)
			ByArray(j) = by(0)
			BzArray(j) = bz(0)

		Next

		ibx = TrapzIntegral(Distance, BxArray)
		iby = TrapzIntegral(Distance, ByArray)
		ibz = TrapzIntegral(Distance, BzArray)

		sbx = ScientificNotation(ibx, ndigits)
		sby = ScientificNotation(iby, ndigits)
		sbz = ScientificNotation(ibz, ndigits)

		objFile.Write CStr(np) & vbTab & sbx & vbTab & sby & vbTab & sbz & vbCrlf

		text = text & vbCrLf & vbCrLf & CStr(np) & vbTab & sbx & vbTab & sby & vbTab & sbz & vbCrlf

	Next

	objFile.Close
	MsgBox(text)

End Sub


Sub Import(strFile)

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set wshShell = CreateObject("Wscript.Shell")
	strFile = WshShell.ExpandEnvironmentStrings(strFile)
	strFile = objFSO.GetAbsolutePathName(strFile)
	Set objFile = objFSO.OpenTextFile(strFile)
	strCode = objFile.ReadAll
	objFile.Close
	ExecuteGlobal strCode

End Sub
