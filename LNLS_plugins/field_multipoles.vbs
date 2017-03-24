Import("auxiliary_functions.vbs")

Const CommentChar = "#"
Const ParametersFilename = "field_multipoles_inputs.txt"
Const NormalComponent = "B y"
Const SkewComponent = "B x"

Call FieldMultipoles()


Sub FieldMultipoles()

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
		Filename = ProcessDocumentName(DocumentName) + "_field_multipoles_pid" + CStr(nproblem) + ".txt"
	Else
		nproblems = nproblem
		Filename = ProcessDocumentName(DocumentName) + "_field_multipoles.txt"
	End If

	Dim normal_coeffs_pn, skew_coeffs_pn
	Dim normal_coeffs, skew_coeffs
	ReDim normal_coeffs(FittingOrder-1, Ubound(nproblems))
	ReDim skew_coeffs(FittingOrder-1, Ubound(nproblems))

	For i=0 to Ubound(nproblems)
		normal_coeffs_pn = GetProblemMultipoles(Doc, nproblems(i), xmin, xpoints, xstep, z, FittingOrder, NormalComponent)
		skew_coeffs_pn = GetProblemMultipoles(Doc, nproblems(i), xmin, xpoints, xstep, z, FittingOrder, SkewComponent)
		For j=0 to FittingOrder-1
			normal_coeffs(j, i) = normal_coeffs_pn(j)
			skew_coeffs(j, i) = skew_coeffs_pn(j)
		Next
	Next

	Call WriteMultipolesFile( objFSO.BuildPath(FilePath, Filename), nproblems, normal_coeffs, skew_coeffs)

	Call OpenMultipolesDialog( Filename, nproblems, normal_coeffs, skew_coeffs)

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


Sub WriteMultipolesFile(FullFilename, nproblems, normal_coeffs, skew_coeffs)

	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile( FullFilename, True)

	Dim TempStr
	TempStr = "n" & vbTab
	For i = 0 To Ubound(nproblems)
		TempStr = TempStr & "Bn (PID " & nproblems(i) & ")" & vbTab & "An (PID " & nproblems(i) & ")" & vbTab
	Next
	objFile.Write TempStr & vbCrLf

	Dim nd
	nd = 7

	For i = 0 To Ubound(normal_coeffs, 1)
		TempStr = CStr(i) & vbTab
		For j = 0 To Ubound(normal_coeffs, 2)
			TempStr = TempStr & ScientificNotation(normal_coeffs(i, j), nd, False) & vbTab
			TempStr = TempStr & ScientificNotation(skew_coeffs(i, j), nd, False) & vbTab
		Next
		objFile.Write TempStr & vbCrLf
	Next
	objFile.Close

End Sub


Sub OpenMultipolesDialog(Filename, nproblems, normal_coeffs, skew_coeffs)

	If Ubound(nproblems) = 0 Then

		Dim nc(5)
		Dim sc(5)
		For i = 0 To 4
			nc(i) = nan
			sc(i) = nan
		Next

		If (Ubound(normal_coeffs) < 5) Then
			For i = 0 To Ubound(normal_coeffs, 1)
				nc(i) = normal_coeffs(i, 0)
			Next
		Else
			For i = 0 To 4
				nc(i) = normal_coeffs(i, 0)
			Next
		End If

		If (Ubound(skew_coeffs) < 5) Then
			For i = 0 To Ubound(skew_coeffs, 1)
				sc(i) = skew_coeffs(i, 0)
			Next
		Else
			For i = 0 To 4
				sc(i) = skew_coeffs(i, 0)
			Next
		End If

		Dim nd
		nd = 4

		MsgBox("Problem " & nproblems(0) & vbTab & vbTab & "Normal" & vbTab & vbTab & "Skew" & _
		vbCrLf & vbCrLf & "Dipole" & vbTab & _
		vbTab & ScientificNotation(nc(0), nd, True) & _
		vbTab & ScientificNotation(sc(0), nd, True) & _
	  vbCrLf & vbCrLf & "Quadrupole" & _
		vbTab & ScientificNotation(nc(1), nd, True) & _
		vbTab & ScientificNotation(sc(1), nd, True) & _
	  vbCrLf & vbCrLf & "Sextupole" & vbTab & _
		vbTab & ScientificNotation(nc(2), nd, True) & _
		vbTab & ScientificNotation(sc(2), nd, True) & _
	  vbCrLf & vbCrLf & "Octupole" & vbTab & _
		vbTab & ScientificNotation(nc(3), nd, True) & _
		vbTab & ScientificNotation(sc(3), nd, True) & _
	  vbCrLf & vbCrLf & "Decapole" & vbTab & _
		vbTab & ScientificNotation(nc(4), nd, True) & _
		vbTab & ScientificNotation(sc(4), nd, True) & _
		vbCrLf & vbCrLf & "Multipoles table save in file: " & vbCrLf & Filename)

	Else
		MsgBox("Multipoles table save in file: " & vbCrLf & Filename)
	End If

End Sub


Function ScientificNotation(floVal, NumberofDigits, AddSpace)
  Dim floAbsVal, intSgnVal, intScale, floScaled, floStr

  If Not isNumeric(floVal) Then
    ScientificNotation = ""
    Exit Function
  End If

  floAbsVal = Abs(floVal)
  If floAbsVal <> 0 Then
    intSgnVal = Sgn(floVal)
    intScale = Int(Log(floAbsVal) / Log(10))
    floScaled = floAbsVal / (10 ^ intScale)

		If intSgnVal < 0 Then
			floStr = FormatNumber(intSgnVal * floScaled, NumberofDigits)
		Else
			floStr = "+" & FormatNumber(intSgnVal * floScaled, NumberofDigits)
		End If

    If Sgn(intScale) < 0 Then
			If AddSpace Then
      	ScientificNotation = " " & floStr & "e" & CStr(intSCale)
			Else
				ScientificNotation = floStr & "e" & CStr(intSCale)
			End If
    Else
      ScientificNotation = floStr & "e+" & CStr(intSCale)
    End If
  Else
    ScientificNotation = "+" & FormatNumber(0, NumberofDigits) & "e+0"
  End If

End Function


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
