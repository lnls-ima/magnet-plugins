Import("auxiliary_functions.vbs")

Const CommentChar = "#"
Const ParametersFilename  = "field_sampler_inputs.txt"
Dim EmptyVar

Call FieldSampler()


Sub FieldSampler()

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
			XRange_FromFile	= Parameters(0)
			YRange_FromFile	= Parameters(1)
			ZRange_FromFile = Parameters(2)
		End If
	End If

	nproblem = GetProblemNumber(Doc, "Field sampling")
	If isNull(nproblem) Then Exit Sub End If
	If not isNumeric(nproblem) Then Exit Sub End If

	CoilInfo = GetCoilInfo(Doc, nproblem, "main")
	If not isNull(CoilInfo) Then
		MainCoilCurrent = CoilInfo(0)
		MainCoilTurns = CoilInfo(1)
	End If

	CoilInfo = GetCoilInfo(Doc, nproblem, "trim")
	If not isNull(CoilInfo) Then
		TrimCoilCurrent = CoilInfo(0)
		TrimCoilTurns = CoilInfo(1)
	End If

	CoilInfo = GetCoilInfo(Doc, nproblem, "ch")
	If not isNull(CoilInfo) Then
		CHCoilCurrent = CoilInfo(0)
		CHCoilTurns = CoilInfo(1)
	End If

	CoilInfo = GetCoilInfo(Doc, nproblem, "cv")
	If not isNull(CoilInfo) Then
		CVCoilCurrent = CoilInfo(0)
		CVCoilTurns = CoilInfo(1)
	End If

	CoilInfo = GetCoilInfo(Doc, nproblem, "qs")
	If not isNull(CoilInfo) Then
		QSCoilCurrent = CoilInfo(0)
		QSCoilTurns = CoilInfo(1)
	End If

	Dim BoxTitle
	BoxTitle = "Field sampling - Problem " & CStr( nproblem )

	Range = GetVariableRange("X coordinate (mm)", BoxTitle, "0 10 1", XRange_FromFile)
	If isNull(Range) Then Exit Sub End If
	xmin    = Range(0)
	xmax    = Range(1)
	xpoints = Range(2)
	xstep   = Range(3)

	Range = GetVariableRange("Y coordinate (mm)", BoxTitle, "0 10 1", YRange_FromFile)
	If isNull(Range) Then Exit Sub End If
	ymin    = Range(0)
	ymax    = Range(1)
	ypoints = Range(2)
	ystep   = Range(3)

	Range = GetVariableRange("Z coordinate (mm)", BoxTitle, "0 10 1", ZRange_FromFile)
	If isNull(Range) Then Exit Sub End If
	zmin    = Range(0)
	zmax    = Range(1)
	zpoints = Range(2)
	zstep   = Range(3)

	DefaultName = ProcessDocumentName(DocumentName)

	If (xpoints <> 1) Then DefaultName = DefaultName & "_X=" & CStr( xmin ) & "_" & CStr( xmax ) & "mm" End If
	If (ypoints <> 1) Then DefaultName = DefaultName & "_Y=" & CStr( ymin ) & "_" & CStr( ymax ) & "mm" End If
	If (zpoints <> 1) Then DefaultName = DefaultName & "_Z=" & CStr( zmin ) & "_" & CStr( zmax ) & "mm" End If

	If not IsEmpty( MainCoilCurrent ) Then DefaultName = DefaultName & "_Imc=" & CStr( MainCoilCurrent ) & "A" End If
	If not IsEmpty( TrimCoilCurrent ) Then DefaultName = DefaultName & "_Itc=" & CStr( TrimCoilCurrent ) & "A" End If
	If not IsEmpty( CHCoilCurrent ) Then DefaultName = DefaultName & "_Ich=" & CStr( CHCoilCurrent ) & "A" End If
	If not IsEmpty( CVCoilCurrent ) Then DefaultName = DefaultName & "_Icv=" & CStr( CVCoilCurrent ) & "A" End If
	If not IsEmpty( QSCoilCurrent ) Then DefaultName = DefaultName & "_Iqs=" & CStr( QSCoilCurrent ) & "A" End If

	DefaultName = DefaultName & ".txt"

	SplitName = Split( DefaultName, "_")
	MagnetName = SplitName(1)
	If (Ubound(SplitName) > 1) Then
		Model = SplitName(2)
	Else
		Model = ""
	End If

	Filename = GetVariableString("Enter the file name:", BoxTitle, DefaultName, EmptyVar)
	If isNull(Filename) Then Exit Sub End If

	FullFilename = objFSO.BuildPath(FilePath, Filename)

	Set objFile = objFSO.CreateTextFile(FullFilename, True)

	Set Mesh = Doc.getSolution.getMesh( nproblem )
	Set Field1 = Doc.getSolution.getSystemField (Mesh,"B x")
	Set Field2 = Doc.getSolution.getSystemField (Mesh,"B y")
	Set Field3 = Doc.getSolution.getSystemField (Mesh,"B z")

	objFile.Write "fieldmap_name:     " & vbTab & MagnetName & " " & Model & vbCrlf
	objFile.Write "timestamp:         " & vbTab & GetDate() & "_" & GetTime() & vbCrlf
	objFile.Write "filename:          " & vbTab & Filename & vbCrlf
	objFile.Write "nr_magnets:        " & vbTab & CStr( 1 ) & vbCrlf
	objFile.Write vbCrlf
	objFile.Write "magnet_name:       " & vbTab & MagnetName & vbCrlf
	objFile.Write "gap[mm]:           " & vbTab & vbCrlf
	objFile.Write "control_gap[mm]:   " & vbTab & "--" & vbCrlf
	objFile.Write "magnet_length[mm]: " & vbTab & vbCrlf

	If not IsEmpty( MainCoilCurrent ) Then
		objFile.Write "current_main[A]:   " & vbTab & CStr( MainCoilCurrent ) & vbCrlf
		objFile.Write "nr_turns_main:     " & vbTab & CStr( MainCoilCurrent*MainCoilTurns) & vbCrlf
	Else
		objFile.Write "current_main[A]:   " & vbTab & "--" & vbCrlf
		objFile.Write "nr_turns_main:     " & vbTab & "--" & vbCrlf
	End If

	If not IsEmpty( TrimCoilCurrent ) Then
		objFile.Write "current_trim[A]:   " & vbTab & CStr( TrimCoilCurrent ) & vbCrlf
		objFile.Write "nr_turns_trim:     " & vbTab & CStr( TrimCoilCurrent*TrimCoilTurns) & vbCrlf
	End If

	If not IsEmpty( CHCoilCurrent ) Then
		objFile.Write "current_ch[A]:     " & vbTab & CStr( CHCoilCurrent ) & vbCrlf
		objFile.Write "nr_turns_ch:       " & vbTab & CStr( CHCoilCurrent*CHCoilTurns ) & vbCrlf
	End If

	If not IsEmpty( CVCoilCurrent ) Then
		objFile.Write "current_cv[A]:     " & vbTab & CStr( CVCoilCurrent ) & vbCrlf
		objFile.Write "nr_turns_cv:       " & vbTab & CStr( CVCoilCurrent*CVCoilTurns ) & vbCrlf
	End If

	If not IsEmpty( QSCoilCurrent ) Then
		objFile.Write "current_qs[A]:     " & vbTab & CStr( QSCoilCurrent ) & vbCrlf
		objFile.Write "nr_turns_qs:       " & vbTab & CStr( QSCoilCurrent*QSCoilTurns ) & vbCrlf
	End If

	objFile.Write "center_pos_z[mm]:  " & vbTab & CStr( 0 ) & vbCrlf
	objFile.Write "center_pos_x[mm]:  " & vbTab & CStr( 0 ) & vbCrlf
	objFile.Write "rotation[deg]:     " & vbTab & CStr( 0 ) & vbCrlf
	objFile.Write vbCrlf
	objFile.Write "X[mm]" & vbTab & "Y[mm]" & vbTab & "Z[mm]" & vbTab & "Bx" & vbTab & "By" & vbTab & "Bz [T]" & vbCrlf
	objFile.Write "------------------------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrlf

	For k=0 to zpoints-1
		For j=0 to ypoints-1
			For i=0 to xpoints-1
				x = xmin + xstep*i
				y = ymin + ystep*j
				z = zmin + zstep*k
				Call Field1.getFieldAtPoint (x, y, z, bx)
				Call Field2.getFieldAtPoint (x, y, z, by)
				Call Field3.getFieldAtPoint (x, y, z, bz)
				objFile.Write CStr(x) & vbTab & CStr(y) & vbTab & CStr(z) & vbTab & CStr(bx(0)) & vbTab & CStr(by(0)) & vbTab & CStr(bz(0)) & vbCrlf
			Next
		Next
	Next

	objFile.Close


End Sub

Function GetCoilInfo(Doc, nproblem, CoilType)
	Dim NrCoils
	Dim CoilCurrent
	Dim CoilTurns
	Dim CoilName
	Dim LCoilName
	Dim CoilInfo
	ReDim CoilInfo(2)
	NrCoils = Doc.getNumberOfCoils()

	CoilType = LCase(CoilType)

	Dim i
	For i=1 to NrCoils
		CoilName = Doc.getPathOfCoil(i)
		LCoilName = LCase(CoilName)
		If InStr(LCoilName, CoilType) Then
			Call Doc.getProblem(nproblem).getParameter(CoilName, "Current", CoilCurrent)
			Call Doc.getProblem(nproblem).getParameter(CoilName, "NumberOfTurns", CoilTurns)
			CoilInfo(0) = CoilCurrent
			CoilInfo(1) = CoilTurns
			GetCoilInfo = CoilInfo
			Exit Function
		End If
	Next

	GetCoilInfo = Null

End Function


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
