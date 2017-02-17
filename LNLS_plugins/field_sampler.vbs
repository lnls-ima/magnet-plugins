Import("auxiliary_functions.vbs")

Const CommentChar = "#"
Const ParametersFilename  = "field_sampler_parameters.txt"
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

	If getDocument().getProblem().IsCoil("MainCoil") Then
		Call getDocument.getProblem( nproblem ).getParameter("MainCoil", "Current", MainCoilCurrent)
		Call getDocument.getProblem( nproblem ).getParameter("MainCoil", "NumberOfTurns", MainCoilTurns)
	End If

	If getDocument().getProblem().IsCoil("TrimCoil") Then
		Call getDocument.getProblem( nproblem ).getParameter("TrimCoil", "Current", TrimCoilCurrent)
		Call getDocument.getProblem( nproblem ).getParameter("TrimCoil", "NumberOfTurns", TrimCoilTurns)
	End If

	If getDocument().getProblem().IsCoil("CHCoil") Then
		Call getDocument.getProblem( nproblem ).getParameter("CHCoil", "Current", CHCoilCurrent)
		Call getDocument.getProblem( nproblem ).getParameter("CHCoil", "NumberOfTurns", CHCoilTurns)
	End If

	If getDocument().getProblem().IsCoil("CVCoil") Then
		Call getDocument.getProblem( nproblem ).getParameter("CVCoil", "Current", CVCoilCurrent)
		Call getDocument.getProblem( nproblem ).getParameter("CVCoil", "NumberOfTurns", CVCoilTurns)
	End If

	If getDocument().getProblem().IsCoil("QSCoil") Then
		Call getDocument.getProblem( nproblem ).getParameter("QSCoil", "Current", QSCoilCurrent)
		Call getDocument.getProblem( nproblem ).getParameter("QSCoil", "NumberOfTurns", QSCoilTurns)
	End If

	Dim BoxTitle
	BoxTitle = "Field sampling - Problem " & CStr( nproblem )

	Range = GetVariableRange("X", BoxTitle, "0 10 1", XRange_FromFile)
	If isNull(Range) Then Exit Sub End If
	xmin    = Range(0)
	xmax    = Range(1)
	xpoints = Range(2)
	xstep   = Range(3)

	Range = GetVariableRange("Y", BoxTitle, "0 10 1", YRange_FromFile)
	If isNull(Range) Then Exit Sub End If
	ymin    = Range(0)
	ymax    = Range(1)
	ypoints = Range(2)
	ystep   = Range(3)

	Range = GetVariableRange("Z", BoxTitle, "0 10 1", ZRange_FromFile)
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

	Filename = GetVariableString("Enter the file name:", BoxTitle, DefaultName, EmptyVar)
	If isNull(Filename) Then Exit Sub End If

	FullFilename = objFSO.BuildPath(FilePath, Filename)

	Set objFile = objFSO.CreateTextFile(FullFilename, True)

	Set Mesh = Doc.getSolution.getMesh( nproblem )
	Set Field1 = Doc.getSolution.getSystemField (Mesh,"B x")
	Set Field2 = Doc.getSolution.getSystemField (Mesh,"B y")
	Set Field3 = Doc.getSolution.getSystemField (Mesh,"B z")

	objFile.Write "fieldmap_name:       " & vbCrlf
	objFile.Write "timestamp:           " & GetDate() & "_" & GetTime() & vbCrlf
	objFile.Write "filename:      			" & Filename & vbCrlf
	objFile.Write "nr_magnets:       		" & CStr( 1 ) & vbCrlf
	objFile.Write vbCrlf
	objFile.Write "magnet_name:         " & vbCrlf
	objFile.Write "gap[mm]:             " & vbCrlf
	objFile.Write "control_gap[mm]:     " & vbCrlf
	objFile.Write "magnet_length[mm]:   " & vbCrlf

	If not IsEmpty( MainCoilCurrent ) Then
		objFile.Write "current_main[A]:     " & CStr( MainCoilCurrent ) & vbCrlf
		objFile.Write "NI_main[A.esp]:      " & CStr( MainCoilCurrent*MainCoilTurns) & vbCrlf
	Else
		objFile.Write "current_main[A]:     " & vbCrlf
		objFile.Write "NI_main[A.esp]:      " & vbCrlf
	End If

	If not IsEmpty( TrimCoilCurrent ) Then
		objFile.Write "current_trim[A]:     " & CStr( TrimCoilCurrent ) & vbCrlf
		objFile.Write "NI_trim[A.esp]:      " & CStr( TrimCoilCurrent*TrimCoilTurns) & vbCrlf
	End If

	If not IsEmpty( CHCoilCurrent ) Then
		objFile.Write "current_ch[A]:       " & CStr( CHCoilCurrent ) & vbCrlf
		objFile.Write "NI_ch[A.esp]:        " & CStr( CHCoilCurrent*CHCoilTurns ) & vbCrlf
	End If

	If not IsEmpty( CVCoilCurrent ) Then
		objFile.Write "current_cv[A]:       " & CStr( CVCoilCurrent ) & vbCrlf
		objFile.Write "NI_cv[A.esp]:        " & CStr( CVCoilCurrent*CVCoilTurns ) & vbCrlf
	End If

	If not IsEmpty( QSCoilCurrent ) Then
		objFile.Write "current_qs[A]:       " & CStr( QSCoilCurrent ) & vbCrlf
		objFile.Write "NI_qs[A.esp]:        " & CStr( QSCoilCurrent*QSCoilTurns ) & vbCrlf
	End If

	objFile.Write "center_pos_z[mm]: 		" & CStr( 0 ) & vbCrlf
	objFile.Write "center_pos_x[mm]: 		" & CStr( 0 ) & vbCrlf
	objFile.Write "rotation[deg]:       " & CStr( 0 ) & vbCrlf
	objFile.Write vbCrlf
	objFile.Write "X[mm]		Y[mm]		Z[mm]		B x, B y, B z   (T)" & vbCrlf
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
				objFile.Write CStr(x) & vbTab & vbTab & CStr(y) & vbTab & vbTab & CStr(z) & vbTab & vbTab & CStr(bx(0)) & vbTab & vbTab & CStr(by(0)) & vbTab & vbTab & CStr(bz(0)) & vbCrlf
			Next
		Next
	Next

	objFile.Close

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
