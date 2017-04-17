
Const electron_rest_energy = 510998.92811 '[eV]
Const light_speed  = 299792458 '[m/s]


Function CalcBeta(ByVal energy)
  CalcBeta = Sqr( 1.0 - (1.0 / ((energy / electron_rest_energy)^2)) )
End Function


Function CalcBrho(ByVal energy, ByVal beta)
  CalcBrho = (beta * energy / light_speed)
End Function


Function ScientificNotation(floVal, NumberofDigits)
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
      If Len(CStr(intScale)) = 2 Then
			   ScientificNotation = floStr & "e-0" & CStr(Abs(intSCale))
      Else
        ScientificNotation = floStr & "e-" & CStr(Abs(intSCale))
      End If
    Else
      If Len(CStr(intScale)) = 1 Then
        ScientificNotation = floStr & "e+0" & CStr(intSCale)
      Else
        ScientificNotation = floStr & "e+" & CStr(intSCale)
      End If
    End If
  Else
    ScientificNotation = "+" & FormatNumber(0, NumberofDigits) & "e+00"
  End If

End Function


Function getDocumentPath(Doc)

	If (Len( Doc.getFilePath()) <> 0) Then

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.GetFile( Doc.getFilePath() )
		getDocumentPath = objFSO.GetParentFolderName( objFile )

	Else
		MsgBox("Document not saved.")
		getDocumentPath = Null
		Exit Function
	End If

End Function


Function getDocumentName(Doc)

	If (Len( Doc.getFilePath()) <> 0) Then

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.GetFile( Doc.getFilePath() )
		getDocumentName = objFSO.GetBaseName( objFile )

	Else
		MsgBox("Document not saved.")
		getDocumentName = Null
		Exit Function
	End If

End Function


Function ProcessDocumentName(DocumentName)

	Dim count
	Dim DefaultNameSplit
	Dim DefaultName
	Dim DateStr

	DefaultName = Split( DocumentName, ".mn")(0)
	DefaultName = Split( DefaultName, "_solved")(0)

	If (InStr( DefaultName, "20") = 1) Then
		DefaultNameSplit = Split( DefaultName, "_")
		DefaultName = DefaultNameSplit(1)
		For count = 2 To (Ubound(DefaultNameSplit))
			DefaultName = DefaultName + "_" + DefaultNameSplit( count )
		Next
	End If

	DateStr = GetDate()
	DefaultName = DateStr & "_" & DefaultName
	ProcessDocumentName = DefaultName

End Function


Function GetProblemNumber(Doc, DialogTitle)

	Dim UserInput
	Dim nproblem

	UserInput = InputBox("Problem Number", DialogTitle, "1")

	If (Len( UserInput ) = 0) Then
		GetProblemNumber = Null
		Exit Function
	End If

	If (StrComp(UserInput, "all", 1) = 0) Then
		GetProblemNumber = Doc.getSolution().getSolvedProblems()

	Else
		nproblem = CDbl( UserInput )

		If Doc.getSolution().isSolved(nproblem) Then
			GetProblemNumber = nproblem
		Else
			MsgBox("Problem " & nproblem & " is not solved.")
			GetProblemNumber = Null
			Exit Function
		End If

	End If

End Function


Function GetVariableRange(VariableLabel, DialogTitle, DefaultValues, FromFile)

  Dim UserInput
  Dim MinValue
  Dim MaxValue
  Dim Nrpts
  Dim VariableStep
  Dim Range
  ReDim Range(4)

  If FromFile <> "" Then
		UserInput = FromFile
	Else
		UserInput = InputBox( VariableLabel & vbLf & "Start, end, iterations", DialogTitle, DefaultValues)
	End If

	If (Len( UserInput ) = 0) Then
		GetVariableRange = Null
		Exit Function
	End If

	UserInput = Split(UserInput)
	MinValue = CDbl(UserInput(0))
	MaxValue = CDbl(UserInput(1))
	Nrpts = CDbl(UserInput(2))

	If (Nrpts <> 1) Then
		VariableStep = (MaxValue-MinValue) / (Nrpts-1)
	Else
		VariableStep = 1
	End If

  Range(0) = MinValue
  Range(1) = MaxValue
  Range(2) = Nrpts
  Range(3) = VariableStep

  GetVariableRange = Range

End Function


Function GetVariableValue(VariableLabel, DialogTitle, DefaultValue, FromFile)

  Dim UserInput
  Dim VariableValue

	If FromFile <> "" Then
		UserInput = FromFile
	Else
		UserInput = InputBox(VariableLabel, DialogTitle, DefaultValue)
	End If

	If (Len( UserInput ) = 0) Then
		GetVariableValue = Null
		Exit Function
	End If

	VariableValue = CDbl(UserInput)
	GetVariableValue = VariableValue

End Function


Function GetVariableString(VariableLabel, DialogTitle, DefaultValue, FromFile)

  Dim UserInput

	If FromFile <> "" Then
		UserInput = FromFile
	Else
		UserInput = InputBox(VariableLabel, DialogTitle, DefaultValue)
	End If

	If (Len( UserInput ) = 0) Then
		GetVariableString = Null
		Exit Function
	End If

	GetVariableString = Trim(UserInput)

End Function


Function GetParametersFromFile(Filename, CommentChar)

	Set objFSO=CreateObject("Scripting.FileSystemObject")

	If (objFSO.FileExists(Filename)) Then
		Dim ParamFile
		Dim Content

		Set ParamFile = objFSO.OpenTextFile(Filename , 1)
		Content = ParamFile.ReadAll
		Content = Split( Content, vbCrLf)
		ParamFile.Close

		Dim Parameters
		ReDim Parameters(Ubound(Content))
		Dim line
		Dim count

		count = 0
		For i=0 to Ubound(Content)
			line = Split( Content(i), CommentChar )(0)
			If line <> "" Then
				Parameters(count) = line
				count = count + 1
			End If
		Next

		ReDim Preserve Parameters(count-1)
		GetParametersFromFile = Parameters

	Else
		GetParametersFromFile = Null
	End If

End Function


Function GetTime()

	Dim time_str
	Dim temp_str
	Dim hour_str, min_str, sec_str

	temp_str = Time
	hour_str = Hour( temp_str )
	min_str = Minute( temp_str )
	sec_str = Second( temp_str )
	hour_str = string( 2 - Len( hour_str ), "0") & hour_str
	min_str = string( 2 - Len( min_str ), "0") & min_str
	sec_str = string( 2 - Len( sec_str ), "0") & sec_str

	time_str = hour_str & "-" & min_str & "-" & sec_str

	GetTime = time_str

End Function


Function GetDate()

	Dim date_str
	Dim temp_str
	Dim year_str, month_str, day_str

	temp_str = Date
	year_str = Year( temp_str )
	month_str = Month( temp_str )
	day_str = Day( temp_str )
	month_str = string( 2 - Len( month_str ), "0") & month_str
	day_str = string( 2 - Len( day_str ), "0") & day_str

	date_str = year_str & "-" & month_str & "-" & day_str

	GetDate = date_str

End Function


Function NewtonLorentzEquation(ByVal alpha, ByVal r, ByVal b)

  Dim drds(6)

  drds(0) = r(3)
  drds(1) = r(4)
  drds(2) = r(5)
  drds(3) = -alpha * (r(4) * b(2) - r(5) * b(1))
  drds(4) = -alpha * (r(5) * b(0) - r(3) * b(2))
  drds(5) = -alpha * (r(3) * b(1) - r(4) * b(0))

  NewtonLorentzEquation = drds

End Function


Function InverseMatrix(Matrix)

  Dim i, k, j, n
  Dim p, temp, kd

  n = Ubound(Matrix, 1)

  Dim vecP()
  ReDim vecP(n)

  Dim MatrixL()
  ReDim MatrixL(n, n)

  Dim MatrixU()
  ReDim MatrixU(n, n)

  Dim MatrixB
  ReDim MatrixB(n, n)

  Dim InvMatrix
  ReDim InvMatrix(n, n)

  Dim vecX()
  ReDim vecX(n)

  Dim vecY()
  ReDim vecY(n)

  'Initial values
  For i = 0 To (n-1)
    vecP(i) = i
    vecX(i) = 0
    vecY(i) = 0

    For j = 0 To (n-1)
      MatrixU(i,j) = 0
      MatrixL(i,j) = 0
      InvMatrix(i,j) = 0
      MatrixB(i,j) = 0
    Next
    MatrixB(i,i) = 1

  Next

  For k = 0 To (n-2)

    'Find pivot
    p = 0
    For i = k To (n-1)
      temp = Abs(Matrix(i, k))
      If (temp > p) Then
        p = temp
        kd = i
      End If
    Next

    'Check if the matrix is singular
    If (p = 0) Then
      MsgBox("Error: Singular Matrix")
      InverseMatrix = InvMatrix
      Exit Function
    End If

    'Exchange matrix rows
    temp = vecP(kd)
    vecP(kd) = vecP(k)
    vecP(k) = temp

    For i = 0 To (n-1)
      temp = Matrix(kd,i)
      Matrix(kd, i) = Matrix(k, i)
      Matrix(k, i) = temp
    Next

  Next

  'Decompose as LU
  For j = 0 To (n-1)
    MatrixL(j,j) = 1

    For i = 0 To j
      s1 = 0
      For k = 0 To i-1
        s1 = s1 + MatrixU(k,j)*MatrixL(i,k)
      Next
      MatrixU(i,j) = Matrix(i,j) - s1
    Next

    For i = j To (n-1)
      s2 = 0
      For k = 0 To j-1
        s2 = s2 + MatrixU(k,j)*MatrixL(i,k)
      Next

			'Check if the matrix is singular
      If (MatrixU(j,j) = 0) Then
        MsgBox("Error: Singular Matrix")
        InverseMatrix = InvMatrix
        Exit Function
      End If

      MatrixL(i,j) = (Matrix(i,j) - s2)/MatrixU(j,j)
    Next

  Next

  For i = 0 To (n-1)

    'Foward solve LY = PB
    For k = 0 To (n-1)
      vecY(k) = MatrixB(vecP(k), i)
      For m = 0 To k-1
        vecY(k) = vecY(k) - MatrixL(k,m)*vecY(m)
      Next
    Next

    'Backward solve UX = Y
    k = n - 1
    Do
      vecX(k) = vecY(k)
      For m = k+1 To (n-1)
        vecX(k) = vecX(k) - MatrixU(k,m)*vecX(m)
      Next
      vecX(k) = vecX(k)/MatrixU(k,k)
      k = k-1
    Loop Until ( k < 0 )

    For j = 0 To (n-1)
      Matrix(i,j) = vecX(j)
    Next

  Next

  For i = 0 To (n-1)
    For j = 0 To (n-1)
      InvMatrix(i,j) = Matrix(j,i)
    Next
  Next

  InverseMatrix = InvMatrix

End Function


Function PolynomialFitting( x(), y(), order )

  Dim i, j
  Dim n

  n = Ubound(x)

  Dim MatrixA
  ReDim MatrixA(order, order)

  Dim VecB
  ReDim VecB(order)

  Dim coeffs
  ReDim coeffs(order)

  For i = 0 To order-1
    For j = 0 To order-1
      MatrixA(i,j) = 0
      For k = 0 To (n-1)
        MatrixA(i,j) = MatrixA(i,j) + x(k)^(i+j)
      Next
    Next
  Next

  Dim InvMatrixA
  InvMatrixA = InverseMatrix(MatrixA)

  For i = 0 To order-1
    VecB(i) = 0
    For k = 0 To (n-1)
      VecB(i) = VecB(i) + (x(k)^i)*y(k)
    Next
  Next

  For i = 0 To order-1
    coeffs(i) = 0
    For k = 0 To order-1
      coeffs(i) = coeffs(i) + InvMatrixA(i, k)*VecB(k)
    Next
  Next

  PolynomialFitting = coeffs

End Function
