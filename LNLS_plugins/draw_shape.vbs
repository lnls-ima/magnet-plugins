
'This function try to approximate an arbitrary shape by a set of line segments with the specified length.
'The error for the start and end point of each line segment stay whitin the tolerance choose for the function value (YTol [%]).
'Whenever possible, the tolerance for the segment length (LenTol [%]) will also be respected.
'The maximum number of segments (MaxNrSegs) and the maximum number of interactions to achieve the length tolerance (MaxLenTolSearch) were included to avoid infinity loops.

Call DrawShape()

Sub DrawShape()

	Const YTol = 0.01
	Const LenTol = 0.01
	Const MaxNrSegs = 1000
	Const MaxLenTolSearch = 1000

	Dim Sucess
	Dim ErrorMsg
	Dim ShapeFunction
	Dim XInput
	Dim LengthUnit
	Dim length
	Dim xmin, xmax, x0, xr, xtemp
	Dim y, y0, yr, yplus, yminus
	Dim a, b, count

	If hasDocument() Then
		Set Doc= getDocument()
	Else
		MsgBox("The application does not have a document open.")
		Exit Sub
	End If

	LengthUnit = LCase(getDocument().getDefaultLengthUnit())
	LengthUnit = LengthUnitSyntax(LengthUnit)

	ShapeFunction = InputBox("Shape:", "Draw Shape", "y = 0.1*(x^2)")

  If (Len( ShapeFunction ) = 0) Then
    Exit Sub
  ElseIf (InStr( ShapeFunction, "=") = 0) Or (InStr( ShapeFunction, "y") = 0) Then
    MsgBox("Invalid shape.")
    Exit Sub
  End If

  XInput = InputBox("x coordinate [" & LengthUnit & "]:" & vbLf & "Start, End", "Draw Shape", "-10 10")
  If (Len( XInput ) = 0) Then Exit Sub End If

  XInput = Split( Trim( XInput) )
  If ((Ubound( Xinput) = 1)) Then
  	xmin = CDbl( XInput( 0 ) )
  	xmax = CDbl( XInput( 1 ) )
  	If (xmax < xmin) Then xtemp = xmin : xmin = xmax : xmax = xtemp End If
  Else
    MsgBox("Wrong number of inputs for the horizontal coordinate x.")
    Exit Sub
  End If

  length = InputBox("Line segment length [" & LengthUnit & "]:", "Draw Shape", "1")
  If (Len( length ) = 0) Then Exit Sub End If

  length = CDbl( length )
  If ( length <= 0) Then
    MsgBox("Line segment length must be greater than zero.")
    Exit Sub
  End If

  x = xmin
  On Error Resume Next
  Execute( ShapeFunction )
  If (Err.Number) Then
		MsgBox( Err.Description )
		Exit Sub
  End If
  On Error GoTo 0

  x0 = x
	y0 = y
  count = 0

	Call getDocument().beginUndoGroup("Draw Shape")

  Do While (x < xmax)
    x = x0 + length

    On Error Resume Next
    Execute( ShapeFunction )
    If (Err.Number) Then
			Call getDocument().endUndoGroup()
      Call getDocument().undo()
      MsgBox( Err.Description )
      Exit Sub
    End If
    On Error GoTo 0

    TolCount = 0
    Do While ( Abs( Sqr( (x - x0)^2 + (y - y0)^2 ) - length)/length > LenTol )

      xr = x : yr = y
      a = (yr - y0)/(xr-x0)
      b = yr - a*xr
      x = (x0 + ( (1 + a^2)*(length^2) - (b + a*x0 - y0)^2 )^(0.5) + a*(-b + y0))/(1 + a^2)
      If (x > xmax) Then x = xmax End If

      On Error Resume Next
      Execute( ShapeFunction )
      If (Err.Number) Then
				Call getDocument().endUndoGroup()
      	Call getDocument().undo()
        MsgBox( Err.Description )
        Exit Sub
      End If
      On Error GoTo 0

      If (x = xmax) Then Exit Do End If

      TolCount = TolCount + 1
      If (TolCount > MaxLenTolSearch) Then Exit Do End If

    Loop

    If ( (length^2 - x^2 + 2*x*x0 - x0^2) >= 0  And x <> xmax) Then

      yplus = y0 + Sqr(length^2 - x^2 + 2*x*x0 - x0^2)
      yminus = y0 - Sqr(length^2 - x^2 + 2*x*x0 - x0^2)

      If ( Abs(y - yplus) <= Abs(y - yminus) ) Then
        If ( Abs( (y - yplus)/ y ) < YTol) Then
          y = yplus
        End If
      Else
        If ( Abs( (y - yminus)/ y ) < YTol) Then
          y = yminus
        End If
      End If

    End If

    Call getDocument().getView().newLine( x0, y0, x, y)
    x0 = x
		y0 = y

    count = count + 1
    If (count > MaxNrSegs) Then
			Call getDocument().endUndoGroup()
			Call getDocument().undo()
      MsgBox( "Maximum number of line segments reached (" & MaxNrSegs & " segments)." )
      Exit Sub
    End If

  Loop

	Call getDocument().endUndoGroup()

End Sub


Function LengthUnitSyntax(LengthUnit)

  Dim Syntax

  If (StrComp(LengthUnit, "kilometers", 1) = 0) Then
    Syntax = "km"
  ElseIf (StrComp(LengthUnit, "meters", 1) = 0) Then
    Syntax = "m"
  ElseIf (StrComp(LengthUnit, "centimeters", 1) = 0) Then
    Syntax = "cm"
  ElseIf (StrComp(LengthUnit, "millimeters", 1) = 0) Then
    Syntax = "mm"
  ElseIf (StrComp(LengthUnit, "microns", 1) = 0) Then
    Syntax = "um"
  ElseIf (StrComp(LengthUnit, "miles", 1) = 0) Then
    Syntax = "mi"
  ElseIf (StrComp(LengthUnit, "yards", 1) = 0) Then
    Syntax = "yd"
  ElseIf (StrComp(LengthUnit, "feet", 1) = 0) Then
    Syntax = "ft"
  ElseIf (StrComp(LengthUnit, "inches", 1) = 0) Then
    Syntax = "in"
  End If

  LengthUnitSyntax = Syntax

End Function
