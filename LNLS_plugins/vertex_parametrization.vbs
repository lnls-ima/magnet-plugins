
Call VertexParametrization()

Sub VertexParametrization()

  n = getDocument().getView().getSelection().getNumberOfObjects()

  If n = 0 Then
    MsgBox("No vertex selected.")
    Exit Sub
  End If

  LengthUnit = LCase(getDocument().getDefaultLengthUnit())
  UnitSyntax = LengthUnitSyntax(LengthUnit)

  ShapeFunction = InputBox("Shape:", "Vertex Parametrization", "y = 0.1*(x^2)")

  If (Len( ShapeFunction ) = 0) Then
    Exit Sub
  ElseIf (InStr( ShapeFunction, "=") = 0) Or (InStr( ShapeFunction, "y") = 0) Then
    MsgBox("Invalid shape.")
    Exit Sub
  End If

  Dim ids
  ReDim ids(n)

  For i=0 to n-1
    ids(i) = getDocument().getView().getSelection().getObjectID(i)(0)
  Next

  Call getDocument().beginUndoGroup("Vertex Parametrization")

  For i=0 to n-1
    Call getDocument().getProblem(1).getVertexGlobalPosition(ids(i), x, ytemp, ztemp)

    yval = Split( ShapeFunction, "y")(1)
    yval = Split( yval, "=")(1)
    yvalsplit = Split( yval, "x")

    If (Ubound(yvalsplit) = 0) Then
      y = "(" & yvalsplit(0) & ")"
    Else
      y = "("
      For k = 0 to Ubound(yvalsplit)-1
        y = y & yvalsplit(k) & x
      Next
      y = y & yvalsplit(Ubound(yvalsplit)) & ")"
    End If

    pos = "[" & x & UnitSyntax & "," & y & UnitSyntax & "]"

    Call getDocument().setParameter(ids(i), "Position", pos, infoArrayParameter)

  Next

  Call getDocument().endUndoGroup()

End Sub


Function LengthUnitSyntax(LengthUnit)

  Dim Syntax

  If (StrComp(LengthUnit, "kilometers", 1) = 0) Then
    Syntax = "%km"
  ElseIf (StrComp(LengthUnit, "meters", 1) = 0) Then
    Syntax = ""
  ElseIf (StrComp(LengthUnit, "centimeters", 1) = 0) Then
    Syntax = "%cm"
  ElseIf (StrComp(LengthUnit, "millimeters", 1) = 0) Then
    Syntax = "%mm"
  ElseIf (StrComp(LengthUnit, "microns", 1) = 0) Then
    Syntax = "%um"
  ElseIf (StrComp(LengthUnit, "miles", 1) = 0) Then
    Syntax = "%mi"
  ElseIf (StrComp(LengthUnit, "yards", 1) = 0) Then
    Syntax = "%yd"
  ElseIf (StrComp(LengthUnit, "feet", 1) = 0) Then
    Syntax = "%ft"
  ElseIf (StrComp(LengthUnit, "inches", 1) = 0) Then
    Syntax = "%in"
  End If

  LengthUnitSyntax = Syntax

End Function
