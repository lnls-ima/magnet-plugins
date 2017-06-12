Option Explicit

Import("auxiliary_functions.vbs")

Const rkstep = 0.0001 '[m]
Const tolerance = 1e-9 '[m]
Const max_loops = 1e8

Dim objFSO
Dim EmptyVar
Dim Window
Dim Doc, Mesh, Fieldx, Fieldy, Fieldz
Dim lim_xmin, lim_ymin, lim_zmin
Dim lim_xmax, lim_ymax, lim_zmax
Dim out_of_lim

Call KickMap()


Sub KickMap()

  Set objFSO = CreateObject("Scripting.FileSystemObject")

  If hasDocument() Then
    Set Doc = getDocument()
  Else
    MsgBox("The application does not have a document open.")
    Exit Sub
  End If

  Dim FilePath, DocumentName, FileName, FullFilename

  FilePath = getDocumentPath(Doc)
	If isNull(FilePath) Then Exit Sub End If

	DocumentName = getDocumentName(Doc)
	If isNull(DocumentName) Then Exit Sub End If

  DocumentName = ProcessDocumentName(DocumentName)

  FileName = DocumentName & "_kickmap.txt"
  FullFilename = objFSO.BuildPath(FilePath, Filename)

  Dim nproblem
  nproblem = GetProblemNumber(Doc, "Kick Map")
	If isNull(nproblem) Then Exit Sub End If
	If not isNumeric(nproblem) Then Exit Sub End If

  Dim Range
  Dim xmin, xmax, xpoints, xstep
  Dim ymin, ymax, ypoints, ystep
  Dim zmin, zmax, ztemp, length
  Dim energy

  energy = GetVariableValue("Particle energy (GeV)", "Kick Map", "3", EmptyVar)
  If isNull(energy) Then Exit Sub End If
  energy = (1e9)*energy

  Range = GetVariableRange("X coordinate (mm)", "Kick Map", "0 10 1", EmptyVar)
	If isNull(Range) Then Exit Sub End If
	xmin    = Range(0)
	xmax    = Range(1)
	xpoints = Range(2)
  xstep   = Range(3)

	Range = GetVariableRange("Y coordinate (mm)", "Kick Map", "0 10 1", EmptyVar)
	If isNull(Range) Then Exit Sub End If
	ymin    = Range(0)
	ymax    = Range(1)
	ypoints = Range(2)
  ystep   = Range(3)

  zmin = GetVariableValue("Initial Z (mm)", "Kick Map", "0", EmptyVar)
  If isNull(zmin) Then Exit Sub End If

  zmax = GetVariableValue("Final Z (mm)", "Kick Map", "500", EmptyVar)
  If isNull(zmax) Then Exit Sub End If

  MsgBox("The kick map calculation may take several minutes." & vbCrlf & "The application will be locked until it is finished.")

  If zmin > zmax Then
    ztemp = zmin
    zmin = zmax
    zmax = ztemp
  End If

  Set Mesh = Doc.getSolution.getMesh(nproblem)

  Call Mesh.getGeometricExtents(lim_xmin, lim_ymin, lim_zmin, lim_xmax, lim_ymax, lim_zmax)

  If xmin < lim_xmin - 1000*tolerance Then
    MsgBox("Initial X is out of the field matrix.")
    Exit Sub
  End If

  If xmax > lim_xmax + 1000*tolerance Then
    MsgBox("Final X is out of the field matrix.")
    Exit Sub
  End If

  If ymin < lim_ymin - 1000*tolerance Then
    MsgBox("Initial Y is out of the field matrix.")
    Exit Sub
  End If

  If ymax > lim_ymax + 1000*tolerance Then
    MsgBox("Final Y is out of the field matrix.")
    Exit Sub
  End If

  If zmin < lim_zmin - 1000*tolerance Then
    MsgBox("Initial Z is out of the field matrix.")
    Exit Sub
  End If

  If zmax > lim_zmax + 1000*tolerance Then
    MsgBox("Final Z is out of the field matrix.")
    Exit Sub
  End If

  xmin   = xmin/1000
	xmax   = xmax/1000
  xstep  = xstep/1000
  ymin   = ymin/1000
	ymax   = ymax/1000
  ystep  = ystep/1000
  zmin   = zmin/1000
  zmax   = zmax/1000

  length = zmax - zmin

  out_of_lim = False

  Set Fieldx = Doc.getSolution.getSystemField(Mesh,"B x")
  Set Fieldy = Doc.getSolution.getSystemField(Mesh,"B y")
  Set Fieldz = Doc.getSolution.getSystemField(Mesh,"B z")

  Dim i, j
  Dim xpos, ypos
  ReDim xpos(xpoints)
  ReDim ypos(ypoints)

  For i=0 To ypoints-1
    For j=0 To xpoints-1
      ypos(i) = ymin + ystep*i
      xpos(j) = xmin + xstep*j
    Next
  Next

  Dim ks, kickx, kicky
  ReDim kickx(ypoints, xpoints)
  ReDim kicky(ypoints, xpoints)

  Dim r(6)
  r(0) = 0 : r(1) = 0 : r(2) = zmin
  r(3) = 0 : r(4) = 0 : r(5) = 1

  For i=0 To ypoints-1
    For j=0 To xpoints-1
      r(0) = xpos(j)
      r(1) = ypos(i)
      ks = GetKicks(energy, r, zmax, rkstep)
      kickx(i,j) = ks(0)
      kicky(i,j) = ks(1)
    Next
  Next

  If (out_of_lim) Then
    Dim button
    button = MsgBox("At least one trajectory travelled out of the field matrix." & vbCrlf & "Save Kick Map?", vbYesNo)
    If (button <> 6) Then Exit Sub End If
  End If

  Call WriteKickMap(FullFilename, energy, length, xpos, ypos, kickx, kicky)
  MsgBox("Kick Map saved in file: " & vbCrlf & vbCrlf & FileName)

End Sub


Function GetMagnetField(ByVal r)
  Dim bx, by, bz
  Dim field(3)

  field(0) = 0
  field(1) = 0
  field(2) = 0

  If (1000*r(0) < lim_xmin - tolerance) or (1000*r(0) > lim_xmax + tolerance) Then
    out_of_lim = True
    GetMagnetField = field
    Exit Function
  End If

  If (1000*r(1) < lim_ymin - tolerance) or (1000*r(1) > lim_ymax + tolerance) Then
    out_of_lim = True
    GetMagnetField = field
    Exit Function
  End If

  If (1000*r(2) < lim_zmin - tolerance) or (1000*r(2) > lim_zmax + tolerance) Then
    out_of_lim = True
    GetMagnetField = field
    Exit Function
  End If

  Call Fieldx.getFieldAtPoint(r(0)*1000, r(1)*1000, r(2)*1000, bx)
  Call Fieldy.getFieldAtPoint(r(0)*1000, r(1)*1000, r(2)*1000, by)
  Call Fieldz.getFieldAtPoint(r(0)*1000, r(1)*1000, r(2)*1000, bz)

  field(0) = bx(0)
  field(1) = by(0)
  field(2) = bz(0)

  GetMagnetField = field
End Function


Function GetKicks(ByVal energy, ByVal r0, ByVal zmax, ByVal rkstep)

  Dim beta, brho, alpha

  Dim b, b1, b2, b3
  Dim r(6), r1(6), r2(6), r3(6)
  Dim drds1, drds2, drds3, drds4

  Dim i, inpts, npts, count

  beta = CalcBeta(energy)
  brho = CalcBrho(energy, beta)
  alpha = 1.0/brho/beta

  inpts = 5
  npts = inpts
  count = 0

  For i=0 To 5
    r(i) = r0(i)
  Next

  Do While (r(2) < (zmax - rkstep))

    If (count >= max_loops) Then Exit Do End If

    b = GetMagnetField(r)
    drds1 = NewtonLorentzEquation(alpha, r, b)
    For i = 0 To 5
      r1(i) = r(i) + (rkstep/2.0)* drds1(i)
    Next

    b1 = GetMagnetField(r1)
    drds2 = NewtonLorentzEquation(alpha, r1, b1)
    For i = 0 To 5
      r2(i) = r(i) + (rkstep/2.0)* drds2(i)
    Next

    b2 = GetMagnetField(r2)
    drds3 = NewtonLorentzEquation(alpha, r2, b2)
    For i = 0 To 5
      r3(i) = r(i) + rkstep* drds3(i)
    Next

    b3 = GetMagnetField(r3)
    drds4 = NewtonLorentzEquation(alpha, r3, b3)

    For i = 0 To 5
      r(i) = r(i) + (rkstep/6.0)*(drds1(i) + 2.0*drds2(i) + 2.0*drds3(i) + drds4(i))
    Next

  Loop

  Dim kicks(2)
  kicks(0) = r(3)*(brho^2)
  kicks(1) = r(4)*(brho^2)

  GetKicks = kicks

End Function


Sub WriteKickMap(filename, energy, length, xpos, ypos, kickx, kicky)

  Dim objFile
  Set objFile = objFSO.CreateTextFile(filename, True)

  Dim nx, ny, i, j, ndig
  nx = Ubound(xpos)
  ny = Ubound(ypos)
  ndig = 7

  objFile.Write "# KICKMAP" & vbCrlf
  objFile.Write "# Author: Magnet User, Date: " & GetDate() & vbCrlf
  objFile.Write "# Total Length of Longitudinal Interval [m]" & vbCrlf
  objFile.Write CStr(length) & vbCrlf
  objFile.Write "# Number of Horizontal Points" & vbCrlf
  objFile.Write CStr(nx) & vbCrlf
  objFile.Write "# Number of Vertical Points" & vbCrlf
  objFile.Write CStr(ny) & vbCrlf
  objFile.Write "# Horizontal KickTable [T2m2]" & vbCrlf
  objFile.Write "START" & vbCrlf

  objFile.Write "               "
  For j=0 To nx-1
    objFile.Write CStr(ScientificNotation((xpos(j)), ndig)) & " "
  Next
  objFile.Write vbCrlf

  For i=0 To ny-1
    objFile.Write CStr(ScientificNotation((ypos(i)), ndig)) & " "
    For j=0 To nx-1
      objFile.Write CStr(ScientificNotation(kickx(i,j), ndig)) & " "
    Next
    objFile.Write vbCrlf
  Next

  objFile.Write "# Vertical KickTable [T2m2]" & vbCrlf
  objFile.Write "START" & vbCrlf

  objFile.Write "               "
  For j=0 To nx-1
    objFile.Write CStr(ScientificNotation(xpos(j), ndig)) & " "
  Next
  objFile.Write vbCrlf

  For i=0 To ny-1
    objFile.Write CStr(ScientificNotation(ypos(i), ndig)) & " "
    For j=0 To nx-1
      objFile.Write CStr(ScientificNotation(kicky(i,j), ndig)) & " "
    Next
    objFile.Write vbCrlf
  Next

  objFile.Close

End Sub


Sub Import(strFile)

  Dim wshShell, objFile, strCode

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set wshShell = CreateObject("Wscript.Shell")
	strFile = WshShell.ExpandEnvironmentStrings(strFile)
	strFile = objFSO.GetAbsolutePathName(strFile)
	Set objFile = objFSO.OpenTextFile(strFile)
	strCode = objFile.ReadAll
	objFile.Close
	ExecuteGlobal strCode

End Sub
