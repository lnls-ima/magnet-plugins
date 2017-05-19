Option Explicit

Import("auxiliary_functions.vbs")

Const rkstep = 0.0001 '[m]
Const max_loops = 1e8

Dim objFSO
Dim EmptyVar
Dim Doc, Mesh, Fieldx, Fieldy, Fieldz
Dim lim_xmin, lim_ymin, lim_zmin
Dim lim_xmax, lim_ymax, lim_zmax
Dim out_of_lim

Call ParticleTrajectory()


Sub ParticleTrajectory()

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

  FileName = DocumentName & "_trajectory.txt"
  FullFilename = objFSO.BuildPath(FilePath, Filename)

  Dim nproblem
  nproblem = GetProblemNumber(Doc, "Particle Trajectory")
	If isNull(nproblem) Then Exit Sub End If
	If not isNumeric(nproblem) Then Exit Sub End If

  Dim xmin
  Dim ymin
  Dim zmin
  Dim zmax
  Dim ztemp
  Dim energy

  energy = GetVariableValue("Particle energy (GeV)", "Kick Map", "3", EmptyVar)
  If isNull(energy) Then Exit Sub End If
  energy = (1e9)*energy

  xmin = GetVariableValue("Initial X (mm)", "Particle Trajectory", "0", EmptyVar)
  If isNull(xmin) Then Exit Sub End If

  ymin = GetVariableValue("Initial Y (mm)", "Particle Trajectory", "0", EmptyVar)
  If isNull(ymin) Then Exit Sub End If

  zmin = GetVariableValue("Initial Z (mm)", "Particle Trajectory", "0", EmptyVar)
  If isNull(zmin) Then Exit Sub End If

  zmax = GetVariableValue("Final Z (mm)", "Particle Trajectory", "500", EmptyVar)
  If isNull(zmax) Then Exit Sub End If

  If zmin > zmax Then
    ztemp = zmin
    zmin = zmax
    zmax = ztemp
  End If

  Set Mesh = Doc.getSolution.getMesh(nproblem)

  Call Mesh.getGeometricExtents(lim_xmin, lim_ymin, lim_zmin, lim_xmax, lim_ymax, lim_zmax)

  If xmin < lim_xmin or xmin > lim_xmax Then
    MsgBox("Initial X is out of the field matrix.")
    Exit Sub
  End If

  If ymin < lim_ymin or ymin > lim_ymax Then
    MsgBox("Initial Y is out of the field matrix.")
    Exit Sub
  End If

  If zmin < lim_zmin Then
    MsgBox("Initial Z is out of the field matrix.")
    Exit Sub
  End If

  If zmax > lim_zmax Then
    MsgBox("Final Z is out of the field matrix.")
    Exit Sub
  End If

  Set Fieldx = Doc.getSolution.getSystemField(Mesh,"B x")
  Set Fieldy = Doc.getSolution.getSystemField(Mesh,"B y")
  Set Fieldz = Doc.getSolution.getSystemField(Mesh,"B z")

  xmin   = xmin/1000
  ymin   = ymin/1000
  zmin   = zmin/1000
  zmax   = zmax/1000

  out_of_lim = False

  Dim r(6)
  r(0) = xmin
  r(1) = ymin
  r(2) = zmin
  r(3) = 0
  r(4) = 0
  r(5) = 1

  Dim finalpos
  finalpos = Array()

  Dim trajectory
  trajectory = RungeKuttaTrajectory(energy, r, zmax, rkstep)

  If (out_of_lim) Then
    MsgBox("Particle travelled out of the field matrix.")
    Exit Sub
  End If

  Call DrawTrajectory(trajectory)

  Call SaveTrajectory(trajectory, FullFilename)
  MsgBox("Trajectory saved in file: " & vbCrlf & vbCrlf & FileName)

End Sub


Function GetMagnetField(ByVal r)
  Dim bx, by, bz
  Dim field(3)

  field(0) = 0
  field(1) = 0
  field(2) = 0

  If (1000*r(0) < lim_xmin) or (1000*r(0) > lim_xmax) or (1000*r(1) < lim_ymin) or (1000*r(1) > lim_ymax) or (1000*r(2) < lim_zmin) or (1000*r(2) > lim_zmax) Then
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


Function RungeKuttaTrajectory(ByVal energy, ByVal r0, ByVal zmax, ByVal rkstep)

  Dim beta, brho, alpha

  Dim b, b1, b2, b3
  Dim r(6), r1(6), r2(6), r3(6)
  Dim drds1, drds2, drds3, drds4

  Dim i, inpts, npts, count

  Dim trajectory
  trajectory = Array(Array())

  beta = CalcBeta(energy)
  brho = CalcBrho(energy, beta)
  alpha = 1.0/brho/beta

  inpts = 5
  npts = inpts
  count = 0

  For i=0 To 5
    r(i) = r0(i)
  Next

  ReDim Preserve trajectory(npts)
  trajectory(count) = r

  Do While (r(2) < zmax - rkstep)

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

    count = count + 1
    If (count > npts) Then
      npts = npts + inpts
      Redim Preserve trajectory(npts)
    End If

    trajectory(count) = r

  Loop

  Redim Preserve trajectory(count)
  RungeKuttaTrajectory = trajectory

End Function


Sub DrawTrajectory(trajectory)

  Dim n, i
  Dim TrajectoryCoordinates

  n = Ubound(trajectory)
  ReDim TrajectoryCoordinates(n-1,2)

  For i = 0 To n-1
    TrajectoryCoordinates(i, 0) = 1000*trajectory(i)(0)
    TrajectoryCoordinates(i, 1) = 1000*trajectory(i)(1)
    TrajectoryCoordinates(i, 2) = 1000*trajectory(i)(2)
  Next

  Call getDocument().beginUndoGroup("Draw Trajectory")
  Call Doc.getView().newPolylineAnnotation(TrajectoryCoordinates, infoFixedToModel, 16)
  Call getDocument().endUndoGroup()

End Sub


Sub SaveTrajectory(trajectory, filename)

  Dim objFile
  Set objFile = objFSO.CreateTextFile(filename, True)

  objFile.Write "x[m]    y[m]    z[m]    dx/ds    dy/ds    dz/ds" & vbCrlf
	objFile.Write "------------------------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrlf

  Dim i
  For i=0 To Ubound(trajectory)
    objFile.Write CStr(trajectory(i)(0)) & vbTab & vbTab
    objFile.Write CStr(trajectory(i)(1)) & vbTab & vbTab
    objFile.Write CStr(trajectory(i)(2)) & vbTab & vbTab
    objFile.Write CStr(trajectory(i)(3)) & vbTab & vbTab
    objFile.Write CStr(trajectory(i)(4)) & vbTab & vbTab
    objFile.Write CStr(trajectory(i)(5)) & vbCrlf
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
