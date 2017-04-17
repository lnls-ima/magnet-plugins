Option Explicit

Import("auxiliary_functions.vbs")

Const rkstep = 0.0001 '[m]
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
  Dim zmin, zmax, length
  Dim energy

  Call Doc.getSolution.getProblem(nproblem).getGeometricExtents("", lim_xmin, lim_ymin, lim_zmin, lim_xmax, lim_ymax, lim_zmax)

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

  zmin = GetVariableValue("Initial Z (mm)", "Kick Map", CStr(lim_zmin), EmptyVar)
  If isNull(zmin) Then Exit Sub End If

  zmax = GetVariableValue("Final Z (mm)", "Kick Map", CStr(lim_zmax), EmptyVar)
  If isNull(zmax) Then Exit Sub End If

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

  Set Mesh = Doc.getSolution.getMesh(nproblem)
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

  Set Window = CreateStatusWindow()
  Window.document.write "<html><body bgcolor=buttonface>Calculating particle trajectories: <span id='output'></span> %</body></html>"
  Window.document.title = "Kick Map"
  Window.resizeto 320, 120
  Window.moveto 200, 200
  UpdateStatusBar(0)

  Dim count, frac
  count = 0
  For i=0 To ypoints-1
    For j=0 To xpoints-1
      r(0) = xpos(j)
      r(1) = ypos(i)
      ks = GetKicks(energy, r, zmax, rkstep)
      kickx(i,j) = ks(0)
      kicky(i,j) = ks(1)
      count = count + 1
      frac = 100*count/(ypoints*xpoints)
      On Error Resume Next
        UpdateStatusBar(frac)
      If (Err.Number) Then
        Window.close
        Exit Sub
      End If
      On Error Goto 0
    Next
  Next

  UpdateStatusBar(100)
  Window.close

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


Sub UpdateStatusBar(value)
    Dim StrValue
    StrValue = FormatNumber(value, 2)
    Window.output.innerhtml = StrValue
End Sub


Function CreateStatusWindow()
    Dim signature, shellwnd, proc
    On Error Resume Next
    Set CreateWindow = nothing
    signature = left(createobject("Scriptlet.TypeLib").guid, 38)

    Dim cmd1, cmd2, cmd3, cmd4
    cmd1 = "<script>moveTo(-32000,-32000);</script>"
    cmd2 = "<hta:application id=app border=dialog minimizebutton=no maximizebutton=no scroll=no showintaskbar=yes contextmenu=no selection=no innerborder=no />"
    cmd3 = "<object id='shellwindow' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object>"
    cmd4 = "<script>shellwindow.putproperty('" & signature & "',document.parentWindow);</script>"
    Set proc = createobject("WScript.Shell").exec("mshta about:""" & cmd1 & cmd2 & cmd3 & cmd4 & """")

    Do
        If proc.status > 0 Then Exit Function
        For Each shellwnd in createobject("Shell.Application").windows
            Set CreateWindow = shellwnd.getproperty(signature)
            if err.number = 0 Then Exit Function
            err.clear
        Next
    Loop
End Function


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
