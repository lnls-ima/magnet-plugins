
Call Install_Plugins()

Sub Install_Plugins()
  Dim objFSO
  Set objFSO = CreateObject("Scripting.FileSystemObject")

  Dim objShell
  Set objShell = CreateObject("Wscript.Shell")

  Dim ExecDir
  ExecDir = getExecutableDirectory()

  Dim ScripyPath
  ScriptPath = objShell.CurrentDirectory

  Dim PluginsDir
  Dim PluginsInstallationDir
  PluginsDir = objFSO.BuildPath(ScriptPath, "LNLS_plugins")
  PluginsInstallationDir = objFSO.BuildPath(ExecDir, "LNLS_plugins")

  Const DeleteReadOnly = True

  On Error Resume Next
  objFSO.DeleteFile(objFSO.BuildPath(PluginsInstallationDir, "*.vbs")), DeleteReadOnly
  objFSO.CopyFolder PluginsDir, PluginsInstallationDir
  If (Err.Number = 70) Then
		MsgBox( "To install LNLS plugins launch MagNet in administrator mode." )
		Exit Sub
  End If
  On Error GoTo 0

  Dim MaterialsDir
  Dim MaterialsFilename
  MaterialsDir = objFSO.BuildPath(ScriptPath, "LNLS_materials")
  MaterialsFilename = objFSO.BuildPath(MaterialsDir, "LNLS Materials.xml")
  objFSO.CopyFile MaterialsFilename, ExecDir

  Dim EventHandlersFilename
  Dim EventHandlersNewFilename
  EventHandlersFilename = objFSO.BuildPath(ExecDir, "EventHandlers.vbs")
  EventHandlersNewFilename = objFSO.BuildPath(ExecDir, "EventHandlers_OriginalFile.vbs")

  On Error Resume Next
  objFSO.CopyFile EventHandlersFilename, EventHandlersNewFilename, 0
  If Err Then
    If (Err.Number <> 58) Then
        MsgBox( Err.Description)
        Exit Sub
    End If
  End If
  On Error GoTo 0

  Dim objFile
  Dim FileContent
  Dim FileContentSplit
  Dim FirstPartContent
  Dim SecondPartContent
  Dim NewLines
  Dim NewContent
  Dim DrawShapePath
  Dim VertexParamPath
  Dim FieldSamplerPath
  Dim FieldMultipolesPath
  Dim CoilResistancePath
  Dim ParticleTrajectoryPath
  Dim KickMapPath

  DrawShapePath = objFSO.BuildPath(PluginsInstallationDir, "draw_shape.vbs")
  VertexParamPath = objFSO.BuildPath(PluginsInstallationDir, "vertex_parametrization.vbs")
  FieldSamplerPath = objFSO.BuildPath(PluginsInstallationDir, "field_sampler.vbs")
  FieldMultipolesPath = objFSO.BuildPath(PluginsInstallationDir, "field_multipoles.vbs")
  CoilResistancePath = objFSO.BuildPath(PluginsInstallationDir, "coil_resistance.vbs")
  ParticleTrajectoryPath = objFSO.BuildPath(PluginsInstallationDir, "particle_trajectory.vbs")
  KickMapPath = objFSO.BuildPath(PluginsInstallationDir, "kick_map.vbs")

  Set objFile = objFSO.OpenTextFile(EventHandlersFilename, 1)
  FileContent = objFile.ReadAll
  objFile.Close

  FileContentSplit = Split(FileContent, "Sub Application_OnLoad()")
  FirstPartContent = FileContentSplit(0)
  SecondPartContent = FileContentSplit(1)

  NewLines =  vbCrLf & _
              Chr(9) & "'LNLS Plugins Start'" & vbCrLf & _
              Chr(9) & "'================================================================================================='" & vbCrLf & _
              Chr(9) & "Const sMacroMenuName = ""LNLS"" " & vbCrLf & _
              Chr(9) & "Dim menubar, macromenu, command, dirpath" & vbCrLf & _
              Chr(9) & "Set menubar = getMenubar()" & vbCrLf & _
              Chr(9) & "Set macromenu = menubar.insertMenu(""&Help"", sMacroMenuName)" & vbCrLf & vbCrLf & _
              Chr(9) & "command = ""runScript("" & Chr(34) & """ & DrawShapePath & """ & Chr(34) & "")"" " & vbCrLf & _
              Chr(9) & "macromenu.appendItem ""Draw Shape"", command" & vbCrLf & vbCrLf & _
              Chr(9) & "command = ""runScript("" & Chr(34) & """ & VertexParamPath & """ & Chr(34) & "")"" " & vbCrLf & _
              Chr(9) & "macromenu.appendItem ""Vertex Parametrization"", command" & vbCrLf & vbCrLf & _
              Chr(9) & "command = ""runScript("" & Chr(34) & """ & FieldSamplerPath & """ & Chr(34) & "")""" & vbCrLf & _
              Chr(9) & "macromenu.appendItem ""Field Sampler"", command" & vbCrLf & vbCrLf & _
              Chr(9) & "command = ""runScript("" & Chr(34) & """ & FieldMultipolesPath & """ & Chr(34) & "")""" & vbCrLf & _
              Chr(9) & "macromenu.appendItem ""Field Multipoles"", command" & vbCrLf & vbCrLf & _
              Chr(9) & "command = ""runScript("" & Chr(34) & """ & CoilResistancePath & """ & Chr(34) & "")""" & vbCrLf & _
              Chr(9) & "macromenu.appendItem ""Coil Resistance"", command" & vbCrLf & vbCrLf & _
              Chr(9) & "command = ""runScript("" & Chr(34) & """ & ParticleTrajectoryPath & """ & Chr(34) & "")""" & vbCrLf & _
              Chr(9) & "macromenu.appendItem ""Particle Trajectory"", command" & vbCrLf & vbCrLf & _
              Chr(9) & "command = ""runScript("" & Chr(34) & """ & KickMapPath & """ & Chr(34) & "")""" & vbCrLf & _
              Chr(9) & "macromenu.appendItem ""Kick Map"", command" & vbCrLf & _
              Chr(9) & "'================================================================================================='" & vbCrLf & _
              Chr(9) & "'LNLS Plugins End'"

  If InStr(SecondPartContent, "LNLS Plugins") Then
    Dim TempA
    Dim TempB
    TempA =  Split(SecondPartContent, "'LNLS Plugins Start'")(0)
    TempB =  Split(SecondPartContent, "'LNLS Plugins End'")(1)
    SecondPartContent = TempA & TempB
  End If

  NewContent = FirstPartContent & "Sub Application_OnLoad()" & NewLines & SecondPartContent

  Set objFile = objFSO.OpenTextFile(EventHandlersFilename, 2)
  objFile.WriteLine NewContent
  objFile.Close

  MsgBox("Installation complete.")

end Sub
