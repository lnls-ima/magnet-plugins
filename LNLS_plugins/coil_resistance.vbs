Option Explicit

Dim coil_name, Doc, Sol, ptype, pid, r, iNbOfCoils
Set Doc= getDocument()
Set Sol = Doc.getSolution()
ptype = Sol.getType
iNbOfCoils= Doc.getNumberOfCoils()
If iNbOfCoils > 0 Then
	If ptype=infoStaticSolution OR ptype=infoTimeHarmonicSolution Then
		coil_name = InputBox("Coil name:", "Get coil resistance", Doc.getPathOfCoil(1))
	   	If coil_name <> "" Then
	   		pid = 1
	   		r=Sol.getDCResistanceOfCoil(pid,coil_name)
			MsgBox("Resistance: " & r)
		End If
	ElseIf ptype=infoTransientSolution Then
		coil_name = InputBox("Coil name:", "Get coil resistance", Doc.getPathOfCoil(1))
	   	If coil_name <> "" Then
			ReDim ArrayOfValues(1)
			ArrayOfValues(0) = 1  ' Problem ID
			ArrayOfValues(1) = 0  ' Time instant
			r=Sol.getDCResistanceOfCoil(ArrayOfValues,coil_name)
			MsgBox("Resistance: " & r)
		End If
	Else
		MsgBox("The model is not solved")
	End If
Else
	MsgBox("No coil in model")
End If
