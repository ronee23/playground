Option Explicit

Dim objFso
Dim shell
Dim inputOption
Dim bMasterFlag
Dim bFlag
Dim inputDrive
Dim readFile
Dim strNextLine

bMasterFlag=True

While bMasterFlag

	inputOption = InputBox(Chr(13) & "Choose Option -" & Chr(13) & Chr(13) & "	Copy Patch: 1" & Chr(13) & Chr(13) & "	Execute Patch: 2", "Engineer Shell Application") 
	
	inputOption = Mid(inputOption,1,1)

	'Initialize FSO
	Set objFso = CreateObject("scripting.filesystemobject")

	If inputOption=CStr(1) Then
		bFlag=False
		While Not (bFlag)
		inputDrive = InputBox("Enter USB Drive" & Chr(13) & Chr(13) & "ex: E or F or G etc", "Engineer Shell Application - Copy Patch")
		inputDrive = Mid(inputDrive,1,1)
		If objFso.DriveExists(inputDrive) Then
			bFlag = True
		Else
			MsgBox "Drive " & inputDrive & ":\ does not exist", 16, "Engineer Shell Application - Copy Patch"
		End If
		If bFlag = True Then
			If objFso.FileExists(inputDrive & ":\patch\" & "FileList.ini") Then
				objFso.CopyFile inputDrive & ":\patch\" & "FileList.ini" , "C:\setup\" , TRUE
				'Read FileList.ini
				set readFile = objFso.OpenTextFile("C:\setup\FileList.ini", 1, false)
				Do Until readFile.AtEndOfStream
					strNextLine = readFile.Readline
					If objFso.FileExists(inputDrive & ":\patch\" & strNextLine) Then
						objFso.CopyFile inputDrive & ":\patch\" & strNextLine, "C:\setup\" , TRUE
					Else
						MsgBox "File defined in FileList.ini: " & Chr(13) & Chr(13) & inputDrive & ":\patch\" & strNextLine & Chr(13) & Chr(13) & "does not Exist.", 48, "Engineer Shell Application - Copy Patch"
					End If
				Loop
				readfile.close
			Else
				MsgBox "File FileList.ini: " & Chr(13) & Chr(13) & inputDrive & ":\patch\" & "FileList.ini" & Chr(13) & Chr(13) & "does not Exist.", 48, "Engineer Shell Application - Copy Patch"
			End If
		End If
		Wend
	End If
		
	If inputOption=CStr(2) Then
		set shell=createobject("wscript.shell")
		If objFso.FileExists("C:\setup\ExecutePatch.bat") Then
			shell.run "C:\setup\ExecutePatch.bat"
		Else
			MsgBox "File: C:\setup\ExecutePatch.bat does not exist", 16, "Engineer Shell Application - Execute Patch"
		End If
	End If

	If inputOption="S" Then
		set shell = createobject("wscript.shell")
		shell.run "C:\psainst\shutdown.bat"
	End If
	
	If Not (inputOption=CStr(1) or inputOption=CStr(2)) Then
		MsgBox "Invalid Option: " & inputOption & Chr(13) & Chr(13) & "Please contact administrator", 16, "Engineer Shell Application"
	End If
	
	Set objFso=Nothing
	set shell=nothing

Wend
