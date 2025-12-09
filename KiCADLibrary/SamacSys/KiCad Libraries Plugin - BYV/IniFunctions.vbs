' Source: http://www.robvanderwoude.com/vbstech_files_ini.php
Function ReadIni (myFilePath, mySection, myKey)
	Dim intEqualPos
	Dim objFSO, objIniFile
	Dim strFilePath, strKey, strLeftString, fileline, strSection
	
	Set objFSO = CreateObject( "Scripting.FileSystemObject" )
	
	ReadIni     = ""
	strFilePath = Trim( myFilePath )
	strSection  = Trim( mySection )
	strKey      = Trim( myKey )
	
	If objFSO.FileExists( strFilePath ) Then
		Set objIniFile = objFSO.OpenTextFile( strFilePath, 1, False )
		Do While objIniFile.AtEndOfStream = False
			fileline = Trim( objIniFile.ReadLine )
			
			' Check if section is found in the current line
			If LCase( fileline ) = "[" & LCase( strSection ) & "]" Then
				fileline = Trim( objIniFile.ReadLine )
				
				' Parse lines until the next section is reached
				Do While Left( fileline, 1 ) <> "["
					' Find position of equal sign in the line
					intEqualPos = InStr( 1, fileline, "=", 1 )
					If intEqualPos > 0 Then
						strLeftString = Trim( Left( fileline, intEqualPos - 1 ) )
						' Check if item is found in the current line
						If LCase( strLeftString ) = LCase( strKey ) Then
							ReadIni = Trim( Mid( fileline, intEqualPos + 1 ) )
							' In case the item exists but value is blank
							If ReadIni = "" Then
								ReadIni = " "
							End If
							' Abort loop when item is found
							Exit Do
						End If
					End If
					
					' Abort if the end of the INI file is reached
					If objIniFile.AtEndOfStream Then Exit Do
					
					' Continue with next line
					fileline = Trim( objIniFile.ReadLine )
				Loop
				Exit Do
			End If
		Loop
		objIniFile.Close
	Else
		ShowMessage(strFilePath & " doesn't exist. Failed to 1: " & myKey & " Exiting...")
		Exit Function
	End If
End Function


' Source: http://www.robvanderwoude.com/vbstech_files_ini.php
Sub WriteIni (myFilePath, mySection, myKey, myValue)
	Dim blnInSection, blnKeyExists, blnSectionExists, blnWritten
	Dim intEqualPos
	Dim objFSO, objNewIni, objOrgIni, wshShell
	Dim strFilePath, strFolderPath, strKey, strLeftString
	Dim fileline, strSection, strTempDir, strTempFile, strValue
	
	strFilePath = Trim( myFilePath )
	strSection  = Trim( mySection )
	strKey      = Trim( myKey )
	strValue    = Trim( myValue )
	
	Set objFSO   = CreateObject( "Scripting.FileSystemObject" )
	Set wshShell = CreateObject( "WScript.Shell" )
	
	strTempDir  = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
	strTempFile = objFSO.BuildPath( strTempDir, objFSO.GetTempName )
	
	Set objOrgIni = objFSO.OpenTextFile( strFilePath, 1, True )
	Set objNewIni = objFSO.CreateTextFile( strTempFile, False, False )
	
	blnInSection     = False
	blnSectionExists = False
	' Check if the specified key already exists
	blnKeyExists     = ( ReadIni( strFilePath, strSection, strKey ) <> "" )
	blnWritten       = False
	
	' Check if path to INI file exists, quit if not
	strFolderPath = Mid( strFilePath, 1, InStrRev( strFilePath, "\" ) )
	
    If Not objFSO.FolderExists ( strFolderPath ) Then
		ShowMessage("Error: WriteIni failed, folder path (" _
		& strFolderPath & ") to ini file " _
		& strFilePath & " not found!")
		Set objOrgIni = Nothing
		Set objNewIni = Nothing
		Set objFSO    = Nothing
		Exit Sub
	End If
	
	While objOrgIni.AtEndOfStream = False
		fileline = Trim( objOrgIni.ReadLine )
		If blnWritten = False Then
			If LCase( fileline ) = "[" & LCase( strSection ) & "]" Then
				blnSectionExists = True
				blnInSection = True
			ElseIf InStr( fileline, "[" ) = 1 Then
				blnInSection = False
			End If
		End If
		
		If blnInSection Then
			If blnKeyExists Then
				intEqualPos = InStr( 1, fileline, "=", vbTextCompare )
				If intEqualPos > 0 Then
					strLeftString = Trim( Left( fileline, intEqualPos - 1 ) )
					If LCase( strLeftString ) = LCase( strKey ) Then
						' Only write the key if the value isn't empty
						' Modification by Johan Pol
						If strValue <> "<DELETE_THIS_VALUE>" Then
							objNewIni.WriteLine strKey & "=" & strValue
						End If
						blnWritten   = True
						blnInSection = False
					End If
				End If
				If Not blnWritten Then
					objNewIni.WriteLine fileline
				End If
			Else
				objNewIni.WriteLine fileline
				' Only write the key if the value isn't empty
				' Modification by Johan Pol
				If strValue <> "<DELETE_THIS_VALUE>" Then
					objNewIni.WriteLine strKey & "=" & strValue
				End If
				blnWritten   = True
				blnInSection = False
			End If
		Else
			objNewIni.WriteLine fileline
		End If
	Wend
	
	If blnSectionExists = False Then ' section doesn't exist
		objNewIni.WriteLine
		objNewIni.WriteLine "[" & strSection & "]"
		' Only write the key if the value isn't empty
		' Modification by Johan Pol
		If strValue <> "<DELETE_THIS_VALUE>" Then
			objNewIni.WriteLine strKey & "=" & strValue
		End If
	End If
	
	objOrgIni.Close
	objNewIni.Close
	
	' Delete old INI file
	objFSO.DeleteFile strFilePath, True
	' Rename new INI file
	objFSO.MoveFile strTempFile, strFilePath
	
	Set objOrgIni = Nothing
	Set objNewIni = Nothing
	Set objFSO    = Nothing
	Set wshShell  = Nothing
End Sub

Sub ShowMessage(msg)
    Msgbox(msg)
End Sub