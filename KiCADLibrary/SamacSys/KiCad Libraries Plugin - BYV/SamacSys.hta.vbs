Dim Records(9,11)
Dim partID
Dim datasheetUrl
Dim ECAD_M
Dim MNA
Dim MPN
Dim ThreeD

Function MySearch()
	Call Search()
End Function

Sub Search()
    Dim Resp, Data
    Dim UrlToGet2
    Dim DownloadedString

	UrlToGet2 = "https://eagle.componentsearchengine.com/alligatorHandler.php?detail=0&searchString=" & document.getElementById("searchText").value & "&offset=0&country=GB&language=en&et=kicad&pv=1.4"
    Set Data = CreateObject("Scripting.Dictionary")
    
    DownloadedString = DownloadString(UrlToGet2)
    If Len(DownloadedString)-Len(Replace(DownloadedString,vbLf,vbNullString)) > 7 Then
	    
	    On Error Resume Next
	    
	    Dim x, respt
	    Dim json
	    Set x = new VbsJson
	    json = DownloadedString
	    Set respt = x.Decode(json)
	
	    Dim i, j
	    j = 0
	    
	    RcrdCount = UBound(respt("parts")) + 1
	    If RcrdCount >= 10 Then
	    	RcrdCount = 10
	    	RcrdSelect = 1
	    Else
	    	RcrdSelect = 10 - RcrdCount +1
	    End If
	    
	    Erase Records	    
	    
	    For i = 0 To RcrdCount-1
	
	        j = j + 1
	
			Records(0,j) = respt("parts")(i)("PartID")
	
			If respt("parts")(i)("ECAD_M") <> "" Then
				Records(1,j) = " Yes"
				Records(2,j) = " Yes"
			Else
				Records(1,j) = "-No-"
				Records(2,j) = "-No-"
			End If
	
			If respt("parts")(i)("Has3D") = "Y" Then
				Records(3,j) = " Yes"
			Else
				Records(3,j) = "-No-"
			End If
			
			If respt("parts")(i)("Datasheet") <> "" Then
				Records(4,j) = " Yes"
			Else
				Records(4,j) = "-No-"
			End If		
	
			desc = respt("parts")(i)("Desc")
			desc = Mid(desc, 1, 37)
			if (len(desc) = 37) Then desc = desc & "..."
			Records(5,j) = desc
			Records(6,j) = respt("parts")(i)("Manuf")
	        Records(7,j) = respt("parts")(i)("PartNo")
	        Records(8,j) = respt("parts")(i)("Datasheet")
	        
	
	    Next
	    
	    SortArray Records, 3
	    	    
	    For j = 1 To 10
	    	document.getElementById("sym" & j).InnerHTML = Records(1,j)
	    	document.getElementById("pcb" & j).InnerHTML = Records(2,j)
	    	document.getElementById("3d" & j).InnerHTML = Records(3,j)
	    	document.getElementById("ds" & j).InnerHTML = Records(4,j)
	    	document.getElementById("desc" & j).InnerHTML = Records(5,j)
	    	document.getElementById("man" & j).InnerHTML = Records(6,j)
	    	document.getElementById("mpn" & j).InnerHTML = Records(7,j)
	    Next
	    
	    if Err.Number <> 0 Then Msgbox("Error: " & Err.Description & VBCRLF & "Source: " & Err.Source)
	    
    	ShowRecord(RcrdSelect)
	    	
    Else
    	Call ClearTable()
    	document.getElementById("iframeWelcome").src = "http://componentsearchengine.com/ExtRef/kicad/homepage.htm"
    	document.getElementById("datasheetLink").style.display = "None"
    	document.getElementById("pinoutLink").style.display = "None"
    	document.getElementById("BuildRequestLink").style.display = "None"
    	document.getElementById("NewBuildRequestLink").style.display = "None"
    	document.getElementById("RequestLink").style.display = "Block"
		document.getElementById("clickToView").style.display = "None"
		document.getElementById("symbolLink").style.display = "None"
		document.getElementById("footprintLink").style.display = "None"
		document.getElementById("3dDelimiter").style.display = "None"
		document.getElementById("3dLink").style.display = "None"
    End If
        
End Sub


Sub ShowRecord(recnum)
    if document.getElementById("sym" & recnum).InnerHTML = "&nbsp;" then exit sub
    for i = 1 to 10
        document.getElementById("row" & i).style.backgroundcolor = "white"
    next
    document.getElementById("row" & recnum).style.backgroundcolor = "#ced2dd" '"gray"
    partID = Records(0, recnum)
    ECAD_M = Records(1, recnum)
    ThreeD = Records(3, recnum)
    MNA = Records(6, recnum)
    MPN = Records(7, recnum)
    datasheetUrl = Records(8, recnum)    
    
    Randomize()
	RndNo = (CStr(Int(1000000 * Rnd())))
    document.getElementById("iframeWelcome").src = "http://componentsearchengine.com/tools/Altium/index.php?partID=" & partID & "&mna=" & MNA & "&mpn=" & MPN & "&t=" & RndNo
    
    document.getElementById("pricingStockLink").style.display = "None"
    document.getElementById("datasheetLink").style.display = "None"
	document.getElementById("pinoutLink").style.display = "None"
	document.getElementById("BuildRequestLink").style.display = "None"
	document.getElementById("NewBuildRequestLink").style.display = "None"
	document.getElementById("RequestLink").style.display = "None"
	
	document.getElementById("clickToView").style.display = "None"
	document.getElementById("symbolLink").style.display = "None"
	document.getElementById("footprintLink").style.display = "None"
	document.getElementById("3dDelimiter").style.display = "None"
	document.getElementById("3dLink").style.display = "None"
	
	If partID = 0 Then
		document.getElementById("NewBuildRequestLink").style.display = "Block"
	ElseIf ECAD_M = "-No-" Then
		document.getElementById("BuildRequestLink").style.display = "Block"
	Else
		document.getElementById("pricingStockLink").style.display = "Block"
    	document.getElementById("pinoutLink").style.display = "Block"
    End If
    
    If datasheetUrl <> "" Then
    	document.getElementById("datasheetLink").style.display = "Block"
    End If
    
    If ECAD_M <> "-No-" Then
    	document.getElementById("clickToView").style.display = "inline-block"
    	document.getElementById("symbolLink").style.display = "inline-block"
    	document.getElementById("footprintLink").style.display = "inline-block"
    End If
    
    If ThreeD = " Yes" Then
    	document.getElementById("3dDelimiter").style.display = "inline-block"
    	document.getElementById("3dLink").style.display = "inline-block"
    End If
    
End Sub

' Clears the HTML results table.
Sub ClearTable()
    For i = 1 to 10
        document.getElementById("row" & i).style.backgroundcolor = "white"
        document.getElementById("sym" & i).InnerHTML = "&nbsp;"
        document.getElementById("pcb" & i).InnerHTML = "&nbsp;"
        document.getElementById("3d" & i).InnerHTML = "&nbsp;"
        document.getElementById("ds" & i).InnerHTML = "&nbsp;"
		document.getElementById("desc" & i).InnerHTML = "&nbsp;"
		document.getElementById("man" & i).InnerHTML = "&nbsp;"
		document.getElementById("mpn" & i).InnerHTML = "&nbsp;"
    Next
End Sub
  
' Starts the login process.  
Sub LoginUserProcess()

    Dim username
    Dim password

    username = document.getElementById("username").value
    password = document.getElementById("password").value

    If Not LoginUser(username,password) Then
        Msgbox "Login failed: Please check your user name and password and try again.", VBCritical
    Else
    	Msgbox "Login successful", vbInformation,"KiCad Libraries Plugin"
    End If

End Sub


Function LoginUser(username,password)

    Dim xHttp: Set xHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xHttp.Open "GET", "https://componentsearchengine.com/ga/auth.txt?", False, username, password
    xHttp.setRequestHeader "Content-Type", "text/html"

    xHttp.Send

    ' Handle HTTP Errors
    If xHttp.Status <> 200 Then
        Set Resp = CreateObject("Scripting.Dictionary")
        Resp.Add "success", "false"
        Resp.Add "error", "HTTP Error: " & xHttp.Status & " " & xHttp.StatusText
    End If

    responseStr = xHttp.responseText
    'Set xHttp = Nothing

    If responseStr = "OK" Then
        LoginUser = True
        strInstalledDir = GetScriptFolder()
        Call WriteIni(strInstalledDir & "install.ini", "General", "username", username)
        Call WriteIni(strInstalledDir & "install.ini", "General", "Password", Password)
        
        'document.getElementById("loginTable").style.display = "None"
        'document.getElementById("mainTable").style.display = "Block"
    Else
        LoginUser = False
    End If

End Function

Sub DownloadFiles()
	Dim bStrm: Set bStrm = createobject("Adodb.Stream")
	Dim ScriptPath : ScriptPath = GetScriptFolder()
    Dim username
    Dim password
    Dim libFound: libFound = False
    Dim overwriteLib: overwriteLib = False
    Dim modFound: modFound = False
    Dim overwriteMod: overwriteMod = False
    Dim dcmFound: dcmFound = False
    Dim overwriteDcm: overwriteDcm = False
    
    If partID = "" Then
    	MsgBox "Please search for a part before clicking 'Add to Library'",,"KiCad Libraries Plugin"
    ElseIf ECAD_M = " Yes" Then
	    strInstalledDir = GetScriptFolder() & "\"
	    username =  ReadIni(strInstalledDir & "install.ini", "General", "username")
	    password =  ReadIni(strInstalledDir & "install.ini", "General", "Password")
	    
	    If Trim(username) = vbNullString Or Trim(password) = vbNullString Then
	    	MsgBox "Please login on 'Login/Settings' tab before clicking 'Add to Library'",,"KiCad Libraries Plugin"
	    Else
	    
		    download3D = ReadIni(strInstalledDir & "install.ini", "General", "Download3D")
		    
		    libDir = ReadIni(strInstalledDir & "install.ini", "General", "LibraryFolder")
		    mdlDir = ReadIni(strInstalledDir & "install.ini", "General", "3dModelFolder")
		    
		    If Trim(libDir) = vbNullString Then
		    	MsgBox "Please select library path on 'Login/Settings' tab before clicking 'Add to Library'",,"KiCad Libraries Plugin"
		    ElseIf Trim(mdlDir) = vbNullString And download3D Then
		    	MsgBox "Please select 3D model path on 'Login/Settings' tab before clicking 'Add to Library'",,"KiCad Libraries Plugin"
		    Else
			    Dim DnldUrl
				DnldUrl = "https://componentsearchengine.com/ga/model.php?partID=" & partID & "&st=10&lt=9&pi=1"
			    
			    Dim xHttp: Set xHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
			    xHttp.Open "GET", DnldUrl, False, username, password
			    xHttp.setRequestHeader "Content-Type", "text/html"
			
			    xHttp.Send
			
			    ' Handle HTTP Errors
			    If xHttp.Status <> 200 Then
			        Set Resp = CreateObject("Scripting.Dictionary")
			        Resp.Add "success", "false"
			        Resp.Add "error", "HTTP Error: " & xHttp.Status & " " & xHttp.StatusText
			    End If
			
			    responseStr = xHttp.responseText
			    
			    If Left(responseStr, 7) = "Error: " Then
			    	MsgBox responseStr, vbCritical
			    	Exit Sub 	
			    End If
		    
			    FilesArray = Split(responseStr,"|**|")
			    
			    dirName = FilesArray(0)
			    libName = FilesArray(1)
			    libFile = FilesArray(2)
			    modName = FilesArray(3)
			    modFile = FilesArray(4)
			    dcmName = FilesArray(5)
			    dcmFile = FilesArray(6)
			    
			    SavePath = ScriptPath + "downloads\" + dirName + "\"
			    
			    libMessages = ReadIni(strInstalledDir & "install.ini", "General", "LibMessages")
			    
			    'Component Library
			    dnldLibFile = Split(FilesArray(2),vbCrLf)
			    oldLibFile = libDir & "\SamacSys_Parts.lib"
			    newLibFile = libDir & "\SamacSys_Temp.lib"
			    
			    Set objFSO = CreateObject("Scripting.FileSystemObject")
			
				Set olf = objFSO.OpenTextFile(oldLibFile)
				Set nlf = objFSO.CreateTextFile(newLibFile)
				
				For i = 0 To UBound(dnldLibFile)
					dnldLibLn = dnldLibFile(i)
					If Mid(dnldLibLn, 1, 21) = "#SamacSys ECAD Model " Then
						dnldPrtLn = dnldLibLn
						oldLibLn = olf.ReadLine
						Do Until oldLibLn = "#End Library"
							nlf.WriteLine oldLibLn
							If oldLibLn = dnldPrtLn Then
								libFound = True		
								Ans = MsgBox("Symbol '" & libName & "' already exists in library SamacSys_Parts.lib'" & vbCrLf & vbCrLf & "OK to overwrite?",vbYesNo,"KiCad Libraries Plugin")
								If Ans = vbYes Then
									overwriteLib = True
									'Write the downloaded component lib to the temp file
									i = i + 1
									dnldLibLn = dnldLibFile(i)
									Do Until dnldLibLn = "ENDDEF"
										nlf.WriteLine dnldLibLn
										i = i + 1
										dnldLibLn = dnldLibFile(i)	
									Loop
									nlf.WriteLine "ENDDEF"
									'Read old lib to "ENDDEF" to skip
									Do Until oldLibLn = "ENDDEF"
										oldLibLn = olf.ReadLine		
									Loop
								End If
							End If
							oldLibLn = olf.ReadLine
						Loop
					End If	
					If Not libFound And dnldPrtLn <> vbNullString Then
						nlf.WriteLine dnldLibLn
					End If
				Next
		
		        If overwriteLib Then
		            nlf.WriteLine "#End Library"
		        End If
		
				olf.Close
				nlf.Close
				
		        If Not libFound Or overwriteLib Then
		        	objFSO.DeleteFile oldLibFile
		            objFSO.MoveFile newLibFile, oldLibFile
		        Else
		        	objFSO.DeleteFile newLibFile
		        End If
		
		        If Not libFound And libMessages Then
		            MsgBox "Symbol " & libName & " has been added to SamacSys_Parts.lib", vbOK, "KiCad Libraries Plugin"
		        End If
		        
			    'Footprint Library
			    dnldModFile = Split(FilesArray(4),vbCrLf)
			    oldModFile = libDir & "\SamacSys_Parts.mod"
			    newModFile = libDir & "\SamacSys_Temp.mod"
			
				Set omf = objFSO.OpenTextFile(oldModFile)
				Set nmf = objFSO.CreateTextFile(newModFile)
				
				For i = 0 To UBound(dnldModFile)
					dnldModLn = dnldModFile(i)
					If Mid(dnldModLn, 1, 8) = "$MODULE " Then
						dnldPcbLn = Mid(dnldModLn, 9)
						oldModLn = omf.ReadLine
						Do Until oldModLn = "$EndINDEX"
							nmf.WriteLine oldModLn
							If oldModLn = dnldPcbLn Then
								modFound = True		
								Ans = MsgBox("Footprint '" & modName & "' already exists in library SamacSys_Parts.mod'" & vbCrLf & vbCrLf & "OK to overwrite?",vbYesNo,"KiCad Libraries Plugin")
								If Ans = vbYes Then
									overwriteMod = True
								End If
							End If
							oldModLn = omf.ReadLine
						Loop
						If Not modFound Then
							dnldModLn = dnldModFile(i)
							nmf.WriteLine dnldPcbLn
							nmf.WriteLine "$EndINDEX"
							Do Until dnldModLn = "$EndLIBRARY"
								nmf.WriteLine dnldModLn
								i = i + 1
								dnldModLn = dnldModFile(i)	
							Loop
							Do Until omf.AtEndOfStream
								oldModLn = omf.ReadLine
								nmf.WriteLine oldModLn	
							Loop
							Exit For									
						End If
					End If
				Next		
						
						
		        If overwriteMod And dnldPcbLn <> vbNullString Then
		        	i = 0
		        	dnldModLn = dnldModFile(i)
		        	Do Until dnldModLn = "$EndINDEX"
		        		i = i + 1
		        		dnldModLn = dnldModFile(i)
		        	Loop       	
		            'Write the new footprint
		            Do Until dnldModLn = "$EndMODULE " & dnldPcbLn
		                nmf.WriteLine dnldModLn
						i = i + 1
						dnldModLn = dnldModFile(i)
		            Loop
		            nmf.WriteLine "$EndMODULE " & dnldPcbLn
		            'Read/Write the old footprint library and skip the "$EndMODULE " & dnldPcbLn section
		    		Do Until omf.AtEndOfStream
						oldModLn = omf.ReadLine
						If oldModLn = "$MODULE " & dnldPcbLn Then
							Do While oldModLn <> "$EndMODULE " & dnldPcbLn
								oldModLn = omf.ReadLine
							Loop
							oldModLn = omf.ReadLine
						End If
						nmf.WriteLine oldModLn	
					Loop
		        End If				
					
				omf.Close
				nmf.Close
		
		        If Not modFound Or overwriteMod Then
		            objFSO.DeleteFile oldModFile
		            objFSO.MoveFile newModFile, oldModFile
		        Else
		        	objFSO.DeleteFile newModFile
		        End If
					
		        If Not modFound And libMessages Then
		            MsgBox "Footprint " & modName & " has been added to SamacSys_Parts.mod", vbOK, "KiCad Libraries Plugin"
		        End If
		        
		        
		        'Documentation File
		        
		        dnldDocLn = vbNullString
			    dnldDcmFile = Split(FilesArray(6),vbCrLf)
			    oldDcmFile = libDir & "\SamacSys_Parts.dcm"
			    newDcmFile = libDir & "\SamacSys_Temp.dcm"
			    
				Set odf = objFSO.OpenTextFile(oldDcmFile)
				Set ndf = objFSO.CreateTextFile(newDcmFile)	            
		        
				For i = 0 To UBound(dnldDcmFile)
					dnldDcmLn = dnldDcmFile(i)
					If Mid(dnldDcmLn, 1, 5) = "$CMP " Then
						dnldDocLn = dnldDcmLn
						oldDcmLn = odf.ReadLine
						Do Until oldDcmLn = "#End Doc Library"
							ndf.WriteLine oldDcmLn
							If oldDcmLn = dnldDocLn Then
								dcmFound = True		
								Ans = MsgBox("Component '" & dcmName & "' already exists in library SamacSys_Parts.dcm'" & vbCrLf & vbCrLf & "OK to overwrite?",vbYesNo,"KiCad Libraries Plugin")
								If Ans = vbYes Then
									overwriteDcm = True
									'Write the downloaded doc entry to the temp file
									i = i + 1
									dnldDcmLn = dnldDcmFile(i)
									Do Until dnldDcmLn = "$ENDCMP"
										ndf.WriteLine dnldDcmLn
										i = i + 1
										dnldDcmLn = dnldDcmFile(i)
									Loop
									ndf.WriteLine "$ENDCMP"
				                    'Read old doc entry to "$ENDCMP" to skip
				                    Do Until oldDcmLn = "$ENDCMP"
				                        oldDcmLn = odf.ReadLine
				                    Loop	
								End If
							End If			
							oldDcmLn = odf.ReadLine
						Loop
					End If
					If Not dcmFound And dnldDocLn <> vbNullString Then
						If dnldDcmLn <> vbNullString Then ndf.WriteLine dnldDcmLn									
					End If
				Next
				
		        If overwriteDcm Then
		            ndf.WriteLine "#End Doc Library"
		        End If
		        
				odf.Close
				ndf.Close
		
		        If Not dcmFound Or overwriteDcm Then
		            objFSO.DeleteFile oldDcmFile
		            objFSO.MoveFile newDcmFile, oldDcmFile
		        Else
		        	objFSO.DeleteFile newDcmFile
		        End If               
		        		
		        If Not dcmFound And libMessages Then
		            MsgBox "Component " & dcmName & " has been added to SamacSys_Parts.dcm", vbOK, "KiCad Libraries Plugin"
		        End If
		        
		        Set objFSO = Nothing		
			  	
			  	If download3D And ThreeD = " Yes" Then
			    	xHttp.Open "GET", "https://componentsearchengine.com/ga/model.php?partID=" & partID & "&step=1", False, username, password'
			    	xHttp.send
			
				    ' Handle HTTP Errors
				    If xHttp.Status <> 200 Then
				        Set Resp = createobject("Scripting.Dictionary")
				        Resp.Add "success", "false"
				        Resp.Add "error", "HTTP Error: " & xHttp.Status & " " & xHttp.StatusText
				    End If
				  	
				   	with bStrm
				    	.Type = 1 '//binary
				      	.open
				      	.write xHttp.responseBody
				      	.savetofile mdlDir + "\" + dcmName + ".step", 2 '//overwrite
				      	.Close
			 		end With
			 		
			    	xHttp.Open "GET", "https://componentsearchengine.com/ga/model.php?partID=" & partID & "&vrml=1", False, username, password'
			    	xHttp.send
			
				    ' Handle HTTP Errors
				    If xHttp.Status <> 200 Then
				        Set Resp = createobject("Scripting.Dictionary")
				        Resp.Add "success", "false"
				        Resp.Add "error", "HTTP Error: " & xHttp.Status & " " & xHttp.StatusText
				    End If
				  	
				   	with bStrm
				    	.Type = 1 '//binary
				      	.open
				      	.write xHttp.responseBody
				      	.savetofile mdlDir + "\" + dcmName + ".wrl", 2 '//overwrite
				      	.Close
			 		end With	 		
			 			
			  	End If
			  	
			  	Set xHttp = Nothing
			  	
			  	'Dim filesys: Set filesys = CreateObject("Scripting.FileSystemObject")
			  	'filesys.CopyFile ScriptPath + "Downloads\Readme.html", ScriptPath + "downloads\" + dirName + "\"
			  	
			  	'MsgBox "Part " + dirName + " has been downloaded to " + SavePath + "...",,"KiCad Libraries Plugin"
			  	
		'	  	openDnldsFdr = ReadIni(strInstalledDir & "install.ini", "General", "OpenDnldsFolder")
			  	
		'	  	If openDnldsFdr Then
		'		  	Set objShell = CreateObject("Wscript.Shell")
		'			objShell.Run "explorer.exe /e," & SavePath
		'		End If
			  	
			  	showInstructions = ReadIni(strInstalledDir & "install.ini", "General", "Instructions")
			  	
			  	If showInstructions Then
				  	Set WShell = CreateObject("WScript.Shell")
				    WShell.Run "http://www.samacsys.com/kicad-libraries/", 1, false
			  	End If
			  	
			  	closePlugin = ReadIni(strInstalledDir & "install.ini", "General", "ClosePlugin")
			  	
			  	If closePlugin Then
			  		Call ExitHTA
			  	End If
			End If
		End If
	Else
		Msgbox "Please click " & Chr(34) & "Build or Free Request" & Chr(34), vbInformation,"KiCad Libraries Plugin"  	
	End If
End Sub

Function GetScriptFolder()
	'Dim WshShell: Set WshShell = CreateObject("Wscript.Shell")
    'GetScriptFolder = WshShell.SpecialFolders("MyDocuments") & "\SamacSys\"
	Dim WshShell, strCurDir
	Set WshShell = CreateObject("WScript.Shell")
	GetScriptFolder    = WshShell.CurrentDirectory & "\"
End Function

Function DownloadString(url)
    Dim xHttp: Set xHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    xHttp.Open "GET", url, False
    xHttp.setRequestHeader "Content-Type", "application/json"

    xHttp.Send

    ' Handle HTTP Errors
    If xHttp.Status <> 200 Then
        Set Resp = CreateObject("Scripting.Dictionary")
        Resp.Add "success", "false"
        Resp.Add "error", "HTTP Error: " & xHttp.Status & " " & xHttp.StatusText
    End If

    DownloadString = xHttp.responseText
    Set xHttp = Nothing
End Function

Function UpdateSettings()
	strInstalledDir = GetScriptFolder()
	Call WriteIni(strInstalledDir & "install.ini", "General", "LibMessages", document.getElementById("chk_LibMessages").checked)
	Call WriteIni(strInstalledDir & "install.ini", "General", "Instructions", document.getElementById("chk_ShowInstruction").checked)
	Call WriteIni(strInstalledDir & "install.ini", "General", "ClosePlugin", document.getElementById("chk_ClosePlugin").checked)	
	Call WriteIni(strInstalledDir & "install.ini", "General", "Download3D", document.getElementById("chk_Download3D").checked)
	Call WriteIni(strInstalledDir & "install.ini", "General", "LibraryFolder", document.getElementById("library").value)
	Call WriteIni(strInstalledDir & "install.ini", "General", "3dModelFolder", document.getElementById("model_folder").value)
End Function
Function UrlEncode(url)
    Dim c
    Dim encoded
    While Len(url) > 0
        c = Left(url, 1)
        url = Mid(url, 2, Len(url) - 1)
        If InStr("[ABCDEFGHIJKLMNOPQRSTUVWXYZZa-z0123456789._~-]", c) <> 0 Then
            encoded = encoded & c
        ElseIf c = " " Then
            encoded = encoded & "+"
        Else
            encoded = encoded & "%" & Right("0" & Hex(Asc(c)), 2)
        End If
    Wend
    UrlEncode = encoded
End Function
Function openDatasheetUrl()
	Set WShell = CreateObject("WScript.Shell")
	WShell.Run datasheetUrl, 1, false
End Function

Function openPinoutUrl()
	Set WShell = CreateObject("WScript.Shell")
	WShell.Run "https://componentsearchengine.com/ga/partCreator.php?partID=" & partID, 1, false
End Function
Function openBuildRequestUrl()
	Set WShell = CreateObject("WScript.Shell")
	WShell.Run "http://componentsearchengine.com/partRequest.html?partID=" & partID, 1, false
End Function
Function openNewBuildRequestUrl()
	Set WShell = CreateObject("WScript.Shell")
	WShell.Run "http://componentsearchengine.com/entry_u.php?mna=" & UrlEncode(MNA) & "&mpn=" & UrlEncode(MPN), 1, false
End Function
Function openRequestUrl()
	Set WShell = CreateObject("WScript.Shell")
	WShell.Run "https://componentsearchengine.com/ga/newPart.php", 1, false
End Function
Function openHelpUrl()
	Set WShell = CreateObject("WScript.Shell")
	WShell.Run "http://www.samacsys.com/kicad-libraries/", 1, false
End Function
Function openLoginDlg()
	document.getElementById("mainTable").style.display = "None"
	document.getElementById("loginTable").style.display = "Block"
End Function
Function openSymbolUrl()
	Set WShell = CreateObject("WScript.Shell")
	WShell.Run "http://componentsearchengine.com/common/footprintPreview.php?partID=" & partID & "&target=symbol", 1, false
End Function
Function openFootprintUrl()
	Set WShell = CreateObject("WScript.Shell")
	WShell.Run "http://componentsearchengine.com/common/footprintPreview.php?partID=" & partID, 1, false
End Function
Function open3dUrl()
	Set WShell = CreateObject("WScript.Shell")
	WShell.Run "http://componentsearchengine.com/viewer/3D.php?partID=" & partID, 1, false
End Function
Function openPricingStockUrl()
	Set WShell = CreateObject("WScript.Shell")
	WShell.Run "https://componentsearchengine.com/detail.html?searchString=" & UrlEncode(MPN) & "&manuf=" & UrlEncode(MNA) & "&country=GB&language=en&source=63", 1, false
End Function
Function libraryBrowse()
    fldr = PickFolder("")
    If fldr <> vbNullString Then
    	document.getElementById("library").value = fldr
    	document.getElementById("dnldsFldr").value = "KiCad Library: " & fldr & "\SamacSys_Parts.lib"
    	Set objFSO = CreateObject("Scripting.FileSystemObject")
    	If Not objFSO.FileExists(fldr & "\SamacSys_Parts.lib") Then
    		Set nlf = objFSO.CreateTextFile(fldr & "\SamacSys_Parts.lib")
			nlf.WriteLine "EESchema-LIBRARY Version 2.3"
			nlf.WriteLine "#encoding utf-8"
			nlf.WriteLine "#SamacSys ECAD Model NE555DR"
			nlf.WriteLine "#/6648/2/2.28/8/3/Integrated Circuit"
			nlf.WriteLine "DEF NE555DR IC 0 30 Y Y 1 F N"
			nlf.WriteLine "F0 " & Chr(34) & "IC" & Chr(34) & " 1000 700 50 H V L CNN"
			nlf.WriteLine "F1 " & Chr(34) & "NE555DR" & Chr(34) & " 1000 600 50 H V L CNN"
			nlf.WriteLine "F2 " & Chr(34) & "SOIC127P600X175-8N" & Chr(34) & " 1000 500 50 H I L CNN"
			nlf.WriteLine "F3 " & Chr(34) & "http://www.ti.com/lit/ds/symlink/ne555.pdf" & Chr(34) & " 1000 400 50 H I L CNN"
			nlf.WriteLine "F4 " & Chr(34) & "Single Precision Timer" & Chr(34) & " 1000 300 50 H I L CNN " & Chr(34) & "Description" & Chr(34)
			nlf.WriteLine "F5 " & Chr(34) & "1.75" & Chr(34) & " 1000 200 50 H I L CNN " & Chr(34) & "Height" & Chr(34)
			nlf.WriteLine "F6 " & Chr(34) & "Texas Instruments" & Chr(34) & " 1000 100 50 H I L CNN " & Chr(34) & "Manufacturer_Name" & Chr(34)
			nlf.WriteLine "F7 " & Chr(34) & "NE555DR" & Chr(34) & " 1000 0 50 H I L CNN " & Chr(34) & "Manufacturer_Part_Number" & Chr(34)
			nlf.WriteLine "F8 " & Chr(34) & "1218414" & Chr(34) & " 1000 -100 50 H I L CNN " & Chr(34) & "RS Part Number" & Chr(34)
			nlf.WriteLine "F9 " & Chr(34) & "http://uk.rs-online.com/web/p/products/1218414" & Chr(34) & " 1000 -200 50 H I L CNN " & Chr(34) & "RS Price/Stock" & Chr(34)
			nlf.WriteLine "F10 " & Chr(34) & "NE555DR" & Chr(34) & " 1000 -300 50 H I L CNN " & Chr(34) & "Arrow Part Number" &Chr(34)
			nlf.WriteLine "F11 " & Chr(34) & "https://www.arrow.com/en/products/ne555dr/texas-instruments" & Chr(34) & " 1000 -400 50 H I L CNN " & Chr(34) & "Arrow Price/Stock" & Chr(34)
			nlf.WriteLine "DRAW"
			nlf.WriteLine "X GND 1 600 -800 200 U 50 50 0 0 B"
			nlf.WriteLine "X TRIG 2 0 0 200 R 50 50 0 0 B"
			nlf.WriteLine "X OUT 3 0 -200 200 R 50 50 0 0 B"
			nlf.WriteLine "X RESET 4 500 600 200 D 50 50 0 0 B"
			nlf.WriteLine "X CONT 5 1200 -300 200 L 50 50 0 0 B"
			nlf.WriteLine "X THRES 6 1200 -200 200 L 50 50 0 0 B"
			nlf.WriteLine "X DISCH 7 1200 0 200 L 50 50 0 0 B"
			nlf.WriteLine "X VCC 8 700 600 200 D 50 50 0 0 B"
			nlf.WriteLine "P 5 0 1 6 200 400 1000 400 1000 -600 200 -600 200 400 N"
			nlf.WriteLine "T 0 600 -100 50 0 0 1 " & Chr(34) & "555" & Chr(34) & " Normal 0 C C"
			nlf.WriteLine "ENDDRAW"
			nlf.WriteLine "ENDDEF"
			nlf.WriteLine "#"
			nlf.WriteLine "#End Library"
    		nlf.Close
    	End If
    
    	If Not objFSO.FileExists(fldr & "\SamacSys_Parts.mod") Then
    		Set nmf = objFSO.CreateTextFile(fldr & "\SamacSys_Parts.mod")
			nmf.WriteLine "PCBNEW-LibModule-V1"
			nmf.WriteLine "# encoding utf-8"
			nmf.WriteLine "Units mm"
			nmf.WriteLine "$INDEX"
			nmf.WriteLine "SOIC127P600X175-8N"
			nmf.WriteLine "$EndINDEX"
			nmf.WriteLine "$MODULE SOIC127P600X175-8N"
			nmf.WriteLine "Po 0 0 0 15 5b50713e 00000000 ~~"
			nmf.WriteLine "Li SOIC127P600X175-8N"
			nmf.WriteLine "Cd D (R-PDSO-G8)"
			nmf.WriteLine "Kw Integrated Circuit"
			nmf.WriteLine "Sc 0"
			nmf.WriteLine "At SMD"
			nmf.WriteLine "AR"
			nmf.WriteLine "Op 0 0 0"
			nmf.WriteLine "T0 0 0 1.27 1.27 0 0.254 N V 21 N " & Chr(34) & "IC**" & Chr(34)
			nmf.WriteLine "T1 0 0 1.27 1.27 0 0.254 N I 21 N " & Chr(34) & "SOIC127P600X175-8N" & Chr(34)
			nmf.WriteLine "DS -3.725 -2.75 3.725 -2.75 0.05 24"
			nmf.WriteLine "DS 3.725 -2.75 3.725 2.75 0.05 24"
			nmf.WriteLine "DS 3.725 2.75 -3.725 2.75 0.05 24"
			nmf.WriteLine "DS -3.725 2.75 -3.725 -2.75 0.05 24"
			nmf.WriteLine "DS -1.95 -2.45 1.95 -2.45 0.1 24"
			nmf.WriteLine "DS 1.95 -2.45 1.95 2.45 0.1 24"
			nmf.WriteLine "DS 1.95 2.45 -1.95 2.45 0.1 24"
			nmf.WriteLine "DS -1.95 2.45 -1.95 -2.45 0.1 24"
			nmf.WriteLine "DS -1.95 -1.18 -0.68 -2.45 0.1 24"
			nmf.WriteLine "DS -1.6 -2.45 1.6 -2.45 0.2 21"
			nmf.WriteLine "DS 1.6 -2.45 1.6 2.45 0.2 21"
			nmf.WriteLine "DS 1.6 2.45 -1.6 2.45 0.2 21"
			nmf.WriteLine "DS -1.6 2.45 -1.6 -2.45 0.2 21"
			nmf.WriteLine "DS -3.475 -2.58 -1.95 -2.58 0.2 21"
			nmf.WriteLine "$PAD"
			nmf.WriteLine "Po -2.712 -1.905"
			nmf.WriteLine "Sh " & Chr(34) & "1" & Chr(34) & " R 0.65 1.525 0 0 900"
			nmf.WriteLine "At SMD N 00888000"
			nmf.WriteLine "Ne 0 " & Chr(34) & Chr(34)
			nmf.WriteLine "$EndPAD"
			nmf.WriteLine "$PAD"
			nmf.WriteLine "Po -2.712 -0.635"
			nmf.WriteLine "Sh " & Chr(34) & "2" & Chr(34) & " R 0.65 1.525 0 0 900"
			nmf.WriteLine "At SMD N 00888000"
			nmf.WriteLine "Ne 0 " & Chr(34) & Chr(34)
			nmf.WriteLine "$EndPAD"
			nmf.WriteLine "$PAD"
			nmf.WriteLine "Po -2.712 0.635"
			nmf.WriteLine "Sh " & Chr(34) & "3" & Chr(34) & " R 0.65 1.525 0 0 900"
			nmf.WriteLine "At SMD N 00888000"
			nmf.WriteLine "Ne 0 " & Chr(34) & Chr(34)
			nmf.WriteLine "$EndPAD"
			nmf.WriteLine "$PAD"
			nmf.WriteLine "Po -2.712 1.905"
			nmf.WriteLine "Sh " & Chr(34) & "4" & Chr(34) & " R 0.65 1.525 0 0 900"
			nmf.WriteLine "At SMD N 00888000"
			nmf.WriteLine "Ne 0 " & Chr(34) & Chr(34)
			nmf.WriteLine "$EndPAD"
			nmf.WriteLine "$PAD"
			nmf.WriteLine "Po 2.712 1.905"
			nmf.WriteLine "Sh " & Chr(34) & "5" & Chr(34) & " R 0.65 1.525 0 0 900"
			nmf.WriteLine "At SMD N 00888000"
			nmf.WriteLine "Ne 0 " & Chr(34) & Chr(34)
			nmf.WriteLine "$EndPAD"
			nmf.WriteLine "$PAD"
			nmf.WriteLine "Po 2.712 0.635"
			nmf.WriteLine "Sh " & Chr(34) & "6" & Chr(34) & " R 0.65 1.525 0 0 900"
			nmf.WriteLine "At SMD N 00888000"
			nmf.WriteLine "Ne 0 " & Chr(34) & Chr(34)
			nmf.WriteLine "$EndPAD"
			nmf.WriteLine "$PAD"
			nmf.WriteLine "Po 2.712 -0.635"
			nmf.WriteLine "Sh " & Chr(34) & "7" & Chr(34) & " R 0.65 1.525 0 0 900"
			nmf.WriteLine "At SMD N 00888000"
			nmf.WriteLine "Ne 0 " & Chr(34) & Chr(34)
			nmf.WriteLine "$EndPAD"
			nmf.WriteLine "$PAD"
			nmf.WriteLine "Po 2.712 -1.905"
			nmf.WriteLine "Sh " & Chr(34) & "8" & Chr(34) & " R 0.65 1.525 0 0 900"
			nmf.WriteLine "At SMD N 00888000"
			nmf.WriteLine "Ne 0 " & Chr(34) & Chr(34)
			nmf.WriteLine "$EndPAD"
			nmf.WriteLine "$EndMODULE SOIC127P600X175-8N"
			nmf.WriteLine "$EndLIBRARY"
			nmf.Close
		End If

    	If Not objFSO.FileExists(fldr & "\SamacSys_Parts.dcm") Then
    		Set ndf = objFSO.CreateTextFile(fldr & "\SamacSys_Parts.dcm")
			ndf.WriteLine "EESchema-DOCLIB  Version 2.0"
			ndf.WriteLine "#"
			ndf.WriteLine "$CMP NE555DR"
			ndf.WriteLine "D Single Precision Timer"
			ndf.WriteLine "K"
			ndf.WriteLine "F http://www.ti.com/lit/ds/symlink/ne555.pdf"
			ndf.WriteLine "$ENDCMP"
			ndf.WriteLine "#"
			ndf.WriteLine "#End Doc Library"
			ndf.Close
		End If
    	Set objFSO = Nothing
    End If
    Call UpdateSettings()
End Function
Function modelFldrBrowse()
    fldr = PickFolder("")
    If fldr <> vbNullString Then document.getElementById("model_folder").value = fldr
    Call UpdateSettings()
End Function
Function SortArray(DArray(), Element)
    Dim gap, doneflag, SwapArray()
    Dim Index
    ReDim SwapArray(2, UBound(DArray, 1), UBound(DArray, 2))
    'Gap is half the records
    gap = Int(UBound(DArray, 2) / 2)
    Do While gap >= 1
        Do
            doneflag = 1
            For Index = 0 To (UBound(DArray, 2) - (gap + 1))
                'Compare 1st 1/2 to 2nd 1/2
                If DArray(Element, Index) > DArray(Element, (Index + gap)) Then
                    For acol = 0 To (UBound(SwapArray, 2) - 1)
                        'Swap Values if 1st > 2nd
                        SwapArray(0, acol, Index) = DArray(acol, Index)
                        SwapArray(1, acol, Index) = DArray(acol, Index + gap)
                    Next
                    For acol = 0 To (UBound(SwapArray, 2) - 1)
                        'Swap Values if 1st > 2nd
                        DArray(acol, Index) = SwapArray(1, acol, Index)
                        DArray(acol, Index + gap) = SwapArray(0, acol, Index)
                    Next
                    CNT = CNT + 1
                    doneflag = 0
                End If
            Next
        Loop Until doneflag = 1
        gap = Int(gap / 2)
    Loop
End Function

Function PickFolder(strStartDir)
Dim SA, F
Set SA = CreateObject("Shell.Application")
Set F = SA.BrowseForFolder(0, "Choose a folder", 0, strStartDir)
If (Not F Is Nothing) Then
  PickFolder = F.Items.Item.path
End If
Set F = Nothing
Set SA = Nothing
End Function 
Function printf(txt)
WScript.StdOut.WriteLine txt
End Function

Sub ExitHTA
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'mshta.exe'")
For Each objProcess in colProcessList
objProcess.Terminate()
Next
End Sub