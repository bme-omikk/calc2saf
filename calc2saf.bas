REM  *****  BASIC  *****

Dim ps, tab, eol As String
Dim dcSep As String
Dim baseFolder As String
Dim sheetName As String
Dim metaRow As Integer
Dim firstDataRow As Integer
Dim lastDataColumn As Integer
Dim bundleColumn As Integer
Dim permissionsColumn As Integer
Dim permissionStringColumn As Integer
Dim descriptionColumn As Integer
Dim primaryColumn As Integer
Dim filenameColumn As String
Dim assetNrColumn As Integer
Dim withoutFileCopy As String
Dim skipCols() As Integer
Dim collectionHandle As String

Function isInArray(searchFor as Integer, arr as Variant)
    Dim i, iStart, iStop
    iStart = LBound(arr)
    iStop = UBound(arr)
    If iStart<=iStop then
        For i = iStart to iStop
            If searchFor = arr(i) then
                isInArray = True
                Exit function
            End if
        Next
    End if
    isInArray = False
End Function

'**
'* https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Strings_(Runtime_Library)
'* 
Function Replace(Source As String, Search As String, NewPart As String)
  Dim Result As String
  Dim StartPos As Long
  Dim CurrentPos As Long
 
  Result = ""
  StartPos = 1
  CurrentPos = 1
 
  If Search = "" Then
    Result = Source
  Else 
    Do While CurrentPos <> 0
      CurrentPos = InStr(StartPos, Source, Search)
      If CurrentPos <> 0 Then
        Result = Result + Mid(Source, StartPos, _
        CurrentPos - StartPos)
        Result = Result + NewPart
        StartPos = CurrentPos + Len(Search)
      Else
        Result = Result + Mid(Source, StartPos, Len(Source))
      End If                ' Position <> 0
    Loop 
  End If 
 
  Replace = Result
End Function

Function l(msg, Optional level)
	Dim iFile As Integer
	Dim lvl As String

	iFile = FreeFile
	
	If IsMissing(level) Then
		lvl = "INFO"
	Else
		lvl = level
	End If
	
	Open baseFolder + "convert.log" For Append As iFile
	Print #iFile, Now & Chr(9) & lvl & Chr(9) & msg
	Close iFile
End Function

Function getNamespace(str)
	Dim parts As Variant
	
	If IsEmpty(str) Or str = "" Then
		getNamespace = ""
	Else
		parts = Split(str, ".")
		getNamespace = parts(0)
	End If
End Function

Function getColumnNrByChar(ch)
	Dim nr1, nr2 As String
	
	If ch = "" Then
		getColumnNrByChar = -1
		Exit Function
	Else
		getColumnNrByChar = 0
	End If
	
	abc = Array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z")
	
	nr1 = Left(ch,1)
	If len(ch) > 1 Then
		nr2 = Right(ch,1)
	Else
		nr2 = ""
	End If

	If nr2 = "" Then
		For i=0 To UBound(abc)
			If abc(i) = nr1 Then
				getColumnNrByChar = i
				Exit Function
			End If
		Next i
	Else
		For i=0 To UBound(abc)
			If abc(i) = nr1 Then
				a = (i+1)*26
			End If
		Next i
		For i=0 To UBound(abc)
			If abc(i) = nr2 Then
				b = i
			End If
		Next i
		getColumnNrByChar = a+b
		Exit Function
	End If	
End Function

Sub Main
	Dim i As Integer
	Dim logline As String
	Dim s As String
	Dim copiedFiles As String
	Dim files As Variant
	
	s = ""

	' path separator
	ps = getPathSeparator
	' tab character
	tab = Chr(9)
	' end of line
	eol = Chr(10)
	
	' parameters
	dcSep            = ThisComponent.Sheets.getByName("Config").getCellRangeByName("B1").String
	baseFolder       = ThisComponent.Sheets.getByName("Config").getCellRangeByName("B2").String
	sheetName        = ThisComponent.Sheets.getByName("Config").getCellRangeByName("B3").String
	metaRow          = CInt(ThisComponent.Sheets.getByName("Config").getCellRangeByName("B4").String)
	firstDataRow     = CInt(ThisComponent.Sheets.getByName("Config").getCellRangeByName("B5").String)
	lastDataColumn   = CInt(getColumnNrByChar(ThisComponent.Sheets.getByName("Config").getCellRangeByName("B6").String))
	skipCols         = Split(ThisComponent.Sheets.getByName("Config").getCellRangeByName("B7").String, ",")
	
	'** https://wiki.lyrasis.org/display/DSDOC6x/Importing+and+Exporting+Items+via+Simple+Archive+Format#ImportingandExportingItemsviaSimpleArchiveFormat-Configuringmetadata_[prefix].xmlforaDifferentSchema
	bundleColumn     = CInt(getColumnNrByChar(ThisComponent.Sheets.getByName("Config").getCellRangeByName("B9").String))
	permissionsColumn= CInt(getColumnNrByChar(ThisComponent.Sheets.getByName("Config").getCellRangeByName("B10").String))
	permissionStringColumn = CInt(getColumnNrByChar(ThisComponent.Sheets.getByName("Config").getCellRangeByName("B11").String))
	descriptionColumn= CInt(getColumnNrByChar(ThisComponent.Sheets.getByName("Config").getCellRangeByName("B12").String))
	primaryColumn    = CInt(getColumnNrByChar(ThisComponent.Sheets.getByName("Config").getCellRangeByName("B13").String))
	filenameColumn   = CInt(getColumnNrByChar(ThisComponent.Sheets.getByName("Config").getCellRangeByName("B14").String))
	assetNrColumn    = CInt(getColumnNrByChar(ThisComponent.Sheets.getByName("Config").getCellRangeByName("B15").String))
	collectionHandle = ThisComponent.Sheets.getByName("Config").getCellRangeByName("B16").String
	withoutFileCopy  = ThisComponent.Sheets.getByName("Config").getCellRangeByName("B18").String
	'**

	l("Converting sheet data to DSpace import package")
	l("**********************************************")
	
	i = firstDataRow
	
	oSheets = ThisComponent.sheets
	oSheet = oSheets.getByName(sheetName)
	
	oCell = oSheet.getCellByPosition(0, i)
	
	do while oCell.String <> ""
		s = oCell.String
		' if the id column contains file names then we use only the file name part

		l(createFolder(baseFolder + s))
		
		copiedFiles = copyFile(s, i, oSheet)
		If InStr(copiedFiles, "Error") Then
			l(copiedFiles, "ERROR")
		Else
			l("Copied files: " & copiedFiles)
		End If
		
		files = Split(copiedFiles)
				
		For j = 0 To UBound(files)
			If files(j) <> "" Then
				Dim perm, ln As String
				
				ln = ""
				
				If oSheet.getCellByPosition(assetNrColumn, i).String<>"" Then
					ln = ln & "-r -s " & oSheet.getCellByPosition(assetNrColumn, i).String & " -f "
					If oSheet.getCellByPosition(filenameColumn, i).String<>"" Then
						ln = ln & oSheet.getCellByPosition(filenameColumn, i).String
					Else
						ln = ln & files(j)
					End If
				Else
					If filenameColumn > -1 Then
						If oSheet.getCellByPosition(filenameColumn, i).String<>"" Then
							ln = ln & oSheet.getCellByPosition(filenameColumn, i).String
						End If
					Else
						ln = ln & files(j)
					End If
				End If
				
				If bundleColumn > -1 Then
					If oSheet.getCellByPosition(bundleColumn, i).String<>"" Then
						ln = ln & tab & "bundle:" & oSheet.getCellByPosition(bundleColumn, i).String
					End If
				End If
				If permissionsColumn > -1 Then
					If oSheet.getCellByPosition(permissionsColumn, i).String<>"" Then
						If permissionStringColumn > -1 Then
							If oSheet.getCellByPosition(permissionStringColumn, i).String<>"" Then
								perm = oSheet.getCellByPosition(permissionStringColumn, i).String
							Else
								perm = "r"
							End If
						End If
						ln = ln & tab & "permissions:-" & perm & " '" & oSheet.getCellByPosition(permissionsColumn, i).String & "'"
					End If
				End If
				If descriptionColumn > -1 Then
					If oSheet.getCellByPosition(descriptionColumn, i).String<>"" Then
						ln = ln & tab & "description:" & oSheet.getCellByPosition(descriptionColumn, i).String
					Else
						ln = ln & tab & "description:" & s
					End If
				End If
				If primaryColumn > -1 Then
					If oSheet.getCellByPosition(primaryColumn, i).String<>"" Then
						ln = ln & tab & "primary:" & oSheet.getCellByPosition(primaryColumn, i).String
					End If
				End If
			End If
		Next j
	
		l(createContentsFile(s, ln))
	
		l(createCollectionsFile(s))
	
		createMetadataContent(s, oSheet, i)
		
		i = i + 1
		oCell = oSheet.getCellByPosition(0, i)
	loop
	Msgbox "Log file has been created: " & Basefolder + ps + "convert.log"
End Sub

Function createMetadataFile(id, schema, content)
	Dim iFile As Integer
	Dim sStr As String
	Dim sHeader As String
	Dim sEnd As String
	Dim sFilename As String
	
	iFile = FreeFile
	
	sFilename = baseFolder + id + "/"
	
	If schema = "dc" Then
		sFilename = sFilename & "dublin_core.xml"
	Else
		sFilename = sFilename & "metadata_" & schema & ".xml"
	End If

	If FileExists(sFilename) Then
		createMetadataFile = "Metadata file " & sFilename & " already exists."
		exit function
	End If

	sHeader = "<?xml version=""1.0"" encoding=""UTF-8""?>" & eol &_
		      "<dublin_core schema=""" & schema & """>" & eol
	sEnd = "</dublin_core>"

	If Not ((LBound(content) = 0) And (UBound(content) = -1)) Then
		sStr = sHeader
			   
	   	For i = 0 To UBound(content)
	   		sStr = sStr & tab & content(i) & eol
		Next
		
		sStr = sStr & sEnd
		
		writeEncodedText(sFilename, sStr, "UTF-8")
	End If
	createMetadataFile = "Metadata file " & sFilename & " for " & id & " has been written."
End Function

'**
'* This function can write text with encodings into a file.
'* The simple file write does not create UTF-8 encoding under Windows
'* instead it encodes them with ISO-8859:(
'* Based on https://forum.openoffice.org/en/forum/viewtopic.php?f=20&t=87895#p412845
Sub writeEncodedText(myPath As String, myText As String, myEncoding As String)
	Dim myTextFile As Object, mySf As Object, myFileStream As Object

	On Error Goto fileKO

	mySf = createUnoService("com.sun.star.ucb.SimpleFileAccess")
	myTextFile = createUnoService("com.sun.star.io.TextOutputStream" )
	myFileStream = mySf.openFileWrite(myPath)
	myTextFile.OutputStream = myFileStream
	myTextFile.Encoding = myEncoding

	myTextFile.writeString(myText)

	myFileStream.closeOutput : myTextFile.closeOutput

	On Error Goto 0
	Exit Sub

	fileKO:
	Resume fileKO2

	fileKO2:
	On Error Resume Next
	msgBox("File write error !", 16)
	myFileStream.closeOutput : myTextFile.closeOutput

	On Error Goto 0
End Sub


Function createDublinCoreString(id, oSheet, row) As String
	Dim res As String
	Dim fullstr As String
	Dim c, dc_i, dcterms_i, local_i, dspace_i, i As Integer
	Dim cellStr As String
	Dim oCell, oDcCell As Variant
	Dim dcParts As Variant
	Dim dc_lines() As String
	Dim dcterms_lines() As String
	Dim local_lines() As String
	Dim dspace_lines() As String
	
	c = 0
	dc_i = 0
	dcterms_i = 0
	local_i = 0
	dspace_i = 0
	
	do while c <= lastDataColumn
		res = ""
		If isInArray(c, skipCols) = False Then
			oCell = oSheet.getCellByPosition(c, row)
			If oCell.Type <> 0 Then ' zero means empty cell
				oDcCell = oSheet.getCellByPosition(c, metaRow)
				If oDcCell.Type <> 0 Then
					cellStr = oCell.String
					dcParts = Split(oDcCell.String, dcSep)
					res = res & "  <dcvalue element=""" & dcParts(1) & """"
					If UBound(dcParts) >= 2 Then
						res = res & " qualifier=""" & dcParts(2) & """>"
					Else
						res = res & " qualifier="""">"
					End If
					fullstr = Replace(cellStr, "&", "&amp;")
					fullstr = Replace(fullstr, "<", "&lt;")
					fullstr = Replace(fullstr, ">", "&gt;")
					res = res & fullstr & "</dcvalue>"
				End If
				If dcParts(0) = "dc" Then
					ReDim Preserve dc_lines(dc_i)
					dc_lines(dc_i) = res
					dc_i = dc_i + 1
				ElseIf dcParts(0) = "dcterms" Then
					ReDim Preserve dcterms_lines(dcterms_i)
					dcterms_lines(dcterms_i) = res
					dcterms_i = dcterms_i + 1
				ElseIf dcParts(0) = "local" Then
					ReDim Preserve local_lines(local_i)
					local_lines(local_i) = res
					local_i = local_i + 1
				ElseIf dcParts(0) = "dspace" Then
					ReDim Preserve dspace_lines(dspace_i)
					dspace_lines(dspace_i) = res
					dspace_i = dspace_i + 1
				End If
			End If
		End If
		c = c + 1
	loop

	createMetadataFile(id, "dc", dc_lines)
	createMetadataFile(id, "dcterms", dcterms_lines)
	createMetadataFile(id, "local", local_lines)
	createMetadataFile(id, "dspace", dspace_lines)

	createDublinCoreString = res
End Function

Function createMetadataContent(id, oSheet, row) As String
	createDublinCoreString(id, oSheet, row)
End Function

Function createContentsFile(id, str) As String
	Dim iFile As Integer
	Dim sStr As String	
	
	writeEncodedText(baseFolder + id + "/contents", str, "UTF-8")
	
	createContentsFile = "Contents file for " + id + " has been written."
End Function

Function createCollectionsFile(id) As String
	Dim iFile As Integer
	Dim sStr As String
	
	sStr = collectionHandle
	
	writeEncodedText(baseFolder + id + "/collections", sStr, "UTF-8")
	
	createCollectionsFile = "Collections file for " + id + " has been written."
End Function

Function createFolder(f) As String
	If Dir(f, 16) = "" Then
		MkDir f
		createFolder = f + " has been created."
	Else createFolder = f + " already exists."
	End If	
End Function

Sub createEmptyFile(fName)
    Dim iNumber As Integer
    
    iNumber = Freefile
    Open fName For Output As #iNumber    
    Print #iNumber, "NEED-OVERWRITE"
    Close #iNumber
    
    iNumber = Freefile
    Open baseFolder & "files-to-copy.lst" For Append As #iNumber    
    Print #iNumber, fName
    Close #iNumber
    iNumber = Freefile
End Sub

Function copyFile(id, row, oSheet) As String
	Dim file As String
	Dim val As String
	
	copyFile = ""
	val = Dir(baseFolder, 0)
	
	On Error Goto Err

	copyFile = ""

	If oSheet.getCellByPosition(assetNrColumn, row).String<>"" Then
		copyFile = oSheet.getCellByPosition(filenameColumn, row).String
		Exit Function
	End If
	
	If withoutFileCopy = "true" Then
		copyFile = oSheet.getCellByPosition(filenameColumn, row).String
		createEmptyFile baseFolder + id + ps + oSheet.getCellByPosition(filenameColumn, row).String
		Exit Function
	End If
	
	Do While (val <> "")
		If val <> "." And val <> ".." Then
			If filenameColumn > -1 Then
				If oSheet.getCellByPosition(filenameColumn, row).String<>"" Then
					val = oSheet.getCellByPosition(filenameColumn, row).String
					FileCopy(baseFolder+val, baseFolder + id + ps + val)
					copyFile = copyFile + val + " "
					val = ""
				End If
			ElseIf InStr(val, id) > 0 Then
				FileCopy(baseFolder+val, baseFolder + id + ps + val)
				copyFile = copyFile + val + " "
				val = Dir()
			Else
				val = Dir()
			End If
		End If
	Loop
	Exit Function
Err:
	'Msgbox "Error during copy file: " & baseFolder+ id + ps + val & " " & Error
	copyFile = copyFile + "Error during copy file: " + baseFolder+ id + ps + val
	Exit Function
End Function
