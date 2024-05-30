''******************************************************************************
' Name:
' verity-health.vbs
'
' Requires:
' mkvdk, mapped drives for collection hosts
'
' Inputs:
' Control File location
'
' Usage:
' cscript verity-health.vbs d:\dev\scripts
'
' Purpose:
' Reads control files, finds Verity collections, 
' gather stats and sends output to screen
'
' Layout:
' Assignment, Sets, Main, Subs, Functions
'
'******************************************************************************

' *******************************************************
' Mapped Drive Assignment
'     Alias = Mounted_Drive:Collection_Directory
' *******************************************************
' Change these server to drive mappings here and verify the server
' names match those in sub WhichMapDrive below
Const bean = "t:\colls"
Const ender = "d:\colls"
Const ForReading=1


Dim oArgs
Dim fso, wso
Dim sMKVDK, sVerityLock, sControlFilePath
Dim dtmDateNow, dtmTimeNow, dtmYYYY, dtmMM, dtmDD, dtmHH, dtmMN, dtmSS
Dim wsoReturn, wsoCommand, sControlFile, dictControlFiles
Dim sLogFolder, sLogFolderPath, sMkvdkLogFolderPath

Set oArgs = Wscript.Arguments
Set fso = CreateObject("Scripting.FileSystemObject")
Set wso = WScript.CreateObject("WScript.Shell")
' 
' Verify these two paths are correct
sMKVDKAbout = "d:\colls\tools\mkvdk-about.cmd"
sLogFolderPath = "d:\colls\tools\logfiles\"

' *******************************************************
' Main
' *******************************************************
' administrative work begins here
Call subBanner()
Call subParseArgs()
Call subCheckControlFolder(oArgs(0))
Call subCheckMKVDKAbout()
Call subCheckTempLogFolderDir(sLogFolderPath)
Call subBuildNames()

Call subBuildFolder(sMkvdkLogFolderPath)
Call subFlightCheck()

' real work begins below
Call subBuildControlDictionary(oArgs(0))
Call subCheckCollections(sCapture)
Call subFinish()

' *******************************************************
' Sub routines
' *******************************************************
Sub subCheckCollections(sTempLogFile)
'
' locate the collecitons and examine the parts directories
' and output from mkvdk -about
'
	Dim strCurrControlFile, strDrive, i,strText
	Dim oFolder, oFile, strControlRead, objTextFile
	Dim objFSO, arrControlLines, arrCollectionLine
	Dim strControlLine, strCollectionFolder

	strCurrControlFile = dictControlFiles.items
	For i = 0 To dictControlFiles.Count -1
		Wscript.echo " "
		wscript.echo "Working " & strCurrControlFile(i)
		strDrive = WhichMapDrive(strCurrControlFile(i))
		if len(strDrive) > 1 then
			wscript.echo "+ Setting local drive " & strDrive

			Set objFSO = CreateObject("Scripting.FileSystemObject")
			strControlRead = oArgs(0) & "\" & strCurrControlFile(i)

			wscript.echo "+ Reading in control file " & strControlRead
			Set objTextFile = objFSO.OpenTextFile(strControlRead, ForReading)
			strText = objTextFile.ReadAll
			objTextFile.Close

			' looking for extra white space
			' another proud VB moment
			' see this sweet hack used for inspiration
			' http://www.microsoft.com/technet/scriptcenter/resources/qanda/may05/hey0520.mspx
			intLength = Len(strText)
			strEnd = Right(strText, 2)
			If strEnd = vbCrLf Then
		    		strText = Left(strText, intLength - 2)
			End If

			wscript.echo "+ Parsing configuration from " & strControlRead						
			arrControlLines = Split(strText, vbCrLf)
		
			For Each strControlLine in arrControlLines
			arrCollectionLine = split(strControlLine, ";")
			strServerName = arrCollectionLine(0)
			strCollectionFolder = arrCollectionLine(1)
			strCollectionPath =  strDrive & "\" & trim(strCollectionFolder)
			strCollectionPartsPath = strDrive & "\" & trim (strCollectionFolder) & "\parts"			

				iDDD = 0
				iMrg = 0
				iDID = 0
				if DoesFolderExist(strCollectionPath) = 0 then
					wscript.echo "+ Working in " & strCollectionPath
					wscript.echo "  " & GetCollectionSize(strCollectionPartsPath) &  " GB"
					wscript.echo"  Server " & strServerName
					Set oFolder = fso.GetFolder(strCollectionPartsPath)
					iTotalFiles = 0
					for each oFile in oFolder.files
						if lcase(right(oFile, 3)) = "did" then
							iDID = iDID + 1
						end if
						if lcase(right(oFile, 3)) = "ddd" then
							iDDD = iDDD + 1
						end if
						if lcase(right(oFile, 3)) = "mrg" then
							iMrg = iMrg + 1
						end if
						iTotalFiles = iTotalFiles + 1
					next
					wscript.echo "  " & iDDD & " ddd files"
					wscript.echo "  " & iMrg & " mrg files"	
					wscript.echo "  " & iDID & " did files"					
					wscript.echo "  " & iTotalFiles & " total files"

					sTempLogFile = sMkvdkLogFolderPath & "\" & arrCollectionLine(1) & ".log"
					wsoCommand = sMKVDKAbout & " " & strCollectionPath & " " &  sTempLogFile
					wscript.echo "  " & "calling mkvdk"
					wsoReturn = wso.Run(wsoCommand,7,true)
					Set objTextFile = objFSO.OpenTextFile(sTempLogFile, ForReading)
					strText = objTextFile.ReadAll
					objTextFile.Close
					wscript.echo "+ Reading MKVDK Results "
					arrMKVDKLines = Split(strText, vbCrLf)
					For Each strMKVDKResult in arrMKVDKLines
						If InStr(strMKVDKResult, "Last Squeeze Date") Then
							wscript.echo "  " & trim(strMKVDKREsult)
						End If
						If InStr(strMKVDKResult, "Number of Documents") Then
							wscript.echo "  " & trim(strMKVDKREsult)
						End If
					Next				
				end if
			Next
		End if

	Next
End Sub

'
' Create a wsh directory structure that contains the names of all
' control files found in the directory given at run time
'
Sub subBuildControlDictionary(strControlFilePath)
	Dim oFolder, oFile
	Set oFolder = fso.GetFolder(strControlFilePath)
	Set dictControlFiles = CreateObject("Scripting.Dictionary")
	
	wscript.echo ""
	wscript.echo "Finding control files"
	i=0
	For Each oFile in oFolder.Files
		if left(oFile.Name, 4) = "ctl_" then
			Wscript.Echo "+ Found " & oFile.Name
			dictControlFiles.add i, oFile.Name
			i = i + 1
		end if
	Next
	if i < 1 then
		Wscript.Echo "Error:"
		Wscript.Echo "====================================="
		Wscript.Echo "Could not find control files in " & oArgs(0)
		wscript.Echo ""
		Wscript.Quit 1
	End if
	if dictcontrolFiles.Count > 4 then
		wscript.Echo ""
		wscript.echo "Configuration Alert"
		Wscript.Echo "====================================="
		Wscript.Echo "Found more than four control files."
		wscript.Echo "This script will continue, but may need to be modified."
		wscript.Echo ""
	end if
End Sub

Sub subParseArgs()
'
' read working directory from command line
'
	if oArgs.Count < 1 then
		Wscript.Echo "Error:"
		Wscript.Echo "====================================="
		Wscript.Echo ""
		wscript.Echo "Requires path to control files"
		wscript.echo "used to handle operations"
		wscript.echo "with Verity collections"
		wscript.quit 1
	End if
	wscript.echo "+ Parsed arguments"
End Sub

Sub subFinish()
	wscript.echo ""
	wscript.echo "Script completed"
	wscript.echo "Start "&  dtmDateNow & " " & dtmTimeNow
	wscript.echo "Stop " & date & " " & Time
End Sub

Sub subFlightCheck()
'
' confirm that we're ready to go by giving user 
' 10 seconds to break out of the script
'
	wscript.echo ""
	wscript.echo "You provided the following arguments"
	for i = 0 to oArgs.Count -1
		wscript.echo "     " & oArgs(i)
	Next
	wscript.echo ""
	wscript.echo "Script will read control files found in "
	wscript.echo oArgs(0)
	wscript.echo "and try to determine their health"
	wscript.echo "====================================="
	wscript.echo "Hit Ctrl+C now to cancel."
	wscript.echo "Script resumes in 5 seconds"
	wscript.echo ""
 	wscript.sleep(5000)
End Sub

Sub subBanner()
	Wscript.echo ""
	Wscript.echo "orb-verity-health.vbs"
	wscript.echo "==================="
	wscript.echo "(c) based on code copyright (c) caseshare 2005"
	Wscript.echo ""
End Sub

Sub subCheckTempLogFolderDir(sFolderPath)
'
' Verify that the 'parent' log directory exists i
' or create it if it does not 
'
If DoesFolderExist(sControlFilePath) = -1 Then
        Wscript.Echo "Warning:"
        Wscript.Echo "====================================="
        Wscript.Echo sFolderPath
        wscript.Echo "Folder does not exist; creating"

	If CreateTempLogFolder(sFolderPath) = -1 Then
	        Wscript.Echo "Error:"
	        Wscript.Echo "====================================="
	        Wscript.Echo sFolderPath & " could not be created"
	        wscript.Echo ""
	        Wscript.Quit 1
	Else
	        wscript.echo "+ Temp log folder created " & sFolderPath
	End If
	
Else
        wscript.echo "+ Control folder exists"
End if

End Sub


Sub subCheckControlFolder(sControlFilePath)
wscript.echo "in CheckControlFolder parm is " & sFolderPath
If DoesFolderExist(sControlFilePath) = -1 Then
	Wscript.Echo "Error:"
	Wscript.Echo "====================================="
	Wscript.Echo oArgs(0)
	wscript.Echo "Folder does not exist"
	wscript.quit 1
Else
	wscript.echo "+ Control folder exists"
End if
End Sub


Sub subCheckMKVDKAbout()
If DoesFileExist(sMKVDKAbout) = -1 Then
	Wscript.Echo "Error:"
	Wscript.Echo "====================================="
	Wscript.Echo sMKVDKAbout & " does not exist."
	wscript.Echo ""
	Wscript.Quit 1
Else
	wscript.echo "+ mkvdkAbout found"
End If
End Sub

Sub subCheckMKVDK()
If DoesFileExist(sMKVDK) = -1 Then
	Wscript.Echo "Error:"
	Wscript.Echo "====================================="
	Wscript.Echo sMKVDK & " does not exist."
	wscript.Echo ""
	Wscript.Quit 1
Else
	wscript.echo "+ mkvdk found"
End If
End Sub

Sub subBuildFolder(sFolderPath)
If CreateTempLogFolder(sFolderPath) = -1 Then
	Wscript.Echo "Error:"
	Wscript.Echo "====================================="
	Wscript.Echo sFolderPath & " could not be created"
	wscript.Echo ""
	Wscript.Quit 1
Else
	wscript.echo "+ Temp log folder created " & sFolderPath
End If
	
End Sub

Sub subBuildNames()
	dtmDateNow = date
	dtmTimeNow = time

	dtmYYYY = DatePart("yyyy", dtmDateNow)
	dtmMM = DatePart("m", dtmDateNow)
	dtmDD = DatePart("d", dtmDateNow)
	dtmHH = DatePart("h", dtmTimeNow)
	dtmMN = DatePart("n", dtmTimeNow)
	dtmSS = DatePart("s", dtmTimeNow)

	If dtmMM < 10 Then
		dtmMM = "0" & dtmMM
	End if
	If dtmHH < 10 Then
		dtmHH = "0" & dtmHH
	End if
	If dtmMN < 10 then
		dtmMN = "0" & dtmMN
	End if
	If dtmDD < 10 Then
		dtmDD = "0" & dtmDD
	End if
	If dtmSS < 10 Then
		dtmSS = "0" & dtmSS
	End If
	sLogFolder = "_orb-" & dtmYYYY & dtmMM & dtmDD & dtmHH & dtmMN & dtmSS
	sMkvdkLogFolderPath = sLogFolderPath & sLogFolder

End Sub


' *******************************************************
' functions
' *******************************************************
function WhichMapDrive(strFileName)
	if lcase(strFileName) = "ctl_ender.txt"  then
		whichMapDrive = ender
	End If
	if lcase(strFileName) = "ctl_bean.txt" then
		whichMapDrive = bean
	End If	
end function

function DoesFileExist(FilePath)
Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	if not fso.FileExists(FilePath) then
		DoesFileExist = -1
	else
		DoesFileExist = 0
	end if
	Set fso = Nothing

end function

function CreateTempLogFolder(sFolderPath)
Dim fso
	set fso = CreateObject("Scripting.FileSystemObject")
	if DoesFolderExist(sFolderPath) < 0 then
		fso.CreateFolder(sFolderPath)
		if DoesFolderExist(sFolderPath) < 0 then
			CreateTempLogFolder = -1
		else
			createTempLogFolder = 0
		end if
	else
		CreateTempLogFolder = 0
	end if
	set fso = nothing

end function

function DoesFolderExist(FolderPath)
Dim fso

	Set fso = CreateObject("Scripting.FileSystemObject")
	If not fso.FolderExists(FolderPath) Then
		DoesFolderExist = -1
   	else
     		DoesFolderExist = 0
	end if
	Set fso = Nothing

End Function

function GetCollectionSize(CollPath)
Dim s, oFolder, oFolderItem, fso, wso

	s = -1
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set wso = CreateObject("Shell.Application")
	Set oFolder = wso.Namespace(CollPath)
	Set oFolderItem = oFolder.Self
	strPath = oFolderItem.Path
	Set oFolder = fso.GetFolder(strPath)
	s = round ((oFolder.size / 1024 / 1024 / 1024 ), 2)
	set oFolder = Nothing
	set fso = Nothing

GetCollectionSize = s
end function

Figure 4: WSH Script
