' this is a script that scan a folder with all the subfolders for the pdfs with higher version than 1.4 
' and it converte it to and old as f*ck version because this server is soo f*ing old that it should be taken for a wolk in the woods 
' now ,we can spoon feed the Dtsearch the  pdf version that he likes
'i used VBS because what else could i use !! 
'USES gostscript 9.5

' yeah ! the comments are in english ! enjoy




'variables setting
objStartFolder = "D:\Netelib_Doc\Docs"    													'the folder for the pdfs files


gs = """C:\gs9.50\bin\gswin32c.exe"""															'the gostscript folder/exe
gs_options = " -sDEVICE=pdfwrite -dCompatibilityLevel=1.3 -dNOPAUSE -dQUIET -dBATCH "			'the options for gostscript , i choosed 1.3 because it is the best supported
ext_vnc = "_version_non_compatible"																'we need to keep a copie of file , because when the shit hit the fun , we will find ourself ready		




filesConverted = 0																				'number of files converted
'give me the time and date
strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
strDateTime = strSafeDate & "-" & strSafeTime

LogFile = objStartFolder&"\convert_files_log_"&strDateTime&".log"			    								'the log file



' we try to elevate script privileges 
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
  WScript.Quit
End If




Const FOR_READING = 1


'the object that manipulate files
Dim objFSO:Set objFSO = CreateObject("Scripting.FileSystemObject")

'we create the log file
Dim objLogFile:Set objLogFile = objFSO.CreateTextFile(logfile, 2, True)

'we get into our folder
Set objFolder = objFSO.GetFolder(objStartFolder)
objLogFile.Write "------- Folder : "&objFolder.Path&"--------"
objLogFile.Writeline
Set colFiles = objFolder.Files
'for eatch file we find we check it and convert it ( and no we  can not get just the pdfs , because it is f*ing vbs)
For Each objFile in colFiles
	check_version objFile
Next

'and we do the same thing for all the subfolders we find
ShowSubfolders objFSO.GetFolder(objStartFolder)










'------------------------------------------------------------
'this is we declare our subs




' get the subfolders in a folder ,and for each file do the samething  
Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
        objLogFile.Write "------- Folder : "& Subfolder.Path &"--------"
        objLogFile.Writeline
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
            check_version objFile
        Next
        ShowSubFolders Subfolder
    Next
End Sub




' here we check the file if it is pdf , that we check the version , if it is so woke for our infrastructure , we beat him into 1.3 
Sub check_version (file)
	if (LCase(objFSO.getextensionname(file.path)) = "pdf") then   'if it is a pdf , then go on..
			Set objTS = objFSO.OpenTextFile(file, FOR_READING)   'open it like a text file the get the first line that it is always the version , like this "%PDF-1.6"
			fullversion = Mid(objTS.Readline,1,8)     
			version_in_string =Replace((Mid(fullversion,6,8)),".",",")	'cut off all the non-sens
			If IsNumeric(version_in_string) Then
				version = CDbl(version_in_string)  ' and with a little bit of magic we have the version as a float
				if (version > 1.4) then														' if it is tooo woke , that is a no no !
									
					objFSO.CopyFile file.Path,file.Path&"_version_non_compatible"			' copy the file to keep a copy
					convert file															' convert the file
					objLogFile.Write "PDF Version :" & version & " Name : "&file.Name              ' log ! log ! always log !!
					objLogFile.Writeline 'skip line
				End if
			else
				objLogFile.Write "****Problem the pdf format ******* Name : "&file.Name      
				objLogFile.Writeline 'skip line				
			End if
			
			

		End if	
End Sub


' the fuction to convert the pdf file using gostscript
Sub convert (file)
	convertArgs = " -sOutputFile="&file.Path&" "&file.Path&"_version_non_compatible"
	Dim objShell
	Set objShell = CreateObject ("WScript.Shell")
	runCmd = gs & gs_options & convertArgs
	objShell.Run runCmd ,0, True
	filesConverted = filesConverted+1
End Sub			

objLogFile.Writeline "+++++++++++++++++++++++++++++++++++++++++++++++++++++++"
objLogFile.Writeline "+++++++++++++++++++++++++++++++++++++++++++++++++++++++"
objLogFile.Write "Number of files converted : " & filesConverted
objLogFile.Close

wScript.Echo "Number of files converted : " & filesConverted

' bilel out
