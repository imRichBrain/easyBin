'#!/   easyBin by RICH Brain v1.0 c2024/dec/31

Set vc=createObject("SAPI.SpVoice")
Set fso=createObject("Scripting.fileSystemObject")
Set obj=createObject("wscript.shell")
Set args=wscript.arguments
bin=Array()

Sub addToBin(val)
	ReDim Preserve bin(uBound(bin)+1)
	bin(uBound(bin))=val
End Sub
Sub clearBin
confirm=msgbox("You Have "&uBound(bin)+1&" files/folders in bin ,"&vbLf&"Do you wanna Delete all these files ?",vbSystemModal+vbYesNo,"Clear Bin ?")
	If confirm=vbYes then
		For Each file in bin
			If fso.fileExists(file) then
				fso.DeleteFile(file)
			ElseIf fso.FolderExists(file) then
				fso.DeleteFolder(file)
			Else
				msgbox "file not exists!"
				wscript.quit
			End If
		Next
		fso.openTextFile(obj.specialFolders("Desktop")&"\bin.txt",2,true).write ""
		vc.speak "the noCycleBin, Deleted "&uBound(bin)+1&" files Successfully!"
		
	End If
End Sub
Sub checkRow(item)
	Set rows=fso.openTextFile(obj.specialFolders("Desktop")&"\bin.txt",1,true)
	Do until rows.atEndOfStream
		If rows.readLine=item then
			vc.speak "You Droped this file or folder before ,"&vbLf&"Please Add Another file or folder to noCycleBin"
			wscript.quit
		End If
	Loop
	rows.close
End Sub
Set files=fso.openTextFile(obj.specialFolders("Desktop")&"\bin.txt",1,true)
Do until files.atEndOfStream
	addToBin(files.ReadLine)

Loop
files.close
If args.count<>0 then
	Set trash=fso.openTextFile(obj.specialFolders("Desktop")&"\bin.txt",8,true)
	For Each arg in args
		checkRow(arg)
		If arg="C:\Users\DAY\OneDrive\Desktop\bin.txt" then
			vc.speak "You can not delete the easyBin database."
			msgbox "you can not delete the easyBin database.",vbSysemModal+vbOkOnly,"Can not delete easyBin DB."
			wscript.quit
		Else
			trash.writeLine arg
			addToBin(arg)
		End If
	Next
	trash.close
	vc.speak "Files added to the easyBin successfully ."
	msgBox "You Added this files to Bin :"&vbLf&vbLf&join(bin,vbLf),vbSystemModal,"File added to bin."
Else
	If uBound(bin)<0 then
		vc.speak "your easyBin is Empty!"
		msgBox "You Have no file in easyBin ,"&vbLf&"Drap a file and drop it on this file to add it to bin"
	Else
		call clearBin
	End If
End If