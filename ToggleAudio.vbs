' Set the audio device names to use the names set
'edited by brule 8-29-15
Const ALL_USERS_DESKTOP = &H1a&
Const USER_DESKTOP = &H1a&
Const nircmd = "C:\Windows\nircmd.exe"
Const Device1 = "Headphones"
Const Device1Name = "Headphones"
Const Device2 = "Speakers"
Const Device2Name = "Speakers"

Set ws = CreateObject("Wscript.Shell")
Set objShell = CreateObject("Shell.Application") 
Set objFolder = objShell.Namespace(USER_DESKTOP)
Set objFolderItem = objFolder.ParseName(Device2Name +".lnk")




if isNull(objFolderItem) or IsEmpty(objFolderItem) or (objFolderItem is Nothing) then
    Set objFolderItem = objFolder.ParseName(Device1Name + ".lnk")
    if isNull(objFolderItem) or IsEmpty(objFolderItem) or (objFolderItem is Nothing) then       
        ' Creates toggle shortcut on desktop n shit
        Set oMyShortcut = ws.CreateShortcut(objFolder.Self.Path + "\"+Device1Name+".lnk")
        oMyShortcut.WindowStyle = 0
        OMyShortcut.TargetPath = WScript.ScriptFullName
        oMyShortCut.Hotkey = "ALT+CTRL+S"
        oMyShortcut.IconLocation = "C:\Windows\System32\mmres.dll, 0"
        oMyShortCut.Save

        ws.run nircmd + " setdefaultsounddevice """+Device1+""" 0", 0
        ws.run nircmd + " setdefaultsounddevice """+Device1+""" 1", 0
        ws.run nircmd + " setdefaultsounddevice """+Device1+""" 2", 0
        msgbox "Desktop link created for """+Device1+""". "+Device1+" set as default!", 0, "Error"
    else
       ' Speakers were set, make headphones drop the bass
        Set objShellLink = objFolderItem.GetLink
        objShellLink.SetIconLocation "C:\Windows\System32\mmres.dll", 2
        objShellLink.Save()
        objFolderItem.Name = Device2Name

        ws.run nircmd + " setdefaultsounddevice """+Device2+""" 0", 0
        ws.run nircmd + " setdefaultsounddevice """+Device2+""" 1", 0
        ws.run nircmd + " setdefaultsounddevice """+Device2+""" 2", 0
		ws.run "powershell -WindowStyle Hidden -NoLogo -executionpolicy bypass -file Speakers.ps1",0
   end if
else
    ' Headphones were set, toggle back to speakers n shit
    Set objShellLink = objFolderItem.GetLink    
    objShellLink.SetIconLocation "C:\Windows\System32\mmres.dll", 0
    objShellLink.Save()
    objFolderItem.Name = Device1Name

    ws.run nircmd + " setdefaultsounddevice """+Device1+""" 0", 0
    ws.run nircmd + " setdefaultsounddevice """+Device1+""" 1", 0
    ws.run nircmd + " setdefaultsounddevice """+Device1+""" 2", 0
	ws.run "powershell -WindowStyle Hidden -NoLogo -executionpolicy bypass -file headphones.ps1",0
end if