Attribute VB_Name = "ModFixer"
Option Explicit
Dim REG As New cRegistry
Public Const rExplorer = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
            rEnum = "Software\Microsoft\Windows\CurrentVersion\Policies\NonEnum", _
            rWapp = "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", _
            rInternet = "Software\Policies\Microsoft\Internet Explorer\Restrictions", _
            rSystem = "Software\Microsoft\Windows\CurrentVersion\Policies\System", _
            rNetwork = "Software\Microsoft\Windows\CurrentVersion\Policies\Network", _
            rDesktop = "Control Panel\Desktop", _
            rAdvanced = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Const SMWC = "Software\Microsoft\Windows\CurrentVersion"

Function CekReg(Nm As Boolean, Root As Long, Path As String, value As String, Tipe As Byte)
    On Error Resume Next
    
    If Nm = True Then
       Select Case Tipe
              Case 1
                    REG.SaveSettingLong Root, Path, value, 1
              Case 2
                    REG.SaveSettingByte Root, Path, value, 1
              Case 3
                    REG.SaveSettingString Root, Path, value, 1
      End Select
    Else
       REG.DeleteValue Root, Path, value
    End If
    
End Function

Public Function FixRegistry()
    On Error Resume Next
    ComName = NameOfTheComputer(PCName)
    UserCom = GetUserCom()
    DoEvents
    
    ' Repair system windows-------------------------------------
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "exefile\shell\open\command", vbNullString, Chr(34) & "%1" & Chr(34) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "lnkfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "piffile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "scrfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "regfile\shell\open\command", "", "regedit.exe %1"
    REG.DeleteValue HKEY_CURRENT_USER, rSystem, "DisableTaskMgr"
    REG.DeleteValue HKEY_LOCAL_MACHINE, rSystem, "DisableTaskMgr"
    REG.DeleteValue HKEY_CURRENT_USER, rSystem, "DisableRegistryTools"
    REG.DeleteValue HKEY_LOCAL_MACHINE, rSystem, "DisableRegistryTools"
    REG.DeleteValue HKEY_CURRENT_USER, rExplorer, "NoFolderOptions"
    REG.DeleteValue HKEY_CURRENT_USER, rExplorer, "NoFind"
    REG.DeleteValue HKEY_CURRENT_USER, rExplorer, "NoRun"
    REG.DeleteValue HKEY_LOCAL_MACHINE, rExplorer, "NoFolderOptions"
    REG.DeleteValue HKEY_LOCAL_MACHINE, rExplorer, "NoFind"
    REG.DeleteValue HKEY_LOCAL_MACHINE, rExplorer, "NoRun"
        
    ' Hidden files or folder-------------------------------------
    REG.SaveSettingLong HKEY_CURRENT_USER, rAdvanced, "Hidden", 2
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden", "CheckedValue", 2
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden", "DefaultValue", 2
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden", "Bitmap", "%SystemRoot%\system32\SHELL32.dll,4"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden", "Text", "@shell32.dll,-30499"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden", "Type", "group"
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\NOHIDDEN", "CheckedValue", 2
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\NOHIDDEN", "DefaultValue", 2
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\NOHIDDEN", "Text", "@shell32.dll,-30501"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\NOHIDDEN", "Type", "radio"
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\SHOWALL", "CheckedValue", 1
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\SHOWALL", "DefaultValue", 2
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\SHOWALL", "Text", "@shell32.dll,-30500"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\Hidden\SHOWALL", "Type", "radio"

    ' Hide extensions--------------------------------------------
    REG.SaveSettingLong HKEY_CURRENT_USER, rAdvanced, "HideFileExt", 1
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\HideFileExt", "CheckedValue", 1
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\HideFileExt", "DefaultValue", 1
    REG.DeleteValue HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\HideFileExt", "HideFileExt"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\HideFileExt", "Text", "@shell32.dll,-30503"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\HideFileExt", "Type", "checkbox"
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, rAdvanced & "\Folder\HideFileExt", "UncheckedValue", 0

    ' Show super hiddens-----------------------------------------
    REG.SaveSettingLong HKEY_CURRENT_USER, rAdvanced, "ShowSuperHidden", 0
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", "Auto", "0"
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Windows\ShellNoRoam\MUICache", "@shell32.dll,-30508", "Hide protected operating system files (Recommended)"
    REG.SaveSettingString HKEY_USERS, "S-1-5-21-1417001333-1060284298-725345543-500\Software\Microsoft\Windows\ShellNoRoam\MUICache", "@shell32.dll,-30508", "Hide protected operating system files (Recommended)"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\SuperHidden", "Text", "@shell32.dll,-30508"

    ' Registered Organization & Registered Owner-----------------
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner", UserCom
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization", PCName

    ' Show Full Path at Address Bar------------------------------
    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", "FullPathAddress", 1

    ' 4k51k4-----------------------------------------------------
    REG.DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    REG.DeleteKey HKEY_USERS, "S-1-5-21-1547161642-1343024091-725345543-500\Software\Policies\Microsoft\Windows\System"
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig", 0
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR", 0
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing", 0
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI", 0
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\", "AlternateShell", "cmd.exe"
    REG.SaveSettingString HKEY_CURRENT_USER, "Control Panel\Desktop\", "SCRNSAVE.EXE", ""
    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled", 0
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled", 0
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe "
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "userinit.exe"

    ' Amburadul.Hokage Killer------------------------------------
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PaRaY_VM"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "ConfigVir"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "NviDiaGT"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "NarmonVirusAnti"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "AVManager"
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Window Title", ""
    REG.DeleteValue HKEY_LOCAL_MACHINE, " SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA"
    REG.DeleteValue HKEY_CLASSES_ROOT, "exefile", "NeverShowExt"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msconfig.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\rstrui.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\wscript.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\mmc.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\procexp.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msiexec.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\taskkill.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\cmd..exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\tasklist.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\HokageFile.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\Rin.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\Obito.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\KakashiHatake.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\HOKAGE4.exe"

    ' Flu_Ikan--------------------------------------------------
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "kebodohan"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "pemalas"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "mulut_besar"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "otak_udang"
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.microsoft.com/isapi/redir.dll?prd={SUB_PRD}&clcid={SUB_CLSID}&pver={SUB_PVER}&ar=home"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.microsoft.com/isapi/redir.dll?prd={SUB_PRD}&clcid={SUB_CLSID}&pver={SUB_PVER}&ar=home"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\dmboot.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\dmio.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\dmload.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\sermouse.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\sr.sys", "", "FSFilter System Recovery"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\vga.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\vgasave.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\dmboot.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\dmiot.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\rdpcdd.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\rdpdd.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\rdpwd.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\sermouse.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\tdpipe.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\tdtcp.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\vga.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\vgasave.sys", "", "Driver"

    LockWindowUpdate (GetDesktopWindow())
    ForceCacheRefresh
    LockWindowUpdate (0)
    DoEvents
End Function



