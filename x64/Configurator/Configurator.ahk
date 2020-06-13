; 
;   ViPER4Windows Configurator 1.0.0.5 (20-04-2019)
;   Author: alanfox2000
;

#NoEnv
#NoTrayIcon
#SingleInstance Force
SetWorkingDir %A_ScriptDir%
AhkPath = %A_WorkingDir%\AutoHotkey.exe
full_command_line := DllCall("GetCommandLine", "str")
if not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)"))
{
    {
        if A_IsCompiled
            Run *RunAs "%A_ScriptFullPath%" /restart
        else
            Run *RunAs "%AhkPath%" /restart "%A_ScriptFullPath%"
    }
    ExitApp
}

V4WGFX = {DA2FB532-3014-4B93-AD05-21B2C620F9C2}
LFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,1
GFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,2
UIRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,3
SFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,5
MFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,6
EFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,7
KDSFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,8
KDMFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,9
KDEFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,10
OSFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,11
OMFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,12
CompositeSFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,13
CompositeMFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,14
CompositeEFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,15
CompositeKDSFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,16
CompositeKDMFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,17
CompositeKDEFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,18
CompositeOSFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,19
CompositeOMFXRegKey = {d04e05a6-594b-4fb6-a80d-01af5eed7d1d}`,20
ProcessingLFXRegKey = {d3993a3f-99c2-4402-b5ec-a92a0367664b}`,1
ProcessingGFXRegKey = {d3993a3f-99c2-4402-b5ec-a92a0367664b}`,2
ProcessingSFXRegKey = {d3993a3f-99c2-4402-b5ec-a92a0367664b}`,5
ProcessingMFXRegKey = {d3993a3f-99c2-4402-b5ec-a92a0367664b}`,6
ProcessingEFXRegKey = {d3993a3f-99c2-4402-b5ec-a92a0367664b}`,7
ProcessingKDSFXRegKey = {d3993a3f-99c2-4402-b5ec-a92a0367664b}`,8
ProcessingKDMFXRegKey = {d3993a3f-99c2-4402-b5ec-a92a0367664b}`,9
ProcessingKDEFXRegKey = {d3993a3f-99c2-4402-b5ec-a92a0367664b}`,10
ProcessingOSFXRegKey = {d3993a3f-99c2-4402-b5ec-a92a0367664b}`,11
ProcessingOMFXRegKey = {d3993a3f-99c2-4402-b5ec-a92a0367664b}`,12
DISABLESYSFXRegKey = {1da5d803-d492-4edd-8c23-e0c0ffee7f0e}`,5
SetACL = %A_WorkingDir%\SetACL.exe
RenderPath = HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render
CapturePath = HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture
ConnectorSubKey = {a45c254e-df1c-4efd-8020-67d146a850e0}`,2
DeviceSubKey = {b3f8fa53-0004-438e-9003-51a46e139bfc}`,6 
LoadPath = %RenderPath%
SoftwareTitle =  ViPER4Windows Configurator
Reg = %A_WinDir%\system32\reg.exe

Menu, Tray, Icon, Configurator.ico
Gui, -MinimizeBox -MaximizeBox
Gui Add, Text, x8 y8 w454 h23 +0x200, Please select the devices for which the V4W shall be installed:
Gui Add, Button, x8 y264 w266 h23 gCopy, Copy Device GUID to clipboard
Gui Add, Button, x408 y264 w80 h23 vbtn_install gInstall, Install
Gui Add, Button, x504 y264 w80 h23 vbtn_uninstall gUninstall, Uninstall
Gui Add, CheckBox, x16 y296 w271 h23 vIsDISABLESYSFX gDisable_SysFx, Disable all Enhacemnts (Current Selected Device)
Gui Add, Tab3, x8 y32 w582 h220, Playback devices|Capture devices
Gui Tab, 1
Gui Add, ListView, AltSubmit vRenderList gRenderListView x16 y56 w565 h185, Connector|Device|Status|GUID
Gui Tab, 2
Gui Add, ListView, AltSubmit vCaptureList gCaptureListView x16 y56 w565 h185, Connector|Device|Status|GUID

Gui Show, w600 h331, %SoftwareTitle%

gosub Refresh
Return

Refresh:
Gui, ListView, RenderList
LV_Delete()
Loop, Reg, %RenderPath%, K
{
    RegRead, DeviceState, %RenderPath%\%A_LoopRegName%, DeviceState
    
    If !(DeviceState = "4" or DeviceState = "268435460" or DeviceState = "536870916" or DeviceState =‭ "805306372"‬)
    {
        Render_GUID = %A_LoopRegName%
        RegRead, Connector, %RenderPath%\%Render_GUID%\Properties, %ConnectorSubKey%
        RegRead, Device, %RenderPath%\%Render_GUID%\Properties, %DeviceSubKey%
        RegRead, GFX, %RenderPath%\%Render_GUID%\FxProperties, %GFXRegKey%
        If GFX = %V4WGFX%
        {
            StatusText = V4W had been installed
        } else 
        {
            StatusText = V4W can be installed
        }
        LV_Add("", Connector, Device, StatusText, Render_GUID)
    }
}
LV_ModifyCol()
LV_ModifyCol(4, 0)

Gui, ListView, CaptureList
LV_Delete()
Loop, Reg, %CapturePath%, K
{
    RegRead, DeviceState, %CapturePath%\%A_LoopRegName%, DeviceState
    
    If !(DeviceState = "4" or DeviceState = "268435460" or DeviceState = "536870916" or DeviceState =‭ "805306372"‬)
    {
        Capture_GUID = %A_LoopRegName%
        RegRead, Connector, %CapturePath%\%Capture_GUID%\Properties, %ConnectorSubKey%
        RegRead, Device, %CapturePath%\%Capture_GUID%\Properties, %DeviceSubKey%
        RegRead, GFX, %CapturePath%\%Capture_GUID%\FxProperties, %GFXRegKey%
        If GFX = %V4WGFX%
        {
            StatusText = V4W had been installed
        } else {
            StatusText = V4W can be installed
        }
        LV_Add("", Connector, Device, StatusText, Capture_GUID)  
    }       
}
LV_ModifyCol()
LV_ModifyCol(4, 0)

Gui, Submit, NoHide
Return

Copy:
gosub GetSelectedDevice
clipboard = %SelectedGUID%
Return

RenderListView:
Gui, ListView, RenderList
RowNumber = 0
Loop
{
RowNumber := LV_GetNext(RowNumber)
if not RowNumber
    break
Type = Render
LV_GetText(SelectedGUID, RowNumber, 4)
LV_GetText(StatusText, RowNumber, 3)
gosub Disable_UninstallButton
gosub Disable_SysFxRead
}
return

CaptureListView:
Gui, ListView, CaptureList
RowNumber = 0
Loop
{
RowNumber := LV_GetNext(RowNumber)
if not RowNumber
    break
Type = Capture
LV_GetText(SelectedGUID, RowNumber, 4)
LV_GetText(StatusText, RowNumber, 3)
gosub Disable_UninstallButton
gosub Disable_SysFxRead
}
return

Disable_UninstallButton:
if (!FileExist(A_WorkingDir "\" SelectedGUID ".reg")) OR (StatusText = "V4W can be installed") {
    GuiControl, Disable, btn_Uninstall
    GuiControl,, btn_install, Install
} else {
    GuiControl, Enable, btn_Uninstall
    GuiControl,, btn_install, Reinstall
}
return

GetSelectedDevice:
if Type = Render
{
    Gui, ListView, RenderList
    LoadPath = %RenderPath%
}
if Type = Capture
{
    Gui, ListView, CaptureList
    LoadPath = %CapturePath%
}
RowNumber = 0
RowNumber := LV_GetNext(RowNumber)
if not RowNumber
{
    msgbox, 0x2040, %SoftwareTitle%, Please select a device
    exit
}
return

Install:
gosub GetSelectedDevice
Gui, +Disabled
SplashTextOn, , , Please Wait...
if !FileExist(A_WorkingDir "\" SelectedGUID ".reg") {
    Runwait, %ComSpec% /c "reg export %LoadPath%\%SelectedGUID% "%A_WorkingDir%\%SelectedGUID%.reg" /y",, Hide
}
Runwait, %ComSpec% /c "net stop Audiosrv /yes",, Hide
gosub RegistryKeyPermissions
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %LFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %GFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %UIRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %SFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %MFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %EFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %KDSFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %KDMFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %KDEFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %OSFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %OMFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %CompositeSFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %CompositeMFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %CompositeEFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %CompositeKDSFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %CompositeKDMFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %CompositeKDEFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %CompositeOSFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %CompositeOMFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %ProcessingLFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %ProcessingGFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %ProcessingSFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %ProcessingMFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %ProcessingEFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %ProcessingKDSFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %ProcessingKDMFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %ProcessingKDEFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %ProcessingOSFXRegKey%
RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %ProcessingOMFXRegKey%
RegWrite, REG_SZ, %LoadPath%\%SelectedGUID%\FxProperties, %GFXRegKey%, {DA2FB532-3014-4B93-AD05-21B2C620F9C2}
;RegWrite, REG_MULTI_SZ, %LoadPath%\%SelectedGUID%\FxProperties, %ProcessingGFXRegKey%, {C18E2F7E-933D-4965-B7D1-1EEF228D2AF3}
If ErrorLevel = 1
{
    Runwait, %ComSpec% /c "%reg% import "%A_WorkingDir%\%SelectedGUID%.reg"",, Hide
    Runwait, %ComSpec% /c "net start Audiosrv",, Hide
    gosub Refresh
    SplashTextOff
    msgbox, 0x2010, %SoftwareTitle%, Access %LoadPath%\%SelectedGUID%\FxProperties failed. Please take ownership of the registry key manually.
}
else
{
    Runwait, %ComSpec% /c "net start Audiosrv",, Hide
    gosub Refresh
    SplashTextOff
    gosub Popup_Finish
}
return

Uninstall:
gosub GetSelectedDevice
SplashTextOn, , , Please Wait...
Runwait, %ComSpec% /c "net stop Audiosrv /yes",, Hide
gosub RegistryKeyPermissions
RegDelete, %LoadPath%\%SelectedGUID%
Runwait, %ComSpec% /c "%reg% import "%A_WorkingDir%\%SelectedGUID%.reg"",, Hide
Runwait, %ComSpec% /c "net start Audiosrv",, Hide
FileDelete, %A_WorkingDir%\%SelectedGUID%.reg
gosub Refresh
SplashTextOff
gosub Popup_Finish
return

Popup_Finish:
MsgBox, 0x2040, %SoftwareTitle%, Operation finished.
Gui, -Disabled
Return

Disable_SysFxRead:
RegRead, IsDISABLESYSFX, %LoadPath%\%SelectedGUID%\FxProperties,%DISABLESYSFXRegKey%
If ErrorLevel = 1
{
    IsDISABLESYSFX := 0
}
GuiControl,, IsDISABLESYSFX, %IsDISABLESYSFX%
Return

Disable_SysFx:
gosub GetSelectedDevice
Gui, +Disabled
SplashTextOn, , , Please Wait...
gosub RegistryKeyPermissions
GuiControlGet, IsDISABLESYSFX
If IsDISABLESYSFX = 0
{
    RegDelete, %LoadPath%\%SelectedGUID%\FxProperties, %DISABLESYSFXRegKey%
}
If IsDISABLESYSFX = 1
{
    RegWrite, REG_DWORD, %LoadPath%\%SelectedGUID%\FxProperties, %DISABLESYSFXRegKey%, 1
}
Runwait, %ComSpec% /c "net stop Audiosrv /yes",, Hide
Runwait, %ComSpec% /c "net start Audiosrv",, Hide
SplashTextOff
Gui, -Disabled
Return

RegistryKeyPermissions:
Runwait, "%SetACL%" -on "%LoadPath%\%SelectedGUID%" -ot reg -actn setowner -ownr "n:Administrators",, Hide
Runwait, "%SetACL%" -on "%LoadPath%\%SelectedGUID%" -ot reg -actn ace -ace "n:Administrators;p:full",, Hide
Runwait, "%SetACL%" -on "%LoadPath%\%SelectedGUID%" -ot reg -actn ace -ace "n:Users;p:full",, Hide
Runwait, "%SetACL%" -on "%LoadPath%\%SelectedGUID%\FxProperties" -ot reg -actn setowner -ownr "n:Administrators",, Hide
Runwait, "%SetACL%" -on "%LoadPath%\%SelectedGUID%\FxProperties" -ot reg -actn ace -ace "n:Administrators;p:full",, Hide
Runwait, "%SetACL%" -on "%LoadPath%\%SelectedGUID%\FxProperties" -ot reg -actn ace -ace "n:Users;p:full",, Hide
Runwait, "%SetACL%" -on "%LoadPath%\%SelectedGUID%\Properties" -ot reg -actn setowner -ownr "n:Administrators",, Hide
Runwait, "%SetACL%" -on "%LoadPath%\%SelectedGUID%\Properties" -ot reg -actn ace -ace "n:Administrators;p:full",, Hide
Runwait, "%SetACL%" -on "%LoadPath%\%SelectedGUID%\Properties" -ot reg -actn ace -ace "n:Users;p:full",, Hide
Return

GuiEscape:
GuiClose:
    ExitApp
