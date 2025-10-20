; Based on this forum post by 0x00
;https://www.autohotkey.com/boards/viewtopic.php?f=76&t=49980&p=387886#p387886


; Press the Menu key on keyboard once to switch to the next audio output
AppsKey:: SetNextDevice()

; Switches the audio output to the next available device
SetNextDevice(){
    Devices := GetAudioDevices()
    CurrentDevice := GetCurrentDevice()
    nextDevice := ""

    ; MsgBox % "Audio Devices:`n" . ObjDump(Devices)
    ; MsgBox Current Device: %CurrentDevice%

    if InStr(CurrentDevice, "LEN Y44w-10") {
        ; If playing through LEN Y44w-10, switch to Bluetooth headphones if available; otherwise, switch to speakers
        headphoneDeviceId := GetDeviceID(Devices, "Headphones")
    
        nextDevice := (StrLen(headphoneDeviceId) > 0) ? "Headphones" : "Speakers"
    } 
    else if InStr(CurrentDevice, "Headphones") {
        ; If playing through Bluetooth headphones, switch to speakers
        nextDevice := "Speakers"
    } 
    else if InStr(CurrentDevice, "Speakers") {
        ; If playing through speakers, switch to LEN Y44w-10
        nextDevice := "LEN Y44w-10"
    }

	SetDefaultEndpoint(GetDeviceID(Devices, nextDevice))
    TrayTip, AudioSwitch, Switching to %nextDevice%,10,17
}

; Get the list of all active audio output devices
GetAudioDevices() {
    Devices := {}
    IMMDeviceEnumerator := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}")
    DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+3*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 0x1, "UPtrP", IMMDeviceCollection, "UInt")
    ObjRelease(IMMDeviceEnumerator)

    DllCall(NumGet(NumGet(IMMDeviceCollection+0)+3*A_PtrSize), "UPtr", IMMDeviceCollection, "UIntP", Count, "UInt")
    Loop % Count {
        DllCall(NumGet(NumGet(IMMDeviceCollection+0)+4*A_PtrSize), "UPtr", IMMDeviceCollection, "UInt", A_Index-1, "UPtrP", IMMDevice, "UInt")
        
        DllCall(NumGet(NumGet(IMMDevice+0)+5*A_PtrSize), "UPtr", IMMDevice, "UPtrP", pBuffer, "UInt")
        DeviceID := StrGet(pBuffer, "UTF-16")
        DllCall("Ole32.dll\CoTaskMemFree", "UPtr", pBuffer)
        
        DllCall(NumGet(NumGet(IMMDevice+0)+4*A_PtrSize), "UPtr", IMMDevice, "UInt", 0x0, "UPtrP", IPropertyStore, "UInt")
        ObjRelease(IMMDevice)

        VarSetCapacity(PROPVARIANT, A_PtrSize == 4 ? 16 : 24)
        VarSetCapacity(PROPERTYKEY, 20)
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", "UPtr", &PROPERTYKEY)
        NumPut(14, &PROPERTYKEY + 16, "UInt")
        DllCall(NumGet(NumGet(IPropertyStore+0)+5*A_PtrSize), "UPtr", IPropertyStore, "UPtr", &PROPERTYKEY, "UPtr", &PROPVARIANT, "UInt")
        DeviceName := StrGet(NumGet(&PROPVARIANT + 8), "UTF-16")
        DllCall("Ole32.dll\CoTaskMemFree", "UPtr", NumGet(&PROPVARIANT + 8))
        ObjRelease(IPropertyStore)
        
        ObjRawSet(Devices, DeviceName, DeviceID)
    }
    ObjRelease(IMMDeviceCollection)
    return Devices
}

; Get the current active output device
GetCurrentDevice()
{
    CurrentDevice := ""
    RunWait, %ComSpec% /c "powershell -command "(Get-AudioDevice -Playback).Name" > %temp%\currentdevice.txt",, Hide
    FileRead, CurrentDevice, %temp%\currentdevice.txt
    CurrentDevice := Trim(CurrentDevice)
    Return CurrentDevice
}

SetDefaultEndpoint(DeviceID)
{
    ;MsgBox "DeviceID" + %DeviceID%
    IPolicyConfig := ComObjCreate("{870af99c-171d-4f9e-af0d-e63df40c2bc9}", "{F8679F50-850A-41CF-9C72-430F290290C8}")
    DllCall(NumGet(NumGet(IPolicyConfig+0)+13*A_PtrSize), "UPtr", IPolicyConfig, "UPtr", &DeviceID, "UInt", 0, "UInt")
    ObjRelease(IPolicyConfig)
}

GetDeviceID(Devices, Name)
{
    ;MsgBox Devices: %Devices%, Name: %Name%
    For DeviceName, DeviceID in Devices
        If (InStr(DeviceName, Name))
            Return DeviceID
}

; Used to print out the device obj when debugging
ObjDump(obj) {
    output := ""
    for key, value in obj
        output .= key . " : " . value . "`n"
    return output
}
