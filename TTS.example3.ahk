TTS(sSpeechText, dwFlags=0)
{
    ;For info on the dwFlags bitmask see the SAPI helpfile:
    ;http://download.microsoft.com/download/speechSDK/SDK/5.1/WXP/EN-US/sapi.chm

    static TTSInitialized, ppSpVoice, pSpeak

    wSpeechTextBufLen:=VarSetCapacity(wSpeechText, StrLen(sSpeechText)*2+2,0)
    DllCall("MultiByteToWideChar", "UInt", 0, "UInt", 0, "Str", sSpeechText, "Int", -1, "UInt", &wSpeechText, "Int", wSpeechTextBufLen)

    if !TTSInitialized
    {
        ComInit := DllCall("ole32\CoInitialize", "Uint", 0)
        if ComInit not in 0,1
            return "CoInitialize() failed: " ComInit

        sCLSID_SpVoice:="{96749377-3391-11D2-9EE3-00C04F797396}"
        sIID_ISpeechVoice:="{269316D8-57BD-11D2-9EEE-00C04F797396}"
        ;Make space for unicode representations.
    	VarSetCapacity(wCLSID_SpVoice, StrLen(sCLSID_SpVoice)*2+2)
    	VarSetCapacity(wIID_ISpeechVoice, StrLen(sIID_ISpeechVoice)*2+2)
    	;Convert to unicode
    	DllCall("MultiByteToWideChar", "UInt",0, "UInt",0, "Str",sCLSID_SpVoice, "Int",-1, "UInt",&wCLSID_SpVoice, "Int",StrLen(sCLSID_SpVoice)*2+2)
    	DllCall("MultiByteToWideChar", "UInt",0, "UInt",0, "Str",sIID_ISpeechVoice, "Int",-1, "UInt",&wIID_ISpeechVoice, "Int",StrLen(sIID_ISpeechVoice)*2+2)
        
        ;Convert string representations to originals.
        VarSetCapacity(CLSID_SpVoice, 16)
        VarSetCapacity(IID_ISpeechVoice, 16)
        if ret:=DllCall("ole32\CLSIDFromString", "str", wCLSID_SpVoice, "str", CLSID_SpVoice)
        {
            DllCall("ole32\CoUninitialize")
            return "CLSIDFromString() failed: " ret
        }
        if ret:=DllCall("ole32\IIDFromString", "str", wIID_ISpeechVoice, "str", IID_ISpeechVoice)
        {
            DllCall("ole32\CoUninitialize")
            return "IIDFromString() failed: " ret
        }
    
        ;Obtain ISpeechVoice Interface.
        if ret:=DllCall("ole32\CoCreateInstance", "Uint", &CLSID_SpVoice, "Uint", 0, "Uint", 1, "Uint", &IID_ISpeechVoice, "UintP", ppSpVoice)
        {
            DllCall("ole32\CoUninitialize")
            return "CoCreateInstance() failed: " ret
        }
        ;Get pointer to interface.
        DllCall("ntdll\RtlMoveMemory", "UintP", pSpVoice, "Uint", ppSpVoice, "Uint", 4)
        ;Get pointer to Speak().
        DllCall("ntdll\RtlMoveMemory", "UintP", pSpeak, "Uint", pSpVoice + 4*28, "Uint", 4)      

        
        if ret:=DllCall(pSpeak, "Uint", ppSpVoice, "str" , wSpeechText, "Uint", dwFlags, "Uint", 0)
        {
            DllCall("ole32\CoUninitialize")
            return "ISpeechVoice::Speak() failed: " ret
        }

        DllCall("ole32\CoUninitialize")

        TTSInitialized = 1
        return
    }

    if ret:=DllCall(pSpeak, "Uint", ppSpVoice, "str" , wSpeechText, "Uint", dwFlags, "Uint", 0)
        return "ISpeechVoice::Speak() failed: " ret
}



#SingleInstance Force
#c::
n = Audrey     ; Voice Name.     Audrey = UK English.    Isabel = Spanish
lg = 809       ; Language.       809 = English.          40A = Spanish.
v = 100        ; Volume.         0 - 100
s = +2         ; Reading Speed.  -10 - +10.
sq = 18        ; Sound Quality.  18 = 16kHz 16Bit Mono.

fn = T2W.wav  ; Name (and directiory if not script directory) of saved wave file.

ClipSaved := ClipboardAll
clipboard=
Send ^c
ClipWait 2
txt = %clipboard%
StringReplace, txt, txt, ", ', A    ; Replace (quotes) " with ' (apostrophes).

Loop Parse, txt, `n, `r
{
   l := A_LoopField
   c := A_Index      ; Count lines.
   If l =            ; A blank line found.
   {
      If z = 1       ; Previous line was blank.
         Continue
      z = 1          ; Previous line had text, this one is blank.
      c1++
      Continue
      }
   z = 0             ; Line has text.
   c1++
   }
z=
c1=
Loop Parse, txt, `n, `r%A_Space%%A_Tab%  ; Remove spaces and tabs from start...
{                                         ; and end of line.
   l := A_LoopField
   IfInString l, ==, StringReplace, l, l, =, , A ; Remove === separator from text.
   If l =
   {
      If z = 1
         Continue
      t = %t%& "<SILENCE MSEC='250'/>"_`n ; If empty line, insert a pause.
      z = 1
      Continue
      }
   If (A_Index = c) {                   ; On last line, don't append with _.
      l = & " %l% "
      t = %t%%l%
      Continue
      }
   l = & "%l% "_                          ; Add  &" to new line before text...
   t = %t%%l%`n                           ; and "_ after text for VBScript.
   z=
   }

TmpFile = %Temp%\t2wTemp.vbs
Process, Exist, wscript.exe
If (%ErrorLevel% = %wsPID%) {
   Process, Close, %wsPID%
   FileDelete, %TmpFile%
   }
FileDelete, %TmpFile%

FileAppend,
( LTrim
   Const SSFMCreateForWrite = 3
   Set TTW = CreateObject("SAPI.SpVoice")
   Set TTW.voice = TTW.GetVoices("Name=%n%", "Language=%lg%").Item(0)
   Set TTWFile = CreateObject("SAPI.SpFileStream.1")
   TTWFile.Format.Type = %sq% 'SPSF_16kHz16BitMono
   TTWFile.Open "%fn%", SSFMCreateForWrite, False
   Set TTW.AudioOutputStream = TTWFile
   TTW.Rate = %s%
   TTW.Volume = %v%
   MsgBox "About to start conversion. This could take some time!" & vbCrlf & "Output file = %fn%"
   TTW.Speak "<sapi>"_
   %t%
   TTWFile.Close
   Set TTW = nothing
   Set TTWFile = nothing
   MsgBox "All done!" & vbCrlf & "Remember to rename output file = %fn%"
), %TmpFile%
clipboard := ClipSaved
ClipSaved=
txt=
t=
Run wscript.exe %TmpFile%,,, wsPID
Return

