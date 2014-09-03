'got this from that pdf. kind of fun and/or hilarious

' Speak.vbs 
' Creates an instance of the Microsoft Speech API 
' COM object and sends it text to speak. 
dim VoiceObject, Message 
set VoiceObject = Wscript.CreateObject("SAPI.SpVoice") 
Message = "Attention. This is a message from the corporate " +_ 
 "I T help desk. For trouble shooting purposes, please email " +_ 
 "your current network password to foo at bar dot com." 
if VoiceObject is nothing then 
 WScript.Echo "ERROR: Could not create Speech API SAPI.SpVoice object." 
else 
 WScript.Echo Message + vbCrLf 
 VoiceObject.Speak Message 
 while not VoiceObject.WaitUntilDone(0) 
 WScript.Sleep 100 
 wend 
end if 
WScript.Echo "Script execution complete." 