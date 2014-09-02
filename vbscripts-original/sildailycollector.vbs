''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Copyright (C) 2013, Microsoft Corporation.
'  All rights reserved.
'
'  File Name:
'      sildailycollector.vbs
'
'  Abstract:
'      This script invokes a software inventory collection via WMI.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Event level "constants"
EventLevelTypeSuccess = 0
EventLevelTypeError = 1
EventLevelTypeWarning = 2
EventLevelTypeInformational = 4
EventLevelTypeAuditSuccess = 8
EventLevelTypeAuditFailure = 16

' Create the shell.
Set shell = WScript.CreateObject("WScript.Shell")

' Connect to WMI.
computerName = "."
wmiConnectionString = "winmgmts:{impersonationLevel=impersonate}!\\" & computerName & "\root\InventoryLogging"
Call Log(EventLevelTypeInformational, "Software inventory: Connecting to WMI: " & wmiConnectionString)

Set wmiServices = GetObject(wmiConnectionString)
Set tasks = wmiServices.Get("Msft_MiStreamTasks")
Call Log(EventLevelTypeInformational, "Software inventory: Connected to WMI.")

Set inParams = tasks.Methods_("Push").inParameters.SpawnInstance_()
inParams.filename = "silstream.mof"
Call Log(EventLevelTypeInformational, "Software inventory: Input filename: " & inParams.filename)

' Execute the collection.
On Error Resume Next
Set outParams = tasks.ExecMethod_("Push", inParams)

If Err.Number = 0 Then
    Call Log(EventLevelTypeInformational, "Software inventory: Collection successful.")
Else
    Call Log(EventLevelTypeError, "Software inventory: Collection failed. Error Number: " & Err.Number & ", Source: " & Err.Source & ", Description: " & Err.Description)
End If

Function Log(level, message)
    If shell Is Nothing Then
        WScript.Echo level & ": " & message
    Else
        shell.LogEvent level, message, computerName
    End If
End Function