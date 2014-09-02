EventLevelTypeSuccess = 0
EventLevelTypeError = 1
EventLevelTypeWarning = 2
EventLevelTypeInformational = 4
EventLevelTypeAuditSuccess = 8
EventLevelTypeAuditFailure = 16

Dim ArgIndex

Set Wshell = Wscript.CreateObject("WScript.Shell")
Call Log(EventLevelTypeInformational, "Executed @ " & Now)

strComputer = "."
strWmiConnectionString = "winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\Microsoft\Windows\ServerManager"
Call Log(EventLevelTypeInformational, strWmiConnectionString)

Set objWMIService = GetObject(strWmiConnectionString)
Call Log(EventLevelTypeInformational, "Successfully connected to wmi")

Set objSmTasks = objWMIService.Get("MSFT_ServerManagerTasks")
Call Log(EventLevelTypeInformational, "Successfully got the class")

If Wscript.Arguments.Count > 0 Then
    strMethodName = Wscript.Arguments(0)
    Call Log(EventLevelTypeInformational, "Method name: " & strMethodName)
    
    ' Obtain an InParameters object specific to the strMethodName method.
    Set objInParams = Nothing
    If Wscript.Arguments.Count > 1 Then
        Call Log(EventLevelTypeInformational, "Parameters specified.")
        Set objInParams = objSmTasks.Methods_(strMethodName).inParameters.SpawnInstance_()
        Call Log(EventLevelTypeInformational, "Successfully created the input parameters object.")

        ' Add the input parameters.
        ArgIndex = 1
        For Each objInParam in objInParams.Properties_
            If ArgIndex >= Wscript.Arguments.Count Then
                Call Log(EventLevelTypeWarning, "Number of specified parameters to the script{" & ArgIndex & _
                         "} is less than the required number of parameters{" & _
                         "}, some parameters to the function will be NULL")
                Exit For
            End If
            Call Log(EventLevelTypeInformational, "Setting parameter: [" & objInParam.Name & " = " & Wscript.Arguments(ArgIndex) & "]")
            objInParam.Value = Wscript.Arguments(ArgIndex)
            ArgIndex = ArgIndex + 1
        Next
    End If

    Set objOutParams = Nothing
    If objInParams Is Nothing Then
        Call Log(EventLevelTypeInformational, "Calling " & strMethodName & " with no input parameters.")
        Set objOutParams = objSmTasks.ExecMethod_(strMethodName)
    Else
        Call Log(EventLevelTypeInformational, "Calling " & strMethodName & " with the specified parameters.")
        Set objOutParams = objSmTasks.ExecMethod_(strMethodName, objInParams)
    End If

    If Error = 0 Then
        Call Log(EventLevelTypeInformational, "Method " & strMethodName & " succeeded.")
        If objOutParams Is Nothing Then
            Call Log(EventLevelTypeWarning, "No output returned.")
        Else
            strWmiObject = WmiObjectToString(objOutParams, 0)
            Call Log(EventLevelTypeInformational, strWmiObject)
        End If
    Else
        Call Log(EventLevelTypeError, "Method " & strMethodName & " failed. [" & Error & "]")
    End If
Else
    Call Log(EventLevelTypeError, "Method name must be provided.")
End If

Function WmiObjectToString(wmiObject, intendation)
    Dim strResult
    Dim strIndentation
    
    strIndentation = CStr("")
    For i = 1 to intendation
        strIndentation = strIndentation & " "
    Next
    
    Set objPath = wmiObject.Path_
    strResult = strIndentation & "Namespace\Class: " & objPath.Namespace & "\" & objPath.Class & Vbcrlf
    strResult = strResult & strIndentation & "Path: " & objPath.Path & Vbcrlf
    strResult = strResult & strIndentation & "Properties: " & Vbcrlf

    For Each objProperty in wmiObject.Properties_
        strResult = strResult & strIndentation & "    Name: " & objProperty.Name & " Value: " & objProperty.Value & Vbcrlf
    Next
    
    WmiObjectToString = strResult
End Function

Function Log(levelType, strMessage)
    If Wshell Is Nothing Then
        Wscript.Echo levelType & ": " & strMessage
    Else
        Wshell.LogEvent levelType, strMessage, strComputer
    End If
End Function
