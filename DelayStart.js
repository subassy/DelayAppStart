var wbemFlagReturnImmediately = 0x10; 
var wbemFlagForwardOnly = 0x20; 
 
   var objWMIService = GetObject("winmgmts:\\\\.\\root\\CIMV2"); 
   var colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ClassicCOMClass", "WQL", 
                                          wbemFlagReturnImmediately | wbemFlagForwardOnly); 
 
   var enumItems = new Enumerator(colItems); 
   for (; !enumItems.atEnd(); enumItems.moveNext()) { 
      var objItem = enumItems.item(); 
 
      WScript.Echo("Caption: " + objItem.Caption); 
      WScript.Echo("Component Id: " + objItem.ComponentId); 
      WScript.Echo("Description: " + objItem.Description); 
      WScript.Echo("Install Date: " + objItem.InstallDate); 
      WScript.Echo("Name: " + objItem.Name); 
      WScript.Echo("Status: " + objItem.Status); 
      WScript.Echo(); 
   } 
