// Create an instance of the scripting Shell Object
 WshShell = WScript.CreateObject("WScript.Shell");
 
 // Have the Shell Object call ShellExecute on our HTML // file
 WshShell.Run("Start.htm", 1, 0);
 
 // Destroy the Shell Object
 WScript.DisconnectObject(WshShell);
