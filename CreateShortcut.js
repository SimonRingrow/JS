
// CreateShortcut.js
// ==================================================================================================================================
// Function: Creates a shortcut.
//
// Usage: [cscript] CreateShort.js /ShortcutPath:<shortcutpath>, /TargetPath:<targetpath>, [/Arguments:<arguments>],
//                                 [/WorkingDirectory:<workingdirectory>], [/Description:<description>, /IconLocation:<iconlocation>
// Example: CreateShortcut /S:"%ONEDRIVE%\Desktop\Notepad.lnk" /T:c:\windows\notepad.exe
// ==================================================================================================================================
// Revision  Author          Date         Notes
// 1.0       simon Ringrow   27/07/2023   First Release
// 1.2       Simon Ringrow   20/10/2023   Converted to VBscript to Jscript
// ==================================================================================================================================

sArguments = "", sDescription = "", sHotKey = "", sIconLocation = "", sShortcutPath = "", sTargetPath = "", iWindowStyle = 1, sWorkingDirectory = "";

// Load command line arguments into dictionary
var oArguments = WScript.Arguments;
var oDictionary = new ActiveXObject("Scripting.Dictionary");
for (iCount = 0; iCount < oArguments.length; iCount++)
{
	sParam = oArguments(iCount).split(":")[0];
	sValue = oArguments(iCount).substring(oArguments(iCount).indexOf(":")+1, oArguments(iCount).length - oArguments(iCount).indexOf(":")+2)
	oDictionary.Add(sParam, sValue);
}

// Parse the dictionary to update shortcut parameters
var aKeys = (new VBArray(oDictionary.Keys()));
for (iCount = 0; iCount < oDictionary.count; iCount++)
{
	switch ((aKeys.getItem(iCount)).toLowerCase())
	{
		case "/arguments", "/a":
			sArguments = oDictionary(aKeys.getItem(iCount));
		case "/description", "/d":
			sDescription = oDictionary(aKeys.getItem(iCount));
		case "/hotkey", "/h":
			sHotKey = oDictionary(aKeys.getItem(iCount));
		case "/iconlocation", "/i":
			sIconLocation = oDictionary(aKeys.getItem(iCount));
		case "/shortcutpath", "/s":
			sShortcutPath = oDictionary(aKeys.getItem(iCount));
		case "/targetpath", "/t":
			sTargetPath = oDictionary(aKeys.getItem(iCount));
//		case "/windowstyle", "/w":
//			sWindowStyle = oDictionary(aKeys.getItem(iCount));
		case "/workingdirectory", "/w":
			sWorkingDirectory = oDictionary(aKeys.getItem(iCount));
	}
}

// Exit if minimum paramaters not present
if (sShortcutPath == "" || sTargetPath == "")
{
	WScript.Echo ("Incorrect commandline passed:\n" +
				  "Usage:\t\[cscript] CreateShort.vbs /ShortcutPath:<shortcutpath>, /TargetPath:<targetpath>, [/Arguments:<arguments>]\n" +
				  "\t[/WorkingDirectory:<workingdirectory>], [/Description:<description>], [/IconLocation:<iconlocation>]");
	WScript.Quit();
}

CreateShortcut(sShortcutPath, sTargetPath, sArguments, sWorkingDirectory, sDescription, sIconLocation, sHotKey, iWindowStyle)

function CreateShortcut(sShortcutPath, sTargetPath, sArguments, sWorkingDirectory, sDescription, sIconLocation, sHotKey, iWindowStyle)
{
	var oShell = new ActiveXObject("WScript.Shell");
	var oShortcut = oShell.CreateShortcut(oShell.ExpandEnvironmentStrings(sShortcutPath));
	oShortcut.Arguments = oShell.ExpandEnvironmentStrings(sArguments);
	oShortcut.Description = sDescription;
	oShortcut.HotKey = sHotKey;
	if (sIconLocation == "")
	{
		oShortcut.IconLocation = oShell.ExpandEnvironmentStrings(sTargetPath) + ",0";
	}
	else
	{
		oShortcut.IconLocation = oShell.ExpandEnvironmentStrings(sIconLocation);
	}
	oShortcut.TargetPath = oShell.ExpandEnvironmentStrings(sTargetPath);
	oShortcut.WindowStyle = iWindowStyle;
	oShortcut.WorkingDirectory = oShell.ExpandEnvironmentStrings(sWorkingDirectory);
	oShortcut.Save();

} // CreateShortcut
