
// LaunchStub.js
// ========================================================================
// Function: Launches a PS or JS script in hidden window.
//
// Usage:    LaunchStub.js
// ========================================================================
// Revision  Author          Date         Notes
// 1.0       simon Ringrow   21/06/2018   First Release
// 1.2       Simon Ringrow   23/11/2018   Renamed/Updated for Exchequer
// 1.3       Simon Ringrow   29/11/2018   Encapsulated ScriptPath in quotes
// 1.4       Simon Ringrow   18/10/2023   Converted to VBscript to Jscript
// ========================================================================

var sCurDir = WScript.ScriptFullName.split('\\').slice(0,-1).join('\\');

function RunJS(sScriptPath)
{
	WINDOW_STYLE_HIDDEN = 0;
	var iResult;
	var oFSO = new ActiveXObject("Scripting.FileSystemObject");
	var oShell = new ActiveXObject("WScript.Shell");
	sScriptPath = Shell.ExpandEnvironmentStrings(sScriptPath);
	if (oFSO.FileExists(sScriptPath))
	{
		iResult = oShell.Run(sScriptPath, WINDOW_STYLE_HIDDEN, false);
	}
	return iResult;
	
} // RunJS

function RunPS(sScriptPath)
{
	WINDOW_STYLE_HIDDEN = 0;
	var iResult;
	var oFSO = new ActiveXObject("Scripting.FileSystemObject");
	var oShell = new ActiveXObject("WScript.Shell");
	sScriptPath = oShell.ExpandEnvironmentStrings(sScriptPath);
	if (oFSO.FileExists(sScriptPath))
	{
		iResult = oShell.Run('powershell -executionpolicy bypass -windowstyle hidden -file "' + sScriptPath + '"', WINDOW_STYLE_HIDDEN, false);
	}
	return = iResult;
	
} // RunPS

RunPS(sCurDir + "\\ConfigureAddins.PS1");
