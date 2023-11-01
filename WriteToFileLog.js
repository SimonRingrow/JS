
function pad(sValue)
{
	return sValue < 10 ? '0'+sValue : sValue;
}

function WriteToLog(sFilePath, sData)
{
	var ForAppending = 8
	var Now = new Date();
	var Time = pad(Now.getDate()) + "-" + pad((Now.getMonth() + 1)) + "-" + Now.getFullYear() + " " + pad(Now.getHours()) + ":" + pad(Now.getMinutes()) + ":" + pad(Now.getSeconds());
	var oFSO = new ActiveXObject("Scripting.FileSystemObject");
	var oFile = oFSO.OpenTextFile(sFilePath, ForAppending, true);
	oFile.WriteLine (Time + " " + sData);
	oFile.Close();

} // WriteToLog
