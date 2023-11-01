
function pad(sValue)
{
	return sValue < 10 ? '0'+sValue : sValue;
}

function WriteToLog(Data)
{
	var Now = new Date();
	var Time =  pad(Now.getHours()) + ":" + pad(Now.getMinutes()) + ":" + pad(Now.getSeconds()) + ":" + Now.getMilliseconds();
	var msiMessageTypeInfo = 0x04000000;
	var Record;
	Record = Session.Installer.CreateRecord(1);
	Record.StringData(0) = "[1]";
	Record.StringData(1) = "CustomAction.js [" + Time + "]: " + Data;
	Session.Message(msiMessageTypeInfo, Record);

} // WriteToLog
