//============================================================================//
//                              Шифрование пароля                             //
//============================================================================//
function Main()
{
	var Encoder, objFS, objFile, code;

	code = new String();
	code = code.concat("var DBLink;");
	code = code.concat("DBLink = new Array();");
	code = code.concat("DBLink['BACK'] = new Array();");
	code = code.concat("DBLink['BACK']['TNSNAME'] = new String('(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = osma.cgs.sbrf.ru)(PORT = 1521))(CONNECT_DATA = (SID = dbdpc)))');");
	code = code.concat("DBLink['BACK']['USERLOGIN'] = new String('wsback');");
	code = code.concat("DBLink['BACK']['PASSWORD'] = new String('wsback');");
	code = code.concat("DBLink['FRONT'] = new Array();");
	code = code.concat("DBLink['FRONT']['TNSNAME'] = new String('(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = chalna.cgs.sbrf.ru)(PORT = 1521))(CONNECT_DATA = (SERVICE_NAME = uzb_pcod)))');");
	code = code.concat("DBLink['FRONT']['USERLOGIN'] = new String('wsback');");
	code = code.concat("DBLink['FRONT']['PASSWORD'] = new String('wsback');");

	Encoder = WScript.CreateObject("Scripting.Encoder");
	objFS = WScript.CreateObject("Scripting.FileSystemObject");
	objFile = objFS.CreateTextFile("reports.jse", 1);
	objFile.Write(Encoder.EncodeScriptFile(".js", code, 0, "JScript"));
	objFile.Close();

	return 0;
}


{
	Main();
}
