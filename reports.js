//============================================================================//
//                                                      //
//============================================================================//
function FormatDate()
{
	var Result, now, msg;

	Result = new String();

	now = new Date();
	msg = new String();
	msg = msg.concat('0');
	msg = msg.concat(now.getYear() + 0);
	msg = msg.slice(-4);
	Result = Result.concat(msg);
	msg = null;
	msg = new String();
	msg = msg.concat('0');
	msg = msg.concat(now.getMonth() + 1);
	msg = msg.slice(-2);
	Result = Result.concat(msg);
	msg = null;
	msg = new String();
	msg = msg.concat('0');
	msg = msg.concat(now.getDate() + 0);
	msg = msg.slice(-2);
	Result = Result.concat(msg);
	msg = null;
	msg = new String();
	msg = msg.concat('0');
	msg = msg.concat(now.getHours() + 0);
	msg = msg.slice(-2);
	Result = Result.concat(msg);
	msg = null;
	msg = new String();
	msg = msg.concat('0');
	msg = msg.concat(now.getMinutes() + 0);
	msg = msg.slice(-2);
	Result = Result.concat(msg);
	msg = null;
	msg = new String();
	msg = msg.concat('0');
	msg = msg.concat(now.getSeconds() + 0);
	msg = msg.slice(-2);
	Result = Result.concat(msg);
	msg = null;
	now = null;

	return Result;
}

function WriteLog(Message)
{
	var objFS, objFile;

	objFS = new ActiveXObject('Scripting.FileSystemObject');
	objFile = objFS.OpenTextFile('reports.log', 8, 1);
	objFile.WriteLine(FormatDate() + ': ' + Message);
	objFile.Close();
	objFile = null;
	objFS = null;

	return 0;
}

function DeleteFiles(DirName, DaysForKeep)
{
	var objFS, Files, now;

	now = new Date();
	objFS = new ActiveXObject('Scripting.FileSystemObject');
	Files = new Enumerator(objFS.GetFolder(DirName).Files);
	for ((Files.moveFirst()); (!(Files.atEnd())); (Files.moveNext()))
	{
		if ((!(((now - Files.item().DateLastModified) / 86400000) <= DaysForKeep)))
		{
			if((objFS.FileExists(Files.item().Path)))
			{
				objFS.DeleteFile(Files.item().Path);
			}
		}
	}
	Files = null;
	objFS = null;
	now = null;

	return 0;
}

function SendMessage(msgServer, msgFrom, msgTo, msgText)
{
	var objMail, i;

	objMail = new ActiveXObject('CDO.Message');
	objMail.Configuration.Fields.Item('http://schemas.microsoft.com/cdo/configuration/smtpserver') = msgServer;
	objMail.Configuration.Fields.Item('http://schemas.microsoft.com/cdo/configuration/sendusing') = 2;
	objMail.Configuration.Fields.Item('http://schemas.microsoft.com/cdo/configuration/smtpserverport') = 25;
	objMail.Configuration.Fields.Item('http://schemas.microsoft.com/cdo/configuration/smtpauthenticate') = 2;
	objMail.Configuration.Fields.Update();
	objMail.To = msgTo;
	objMail.From = msgFrom;
	objMail.Subject = 'Report';
	objMail.HTMLBody = msgText;
	objMail.Send();
	objMail = null;

	return 0;
}

function LoadXML(FileName)
{
	var Result, objXML, str;

	objXML = new ActiveXObject('MSXML2.DOMDocument');
	objXML.load(FileName);
	if((!(objXML.parseError.errorCode == 0)))
	{
		str = new String();
		str = str.concat('File:');
		str = str.concat(' ');
		str = str.concat(objXML.parseError.url);
		str = str.concat('\n');
		str = str.concat('Code:');
		str = str.concat(' ');
		str = str.concat(objXML.parseError.errorCode);
		str = str.concat('\n');
		str = str.concat('Message:');
		str = str.concat(' ');
		str = str.concat(objXML.parseError.reason);
//		str = str.concat('\n');
		str = str.concat('Source:');
		str = str.concat(' ');
		str = str.concat(objXML.parseError.srcText);
		str = str.concat('\n');
		str = str.concat('Line:');
		str = str.concat(' ');
		str = str.concat(objXML.parseError.line);
		str = str.concat('\n');
		str = str.concat('Column:');
		str = str.concat(' ');
		str = str.concat(objXML.parseError.linepos);
//		str = str.concat('\n');
		throw new Error(str);
	}
	Result = objXML.documentElement;
	objXML = null;

	return Result;
}

function SQL2XML(str, sql)
{
	var Result, objConn;

	Result = new ActiveXObject('MSXML2.DOMDocument');

	objConn = new ActiveXObject('ADODB.Connection');
	objConn.ConnectionString = str;
	objConn.Open();
	objConn.Execute(sql).Save(Result, 1);
	objConn.Close();
	objConn.ConnectionString = '';
	objConn = null;

	return Result;
}

function XSL4XML(objXML, objXSL)
{
	var Result;

	Result = new ActiveXObject('MSXML2.DOMDocument');
	objXML.transformNodeToObject(objXSL, Result);

	return Result;
}

function XML2ZIP(objXML, FileName, ArchName)
{
	var objZIP, objFS, i;

	objFS = new ActiveXObject('Scripting.FileSystemObject');
	objZIP = new ActiveXObject('XStandard.Zip');
	objXML.save(FileName);
	objZIP.Pack(FileName, ArchName, 0);
	objFS.DeleteFile(FileName);
	objFS = null;
	objZIP = null;

	return 0;
}

function getReportParamByTask(objXML, Task)
{
	var Result, ID, Nodes;

	Result = new Array();

	WriteLog('    ID = ' + Task['Report'] + '.');
	objXML.ownerDocument.setProperty('SelectionLanguage', 'XPath');
	Result['BASE'] = new Array();
	Nodes = new Enumerator(objXML.selectNodes('/root/report'));
	for((Nodes.moveFirst()); (!(Nodes.atEnd())); (Nodes.moveNext()))
	{
		if((Nodes.item().getAttribute('id') == Task['Report']))
		{
			if((Nodes.item().selectSingleNode('@id') == null))
			{
				ID = new String();
			}
			else
			{
				ID = Nodes.item().selectSingleNode('@id').text;
			}
			Result['BASE'][ID] = new Array();
			if((Nodes.item().selectSingleNode('@name') == null))
			{
				Result['BASE'][ID]['NAME'] = new String();
			}
			else
			{
				Result['BASE'][ID]['NAME'] = Nodes.item().selectSingleNode('@name').text;
			}
		}
		
	}
	Nodes = null;
	Result['EXEC'] = new Array();
	Nodes = new Enumerator(objXML.selectNodes('/root/report/batch/exec'));
	for((Nodes.moveFirst()); (!(Nodes.atEnd())); (Nodes.moveNext()))
	{
		if((Nodes.item().parentNode.parentNode.getAttribute('id') == Task['Report']))
		{
			if((Nodes.item().selectSingleNode('id') == null))
			{
				ID = new String();
			}
			else
			{
				ID = Nodes.item().selectSingleNode('id').text;
			}
			Result['EXEC'][ID] = new Array();
			if((Nodes.item().selectSingleNode('type') == null))
			{
				Result['EXEC'][ID]['TYPE'] = new String();
			}
			else
			{
				Result['EXEC'][ID]['TYPE'] = Nodes.item().selectSingleNode('type').text;
			}
			if((Nodes.item().selectSingleNode('name') == null))
			{
				Result['EXEC'][ID]['NAME'] = new String();
			}
			else
			{
				Result['EXEC'][ID]['NAME'] = Nodes.item().selectSingleNode('name').text;
			}
			if((Nodes.item().selectSingleNode('command') == null))
			{
				Result['EXEC'][ID]['COMMAND'] = new String();
			}
			else
			{
				Result['EXEC'][ID]['COMMAND'] = Nodes.item().selectSingleNode('command').text;
			}
			if((Nodes.item().selectSingleNode('dblink') == null))
			{
				Result['EXEC'][ID]['DBLINK'] = new String();
			}
			else
			{
				Result['EXEC'][ID]['DBLINK'] = Nodes.item().selectSingleNode('dblink').text;
			}
		}
	}
	Nodes = null;
	Result['PARAM'] = new Array();
	Nodes = new Enumerator(objXML.selectNodes('/root/report/params/param'));
	for((Nodes.moveFirst()); (!(Nodes.atEnd())); (Nodes.moveNext()))
	{
		if((Nodes.item().parentNode.parentNode.getAttribute('id') == Task['Report']))
		{
			if((Nodes.item().selectSingleNode('id') == null))
			{
				ID = new String();
			}
			else
			{
				ID = Nodes.item().selectSingleNode('id').text;				
			}
			Result['PARAM'][ID] = new Array();
			if((Nodes.item().selectSingleNode('name') == null))
			{
				Result['PARAM'][ID]['NAME'] = new String();
			}
			else
			{
				Result['PARAM'][ID]['NAME'] = Nodes.item().selectSingleNode('name').text;
			}
			if((Nodes.item().selectSingleNode('type') == null))
			{
				Result['PARAM'][ID]['TYPE'] = new String();
			}
			else
			{
				Result['PARAM'][ID]['TYPE'] = Nodes.item().selectSingleNode('type').text;
			}
			if((Nodes.item().selectSingleNode('required') == null))
			{
				Result['PARAM'][ID]['REQUIRED'] = new String();
			}
			else
			{
				Result['PARAM'][ID]['REQUIRED'] = Nodes.item().selectSingleNode('required').text;
			}
		}
	}
	Nodes = null;
	for(i in Result['PARAM'])
	{
		Result['PARAM'][i]['VALUE'] = Task[i];
		switch(Result['PARAM'][i]['REQUIRED'])
		{
			case '0':
			{
				if((Result['PARAM'][i]['VALUE'] == null))
				{
					switch(Result['PARAM'][i]['TYPE'])
					{
						case 'date':
						{
							Result['PARAM'][i]['VALUE'] = '01.01.1800';
						}
						break;
						case 'string':
						{
							Result['PARAM'][i]['VALUE'] = '00000000000000000000';
						}
						break;
						case 'number':
						{
							Result['PARAM'][i]['VALUE'] = '0';
						}
						break;
						default:
						{
							Result['PARAM'][i]['VALUE'] = '0';
						}
						break;
					}
					WriteLog('  ' + i + '   ID = ' + Task['Report'] + '        : ' + Result['PARAM'][i]['VALUE'] + '.');
				}
			}
			break;
			case '1':
			{
				if((Result['PARAM'][i]['VALUE'] == null))
				{
					throw new Error('  "' + i + '"   ID = ' + Task['Report'] + '  .');
				}
			}
			break;
			default:
			{
				throw new Error('    "required"      ID = ' + Task['Report'] + '.');
			}
			break;
		}
	}
	if((Result['PARAM']['__BRANCH__'] == null))
	{
		ID = Task['NORMDEP'].split('-');
		Result['PARAM']['__BRANCH__'] = new Array();
		Result['PARAM']['__BRANCH__']['NAME'] = ' :';
		Result['PARAM']['__BRANCH__']['TYPE'] = 'string';
		Result['PARAM']['__BRANCH__']['VALUE'] = ID[0];
		Result['PARAM']['__BRANCH__']['REQUIRED'] = '0';
		ID = null;
	}
	if((Result['PARAM']['__OFFICE__'] == null))
	{
		ID = Task['NORMDEP'].split('-');
		Result['PARAM']['__OFFICE__'] = new Array();
		Result['PARAM']['__OFFICE__']['NAME'] = ' :';
		Result['PARAM']['__OFFICE__']['TYPE'] = 'string';
		Result['PARAM']['__OFFICE__']['VALUE'] = ID[1];
		Result['PARAM']['__OFFICE__']['REQUIRED'] = '0';
		ID = null;
	}
	WriteLog('    ID = ' + Task['Report'] + '  .');

	return Result;
}

function LoadConfig(FileName)
{
	var objFS, objFile, Result, i, arr;

	Result = new Array();

	objFS = new ActiveXObject('Scripting.FileSystemObject');
	objFile = objFS.OpenTextFile(FileName, 1, 0);
	for((i = 0); (!(objFile.AtEndOfStream)); (i += 1))
	{
		Line = objFile.ReadLine();
		arr = Line.split('=');
		Result[arr[0]] = arr[1];
	}
	objFile = null;
	objFS = null;

	return Result;
}

function ReadFile(FileName)
{
	var Result, objFS, objFile;

	objFS = new ActiveXObject('Scripting.FileSystemObject');
	objFile = objFS.OpenTextFile(FileName, 1, 0);
	Result = objFile.ReadAll();
	objFile.Close();
	objFile = null;
	objFS = null;

	return Result;
}

function WriteFile(FileName, str)
{
	var Result, objFS, objFile;

	objFS = new ActiveXObject('Scripting.FileSystemObject');
	objFile = objFS.OpenTextFile(FileName, 2, 1);
	objFile.Write(str);
	objFile.Close();
	objFile = null;
	objFS = null;

	return 0;
}

function getTaskFromFile(FileName)
{
	var Result, Line, objFS, str, arr, i, k;

	Result = new Array();

	objFS = new ActiveXObject('Scripting.FileSystemObject');
	objFile = objFS.OpenTextFile(FileName, 1, 0);
	for((i = 0); (!(objFile.AtEndOfStream)); (i += 1))
	{
		line = objFile.ReadLine();
		str = line.split('|');
		for(k in str)
		{
			arr = str[k].split(':');
			if((!(Result[arr[0]])))
			{
				Result[arr[0]] = arr[1];
			}
			arr = null;
		}
		str = null;
		line = null;
	}
	objFile.Close();
	if((Result['NAME'] == null))
	{
		Result['NAME'] = objFS.GetFileName(FileName);
	}
	if((Result['ARCH'] == null))
	{
		Result['ARCH'] = FormatDate() + '.' + 'zip';
	}
	objFile = null;
	objFS = null;

	return Result;
}

function ExecuteReportFile(objXML, objXSL, Config, Task)
{
	var Report, Message;

	WriteLog('     ID = ' + Task['Report'] + '.');
	Report = getReportParamByTask(objXML, Task);
	Message = new String();
	Message = Message.concat('DOMAIN:');
	Message = Message.concat(Task['DOMAIN']);
	Message = Message.concat('|');
	Message = Message.concat('FILENAME:')
	Message = Message.concat((Task['ARCH']));
	Message = Message.concat('|');
	Message = Message.concat('REPORTNAME:');
	Message = Message.concat(Report['BASE'][Task['Report']]['NAME']);
	for(i in Report['PARAM'])
	{
		if((Report['PARAM'][i]['REQUIRED'] == '1'))
		{
			Message = Message.concat(' ');
			Message = Message.concat(Report['PARAM'][i]['NAME']);
			Message = Message.concat(' ');
			Message = Message.concat(Report['PARAM'][i]['VALUE']);
		}
	}
	Message = Message.concat('|');
	WriteFile(Config['HOME'] + '/' + 'OUT' + '/' + Task['NAME'], Message);
	Message = null;
	Report = null;
	WriteLog('     ID = ' + Task['Report'] + '  .');

	return 0;
}

function ExecuteReportTask(objXML, objXSL, Config, Task)
{
	var Report, str, sql, i, k;

	WriteLog('     ID = ' + Task['Report'] + '.');
	Report = getReportParamByTask(objXML, Task);
	for(k in Report['EXEC'])
	{
		if((Report['EXEC'][k]['TYPE'] == 'sql'))
		{
			WriteLog(' SQL-   ID = ' + Task['Report'] + '.');
			sql = ReadFile(Config['HOME'] + '/' + 'SQL' + '/' + Report['EXEC'][k]['COMMAND']);
			for(i in Report['PARAM'])
			{
				sql = sql.replace(new RegExp(i, 'g'), Report['PARAM'][i]['VALUE']);
			}
			WriteLog(' SQL-   ID = ' + Task['Report'] + '  .');
			WriteLog('SQL-   ID = ' + Task['Report'] + ':');
			WriteLog(sql);
			WriteLog(' ConnectionString   ID = ' + Task['Report'] + '.');
			str = new String();
			str = str.concat('Provider=MSDAORA;Data Source=TNSNAME;User ID=USERLOGIN;Password=PASSWORD');
			for(i in DBLink[Report['EXEC'][k]['DBLINK']])
			{
				str = str.replace(new RegExp(i, 'g'), DBLink[Report['EXEC'][k]['DBLINK']][i]);
			}
			WriteLog(' ConnectionString   ID = ' + Task['Report'] + '  .');
			WriteLog('ConnectionString   ID = ' + Task['Report'] + ':');
			WriteLog(str);
			WriteLog('  ' + Report['EXEC'][k]['NAME'] + '   ID = ' + Task['Report'] + '.');
			XML2ZIP(XSL4XML(SQL2XML(str, sql), objXSL), Config['HOME'] + '/' + 'ARHIV' + '/' + Report['EXEC'][k]['NAME'], Config['HOME'] + '/' + 'ARHIV' + '/' + Task['ARCH']);
			WriteLog('  ' + Report['EXEC'][k]['NAME'] + '   ID = ' + Task['Report'] + '  .');
			str = null;
			sql = null;
		}
	}
	Report = null;
	WriteLog('     ID = ' + Task['Report'] + '  .');

	return 0;
}

function ExecuteAllReports(objXML, objXSL, Config)
{
	var objFS, FileName, Files, Name;

	objFS = new ActiveXObject('Scripting.FileSystemObject');
	WriteLog('     ' + Config['HOME'] + '/' + 'OUT' + '.');
	DeleteFiles(Config['HOME'] + '/' + 'OUT', Config['HOWLONGKEEPFILES']);
	WriteLog('     ' + Config['HOME'] + '/' + 'OUT' + '  .');
	if(!(objFS.GetFolder(Config['HOME'] + '/' + 'INQUEUE').Files.Count == 0))
	{
		throw new Error('    .');
	}
	Files = new Enumerator(objFS.GetFolder(Config['HOME'] + '/' + 'IN').Files);
	for((Files.moveFirst()); (!(Files.atEnd())); (Files.moveNext()))
	{
		Name = objFS.GetFileName(Files.item().Name);
		if((objFS.FileExists(Config['HOME'] + '/' + 'INQUEUE' + '/' + Name)))
		{
			objFS.DeleteFile(Config['HOME'] + '/' + 'INQUEUE' + '/' + Name);
		}
		objFS.MoveFile(Config['HOME'] + '/' + 'IN' + '/' + Name, Config['HOME'] + '/' + 'INQUEUE' + '/' + Name);
		try
		{
			Task = getTaskFromFile(Config['HOME'] + '/' + 'INQUEUE' + '/' + Name);
			SendMessage(Config['SERVER'], Config['ADMIN'], Task['MAIL'], '    .     OutLook       .');
			ExecuteReportTask(objXML, objXSL, Config, Task);
			ExecuteReportFile(objXML, objXSL, Config, Task);
			SendMessage(Config['SERVER'], Config['ADMIN'], Task['MAIL'], '<p>!</p><p>      <a href=""></a>.</p>');
			Task = null;
		}
		catch(ex)
		{
			if((objFS.FileExists(Config['HOME'] + '/' + 'ERRORS' + '/' + Name)))
			{
				objFS.DeleteFile(Config['HOME'] + '/' + 'ERRORS' + '/' + Name);
			}
			objFS.CopyFile(Config['HOME'] + '/' + 'INQUEUE' + '/' + Name, Config['HOME'] + '/' + 'ERRORS' + '/' + Name);
			WriteLog('   ' + Name + '  :');
			WriteLog(ex.message);
			try
			{
				SendMessage(Config['SERVER'], Config['ADMIN'], Task['MAIL'], '<p>   ' + Name + '  :</p>' + '<p>' + ex.message + '</p>' + '<p>,           .</p>');
			}
			catch(ex)
			{
				WriteLog('      ' + Name + '  :');
				WriteLog(ex.message);
			}
		}
		if((objFS.FileExists(Config['HOME'] + '/' + 'TEMP' + '/' + Name)))
		{
			objFS.DeleteFile(Config['HOME'] + '/' + 'TEMP' + '/' + Name);
		}
		objFS.MoveFile(Config['HOME'] + '/' + 'INQUEUE' + '/' + Name, Config['HOME'] + '/' + 'TEMP' + '/' + Name);
		Name = null;
	}
	Files = null;
	if(!(objFS.GetFolder(Config['HOME'] + '/' + 'INQUEUE').Files.Count == 0))
	{
		throw new Error('   .');
	}
	WriteLog('     ' + Config['HOME'] + '/' + 'ARHIV' + '.');
	DeleteFiles(Config['HOME'] + '/' + 'ARHIV', Config['HOWLONGKEEPFILES']);
	WriteLog('     ' + Config['HOME'] + '/' + 'ARHIV' + '  .');
	objFS = null;

	return 0;
}

function Main()
{
	var objXML, objXSL, Config;

	WriteLog('Start process.');
	try
	{
		Config = LoadConfig('reports.cfg');
		objXML = LoadXML(Config['HOME'] + '/' + 'reports.xml');
		objXSL = LoadXML(Config['HOME'] + '/' + 'reports.xsl');
		ExecuteAllReports(objXML, objXSL, Config);
		objXML = null;
		objXSL = null;
	}
	catch(ex)
	{
		WriteLog('    :');
		WriteLog(ex.message);
	}
	WriteLog('Stop process.');

	return 0;
}

{
	Main();	
}
