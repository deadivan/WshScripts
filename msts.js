/**
* �缯����
*Movie Subtitle Filename Sync
*Author deadivan[ A t ]gmail.com
*If you meet WScript " ** DoOpenPipeStream **" error message.The scrrun.dll  may not   
*registered properly.You need re-register it.
*From Start/Run, Type
*regsvr32 %SystemRoot%\system32\scrrun.dll
*It should be ok.
*/


var fso = new ActiveXObject("Scripting.FileSystemObject");
var wsh = WScript.CreateObject("WScript.Shell");

var regs = {
	"movie" 	: /(\.avi$)|(\.mkv$)|(\.ts$)/i,
	"sub"	: /(\.srt$)|(\.idx$)|(\.sub$)/i,
	"cd"		: /(?!-|\.|^)cd\d+(?=\.|-)/i,
	"rar"	: /\.rar$/i,
	"sample"	: /((?:\.|-)sample(?=\.))/i,
	"argv"	: /^--.+/i,
	"movie_id" : [/s0(\d)e(\d{2})/i , /(\d)x(\d{2})/i]
}
var subFormats = ["%S%%E%", "s0%S%e%E%", "%S%x%E%"];

//WScript.Echo(movieFilter.test(file));
var l = WScript.Arguments.length;
var argv = new Array();
for(var i = 0; i < l;i ++){
	argv[i] = WScript.Arguments(i);
}

var flag = false;

var opts = findByFilter(argv, regs["argv"]);
var dirs = findByFilter(argv, regs["argv"], true);

if(argv.length == 0){
	showHelp();
}else{
	if(opts.length > 0){
		//�������ֻʹ�õ�һ������
		switch(String(opts[0]).toLowerCase()){
			case "--install":
			install();
			exit();
			break;
			case "--uninstall":
			uninstall();
			exit();
			break;
			case "--help":
			default:
			showHelp();
			exit();
			break;
		}
	}
	var l = dirs.length;
	for(var i = 0;i < l;i ++){
		flag = tt(dirs[i]);
	}
	if(flag){
		echo ("�����ɹ�!");
	}else{
		echo ("ʧ��!");
		showHelp();
	}
}

function showHelp(){
	echo("��Ļ���� [Movie Subtitle Filename Sync] v0.1.0 beta\n\n" +
	"��һ����������DVDRipĿ¼�е���Ļ�ļ���������ͬĿ¼�еĵ�Ӱ�ļ������޸ġ�\n\n" +
	"msfs.js [ --install | --uninstall | --help ] names\n\n" +
	"  names      \tָ��һ������������Ҫ�����Ŀ¼���б�\n" +
	"  --install  t��װmsfs.js��WindowsĿ¼�����ҽ�\"��Ļ����\"���ܼ����Ҽ��˵���\n" +
	"  --uninstall\tɾ��WindowsĿ¼�е�msfs.js�������Ƴ��Ҽ��˵��е�\"��Ļ����\"���ܡ�\n" +
	"  --help     \t��ʾ��������Ϣ��\n\n" +
	"�����װ�˱����������ѡ��һ����������DVDRipĿ¼���Ҽ�ѡ��\"��Ļ����\"�����и���\n������\n" +
	"�����ʹ���κβ����������������ʾ������Ϣ��\n" +
	"���DVDRipĿ¼����rar�ļ���Ҳ�Ὣrar�ļ���ѹ���ٽ��и����������������������Ҫ\nϵͳ�а�װWinRar��\n\n" +
	"ע�⣺������������ھ缯���߸����Ӱ����Ļ������");
}

function echo(v){
	wsh.PopUp(v.toString(), 0, "��Ļ����");
}

function exit(){
	WScript.Quit();
}

function install(){
	var env = wsh.Environment("PROCESS");
	var dest = env("SYSTEMROOT")+ "\\msfs.js";
	wsh.RegWrite("HKCR\\Directory\\shell\\msfs\\", 
		"��Ļ����");
	wsh.RegWrite("HKCR\\Directory\\shell\\msfs\\command\\", 
		"WScript " + dest + " %1");
	fso.CopyFile(WScript.ScriptFullName, dest);
	echo("��װ�ɹ���");
}

function uninstall(){
	var env = wsh.Environment("PROCESS");
	var dest = env("SYSTEMROOT")+ "\\msfs.js";
	try{
		wsh.RegDelete("HKCR\\Directory\\shell\\msfs\\command\\");
		wsh.RegDelete("HKCR\\Directory\\shell\\msfs\\");
		fso.DeleteFile(dest);
	}catch(e){
	}
	echo("�ɹ�ж��!");
}

function tt(v){
	var f = false;
	var af = getAllFiles(v);
	var cf = {
		"movie"	: findByFilter(af, regs["movie"]),
		"rar"	: findByFilter(af, regs["rar"])
	};
	var cfrl = cf["rar"].length;
	var t = true;
	for(var i = 0;i < cfrl;i ++){
		t = t && unrar(cf["rar"][i], dirname(cf["rar"][i]));
	}
	if(!t){
		if(7 == wsh.Popup("��ѹ���ļ���ʱ������˴����Ƿ������������", 0, "��Ļ����", 4 + 48)){
			WScript.Quit(0);
		};
	}
	af =  getAllFiles(v);
	cf["sub"] =	findByFilter(af, regs["sub"]);

	var cfml = cf["movie"].length;
	var t = findByFilter(cf["movie"], regs["sample"], true);
	f = syncName(t, cf["sub"]);
	return f;
}

/**
*@param mov Array movie files 
*@param sub Array sub files
*/
function syncName(mov, sub){
	var l = mov.length;
//	var r = regs["cd"];//.compile(regs["cd"].source, "i");
	var r = regs["movie_id"];
	var f = false;

	if(l > 1){
		var tm = null;
		for(var i = 0; i < l; i++){
			for(idx in r){
				tm = mov[i].match(r[idx]);
				if(tm) break;
			}
			for(idx in subFormats){
				var subFormat = subFormats[idx];
				var tre = String(subFormat).replace("%S%", tm[1]);
				tre = tre.replace("%E%", tm[2]);
				var tre = new RegExp(tre, "i");
				var as = findByFilter(sub, tre);
				f = renameSub(mov[i], as);
			}
		}
	}else if(l == 1){
		f = renameSub(mov[0], sub);
	}
	return f;
}
/**
*@param mov String  moviefile name
*@param sub Array sub files name
*/
function renameSub(mov, sub){
	var l = sub.length;
	for(var i = 0;i < l; i++){
		var ext = "." + extName(sub[i]);
		 
		//keep .gb.srt .eng.srt .big5.srt 's lang info
		var ext2 = "." + extName(mainName(sub[i]));
		if(ext2.length > 8 || ext2 == ".") ext2 = "";
		//echo([ext, ext2]);
		//max is ".chinese"
		
		var dest = dirname(mov) + "\\" + mainName(mov) + 
		String(ext2 + ext).replace(regs["cd"], "").replace(/\.+/, ".");
		if(sub[i] == dest) continue;
		var j = 1;
		while(fso.FileExists(dest)){
			dest = dirname(mov) + "\\" + mainName(mov) + ".(" + (j ++) + ")" +
				String(ext2 + ext).replace(regs["cd"], "").replace(/\.+/, ".");
		}
		fso.MoveFile(sub[i],dest);
	}
	return true;
}

function unrar(f, d){
	var wsh = WScript.CreateObject("WScript.Shell");
	try{
		var rarPath = wsh.RegRead("HKLM\\\SOFTWARE\\Microsoft\\" + 
	"Windows\\CurrentVersion\\App Paths\\WinRAR.exe\\Path");
	}catch(e){
		WScript.Echo("û�а�װWinRar!�޷���ѹ��rar�ļ�!");
		return false;
	}
	var cmd = "\"" + rarPath + "\\unrar.exe\" e -y " + " \"" + 
	String(f).replace(/\"/g, "\\\"") + "\" \"" +
	String(d).replace(/\"/g, "\\\"") + "\" ";
	return (wsh.Run(cmd, 7, true) == 0);
}

function basename(s){
	var retval = s;
	if(/\\/.test(s)){
		retval = String(s).split(/\\/).slice(-1).toString();
	}
	return retval;
}


function dirname(s){
	var retval = "";
	if(/\\/.test(s)){
		retval = String(s).split(/\\/).slice(0, -1).join("\\");
	}
	return retval;
}

function mainName(s){
	var retval = "";
	var l = basename(s)
	if(l.length){
		if(/\./.test(l)){
			retval = String(l).split(".").slice(0, -1).join(".");
		}
	}
	return retval;
}
function extName(s){
	var retval = "";
	var l = basename(s)
	if(l.length){
		if(/\./.test(l)){
			retval = String(l).split(".").slice(-1).toString();
		}
	}
	return retval;
}

function findByFilter(ary, reg, invert){
	var invert = invert || false;
	var retval = Array();var i = 0;
	for(var t in ary){
		if(invert != reg.test(ary[t])){
			retval[i ++] = ary[t];
		}
		//echo([invert, reg.test(ary[t]), reg, ary[t]]);
	}
	return retval;
}

function getAllFiles(v){
	var t = fso.GetFolder(v);
	var retval = Array();
	var folder = new Enumerator(t.files);
	for(var i = 0; !folder.atEnd(); folder.moveNext(), i++){
		retval[i] = String(folder.item());
	}
	return retval;
} 