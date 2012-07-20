/**
* 剧集改名
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
		//多个参数只使用第一个参数
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
		echo ("改名成功!");
	}else{
		echo ("失败!");
		showHelp();
	}
}

function showHelp(){
	echo("字幕改名 [Movie Subtitle Filename Sync] v0.1.0 beta\n\n" +
	"将一个或者数个DVDRip目录中的字幕文件名，根据同目录中的电影文件名做修改。\n\n" +
	"msfs.js [ --install | --uninstall | --help ] names\n\n" +
	"  names      \t指定一个或者数个需要处理的目录的列表。\n" +
	"  --install  t安装msfs.js到Windows目录，并且将\"字幕改名\"功能加入右键菜单。\n" +
	"  --uninstall\t删除Windows目录中的msfs.js，并且移除右键菜单中的\"字幕改名\"功能。\n" +
	"  --help     \t显示本帮助信息。\n\n" +
	"如果安装了本软件，可以选中一个或者数个DVDRip目录，右键选择\"字幕改名\"来进行改名\n操作。\n" +
	"如果不使用任何参数，本软件将会显示帮助信息。\n" +
	"如果DVDRip目录中有rar文件，也会将rar文件解压缩再进行改名操作，但是这个功能需要\n系统中安装WinRar。\n\n" +
	"注意：本软件不能用于剧集或者高清电影的字幕改名。");
}

function echo(v){
	wsh.PopUp(v.toString(), 0, "字幕改名");
}

function exit(){
	WScript.Quit();
}

function install(){
	var env = wsh.Environment("PROCESS");
	var dest = env("SYSTEMROOT")+ "\\msfs.js";
	wsh.RegWrite("HKCR\\Directory\\shell\\msfs\\", 
		"字幕改名");
	wsh.RegWrite("HKCR\\Directory\\shell\\msfs\\command\\", 
		"WScript " + dest + " %1");
	fso.CopyFile(WScript.ScriptFullName, dest);
	echo("安装成功！");
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
	echo("成功卸载!");
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
		if(7 == wsh.Popup("解压缩文件的时候出现了错误，是否继续改名程序？", 0, "字幕改名", 4 + 48)){
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
		WScript.Echo("没有安装WinRar!无法解压缩rar文件!");
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