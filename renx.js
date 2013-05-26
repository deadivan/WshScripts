var fso = new ActiveXObject("Scripting.FileSystemObject");
var wsh = WScript.CreateObject("WScript.Shell");

function showHelp(){
	print([
		"",
	].join("\n"));
}
var l = WScript.Arguments.length;
var argv = new Array();
for (var i = 0; i < l; i++) {
    argv[i] = WScript.Arguments(i);
}
var f = new RegExp(argv[0], 'g');
var r = argv[1] || "";
var files = findByFilter(getAllFiles(".\\"), f);
var fl = files.length;
print(f);
for(var i = 0; i < fl; i++){
	var dest = dirname(files[i]) + "\\" + basename(files[i]).replace(f, r);
 	fso.moveFile(files[i], dest);
}
function exit() {
    WScript.Quit();
}
function print(v) {
    wsh.PopUp(v.toString(), 0, "");
}
function findByFilter(ary, reg, invert) {
    var invert = invert || false;
    var retval = Array();
    var i = 0;
    for (var t in ary) {
        if (invert != reg.test(ary[t])) {
            retval[i++] = ary[t];
        }
        //echo([invert, reg.test(ary[t]), reg, ary[t]]);
    }
    return retval;
}
function getAllFiles(v, recursive){
    var recursive = recursive || false;
    var retval = Array();
    if(!fso.FolderExists(v)) return retval;
    var t = fso.GetFolder(v);
    if(t.files){
        var folder = new Enumerator(t.files);
        for(var i = 0; !folder.atEnd(); folder.moveNext(), i++){
            retval[i] = String(folder.item());
        }
    }
    if(recursive == true){
        var folder = new Enumerator(t.SubFolders);
        for(var i = 0; !folder.atEnd(); folder.moveNext(), i++){
            retval = retval.concat(getAllFiles(String(folder.item()), true));
        }
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