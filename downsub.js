var config = {
    'splayer':[
        "C:\\Program Files (x86)\\Splayer\\splayer.exe",
        "C:\\Program Files\\SPlayer\\splayer.exe",
        "d:\\tools\\splayer\\splayer.exe"
    ],
    'name':'射手影音播放器',
    'executeTime':30,
    'checkInterval':5,
    'goOver':6
}


var fso = new ActiveXObject("Scripting.FileSystemObject");
var wsh = WScript.CreateObject("WScript.Shell");
var regs = {
    "movie":/(\.avi$)|(\.mkv$)|(\.ts$)/i,
    "sub":/(\.srt$)|(\.ass$)|(\.idx$)|(\.sub$)/i,
    "cd":/(?!-|\.|^)cd\d+(?=\.|-)/i,
    "rar":/\.rar$/i,
    "sample":/((?:\.|-)sample(?=\.))/i,
    "argv":/^-.+/i
}
//默认运行参数
var param = {recursive:false};
var subDir = wsh.Environment('Process').item('APPDATA') + '\\SPlayer\\SVPSub';

//查找射手播放器位置
var l = config.splayer.length;
var player = '';
for (var i = 0; i < l; i++) {
    if (fso.FileExists(config.splayer[i])) {
        player = config.splayer[i];
        break;
    }
}


//WScript.Echo(movieFilter.test(file));
var l = WScript.Arguments.length;
var argv = new Array();
for (var i = 0; i < l; i++) {
    argv[i] = WScript.Arguments(i);
}

var flag = false;
var moviesAll = false;
var moviesLeft = [];

var opts = findByFilter(argv, regs["argv"]);
var l = opts.length;
for (var i = 0; i < l; i++) {
    if (opts[i] == '-r') param['recursive'] = true;
}
var dirs = findByFilter(argv, regs["argv"], true);
var l = dirs.length;
if (dirs.length > 0) {
    for (var i = 0; i < l; i++) {
        for (var j = 0; j < config.goOver; j++)
            flag = exec(dirs[i]);
    }
    print('Done!');
} else {
    print([
        "通过射手播放器下载字幕。", "请在命令行中运行:",
        "  " + WScript.ScriptFullName + " [-r] 视频文件目录",
        '可选参数：',
        "  -r : 对所有下级目录中的视频文件都执行此操作。"
    ].join("\n"));
}


function escapeReg(text) {
    return text.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, "\\$&");
}

function exec(dir) {
    if(moviesAll == false){
        var files = getAllFiles(dir, param['recursive']);
        moviesAll = findByFilter(findByFilter(files, regs.movie), regs.sample, true);
        moviesLeft = moviesAll;
    }
    var movies = moviesLeft;
    moviesLeft = [];
    var l = movies.length;
    for (var i = 0; i < l; i++) {
        var term = false;
        if (!dealWithSub(movies[i])) {
            var cmd = player + " \"" + movies[i] + "\"";
            var wnd = wsh.Exec(cmd);
            var round = Math.ceil(config.executeTime / config.checkInterval);
            for (var j = 0; j < round; j++) {
                WScript.Sleep(config.checkInterval * 1000);
                if (j == 0) wsh.SendKeys("% n");
                if (dealWithSub(movies[i])) {
                    try {
                        wnd.Terminate()
                    } catch (e) {
                    }
                    term = true;
                    break;
                }
            }
            if (!term) try {
                moviesLeft.push(movies[i]);
                wnd.Terminate()
            } catch (e) {
            }
        }
    }
}
/**
 * 处理、检查电影文件的字幕。如果一个字幕文件在射手播放器的appdata目录下，则拷贝
 * 到和电影文件同一目录下。
 *
 * @param string movie 电影文件的位置
 * @return 如果有字幕，则返回true ，否则返回false
 */
function dealWithSub(movie) {
    var retval = false;
    var dir = dirname(movie) + '\\';
    var s = lookingForSub(movie, dir);
    if (s.length > 0) retval = true;
    else {
        s = lookingForSub(movie, subDir);
        if (s.length > 0) {
            var l = s.length;
            for (var i = 0; i < l; i++)  try {
                fso.MoveFile(s[i], dir);
            } catch (e) {
                fso.CopyFile(s[i], dir);
            }
            ;
            retval = true;
        }
    }
    return retval;
}

/**
 * 在给定目录中寻找符合电影文件的字幕
 *
 * @param string movie 电影文件位置
 * @param string dir 查找字幕文件的位置。
 * @return array 字幕文件位置
 */
function lookingForSub(movie, dir) {
    var subs = findByFilter(getAllFiles(dir), regs.sub);
    var re = new RegExp(escapeReg(mainName(movie)));
    return findByFilter(subs, re);
}

function exit() {
    WScript.Quit();
}
function print(v) {
    wsh.PopUp(v.toString(), 0, "字幕下载");
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