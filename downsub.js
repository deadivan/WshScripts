var config = {
    'splayer':[
        "C:\\Program Files (x86)\\Splayer\\splayer.exe",
        "C:\\Program Files\\SPlayer\\splayer.exe",
        "d:\\tools\\splayer\\splayer.exe"
    ],
    'name':'����Ӱ��������',
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
//Ĭ�����в���
var param = {recursive:false};
var subDir = wsh.Environment('Process').item('APPDATA') + '\\SPlayer\\SVPSub';

//�������ֲ�����λ��
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
        "ͨ�����ֲ�����������Ļ��", "����������������:",
        "  " + WScript.ScriptFullName + " [-r] ��Ƶ�ļ�Ŀ¼",
        '��ѡ������',
        "  -r : �������¼�Ŀ¼�е���Ƶ�ļ���ִ�д˲�����"
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
 * ��������Ӱ�ļ�����Ļ�����һ����Ļ�ļ������ֲ�������appdataĿ¼�£��򿽱�
 * ���͵�Ӱ�ļ�ͬһĿ¼�¡�
 *
 * @param string movie ��Ӱ�ļ���λ��
 * @return �������Ļ���򷵻�true �����򷵻�false
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
 * �ڸ���Ŀ¼��Ѱ�ҷ��ϵ�Ӱ�ļ�����Ļ
 *
 * @param string movie ��Ӱ�ļ�λ��
 * @param string dir ������Ļ�ļ���λ�á�
 * @return array ��Ļ�ļ�λ��
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
    wsh.PopUp(v.toString(), 0, "��Ļ����");
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