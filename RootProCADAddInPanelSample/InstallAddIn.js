if (WScript.Arguments.length != 2) WScript.Quit(-1);

var wshShell = new ActiveXObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");

var myDocs = wshShell.SpecialFolders("MyDocuments");
var targetPath = fso.BuildPath(myDocs, WScript.Arguments(1));
var command = "xcopy \"" + WScript.Arguments(0) + "\" \"" + targetPath + "\\\" /y"
var exec = wshShell.Exec(command);
while (exec.Status == 0) WScript.Sleep(100);
WScript.Echo("Add in installed to " + targetPath);
