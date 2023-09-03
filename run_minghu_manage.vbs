gui_dir = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path
CreateObject("WScript.Shell").Run "cmd /c python "+gui_dir+"\MinghuManage.py",0