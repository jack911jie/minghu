Set objShell = CreateObject("Wscript.Shell")
' 启动 app.py，隐藏命令行窗口
objShell.Run "cmd /c python app.py > C:\Users\admin\Desktop\app_log.txt 2>&1", 0, True

' 等待一段时间，确保 Flask 应用已经启动
WScript.Sleep 2000

' 打开默认浏览器并访问 http://127.0.0.1:5000
objShell.Run "cmd /c start http://127.0.0.1:5000", 0, False