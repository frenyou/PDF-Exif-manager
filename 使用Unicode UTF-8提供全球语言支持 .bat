::https://answers.microsoft.com/zh-hans/windows/forum/all/%E5%8B%BE%E9%80%89beta%E7%89%88%E4%BD%BF%E7%94%A8u/50de2823-4ff3-43e6-9cef-6d557abd4eac

::Windows 默认以管理员身份运行批处理bat文件
%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit cd /d "%~dp0"

::使用Unicode UTF-8提供全球语言支持 
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Command Processor" /v autorun /t REG_SZ /d "chcp 65001" /f
pause

::当您需要取消这个操作时，您可以尝试在“管理员：命令提示符”中执行以下的命令：
::reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Command Processor" /v autorun

