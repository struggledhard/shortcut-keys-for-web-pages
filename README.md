# shortcut keys for web pages
vbs脚本生成网页快捷键，用于桌面快速打开指定网站或内部系统

1、运行```deskBat.bat```时会在C盘下创建一个文件，把图标从当前文件夹复制到C盘相应文件夹下，这样操作是如果生成快捷键后，删除文件后图标任可读取。

2、用IE打开指定网页  
```
（1）oShellLink.TargetPath改成oShellLink.TargetPath = "C:\Program Files\internet explorer\iexplore.exe"  

（2）oShellLink.TargetPath下面加上oShellLink.Arguments="http://www.baidu.com"  

（3）oShellLink.WorkingDirectory改成oShellLink.WorkingDirectory = "C:\Program Files\internet explorer"  
```
