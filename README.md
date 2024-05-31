使用ahk制作的pdf管理器。
windows only，仅适用windows系统。



介绍：

预制常用ExifTool命令，在此基础上制作了可视化窗口。可显示、修改PDF元数据。
利用mutool生成PDF封面图像文件。
可调用裁剪程序裁剪封面图像。

在ansi编码下，exiftool输出中文会乱码，需要先设置win系统的unicode支持。



使用方法：

下载解压后，先运行bat以使用unicode。首次运行主程序会自动解压exiftool和mutool及裁剪图像等程序。



注意：

命令行运行exiftool后，会利用剪贴板传递输出给主程序。因此读取pdf和保存pdf元数据时，不要进行复制粘贴操作。
命令行用其他方法传递输出会在屏幕闪动一下，使用系统剪贴板则不会。暂时还没找到更好的方法。



运行截图：
![运行截图](https://github.com/frenyou/PDF-Exif-manager/assets/101919925/7ae0e20b-41a6-40bc-81a7-913bebc8b9fe)
