; 脚本名称：PDF元数据管理器

; 脚本用途：预制常用ExifTool命令，在此基础上制作了可视化窗口，
;           可显示、修改PDF元数据。生成和裁剪PDF封面图像文件。

; 脚本用法：点击Listview窗格，输入信息后点击保存。
;           保存元数据时，Exiftool会将出版日期的年月日，统一用“:”间隔。

; ===============================================================================
; 注意！Exiftool对中文支持不佳，命令行读取和输出中文时会乱码。
; 从0.56开始，勾选系统区域设置的“use utf-8 for worldwide language support”
; 不再使用临时文本文件作为中转，不再生成及删除中转文件。

; 注意！Exiftool总是覆盖原来的文件，而不只是改动某些部位。
; Exiftool always rewrites the file.  There is no edit in place functionality. 
; https://exiftool.org/forum/index.php?msg=82132
; -overwrite_original              Overwrite original by renaming tmp file
; -overwrite_original_in_place     Overwrite original by copying tmp file

; 搜索「exif项目相关命令」定位代码行，方便增删exif项目。
; ===============================================================================

; 参考资料↓————————————————————————————————

; 修改listview网格 Class_LV_InCellEdit.ahk
; https://github.com/AHK-just-me/Class_LV_InCellEdit

; 获取选中的列数
; https://www.autohotkey.com/board/topic/80265-solved-which-column-is-clicked-in-listview/

; 更改dropdownlist背景颜色 Class_OD_Colors.ahk
; https://www.autohotkey.com/boards/viewtopic.php?t=338

; 更改combobox背景颜色 Class_CtlColors.ahk
; https://www.autohotkey.com/boards/viewtopic.php?f=6&t=2197
; by Justme

; 输出pdf页面为图像命令示例
; mutool convert -o output.jpg -O width=800 -r 300 1.pdf 1    输出output.jpg 宽度800 dpi300 输入1.pdf 转换第1页
; mutool draw -o output.png -r 96 1.pdf 1
; https://mupdf.readthedocs.io/en/latest/mutool-convert.html

; 参考资料↑————————————————————————————————

SetBatchLines -1
SetWorkingDir %A_ScriptDir%
#NoEnv
#Include Class_LV_InCellEdit.ahk
#Include Class_OD_Colors.ahk
#Include Class_CtlColors.ahk
#Include Class_StdOutToVar.ahk
#SingleInstance Off
#NoTrayIcon
#If WinActive("ahk_id" GUIMainHD)     ; 当程序激活时启用热键


; 常用可改变量↓————————————————————————————————

; v0.68
; 修改功能，使用新的类传递命令行输出，不再占用剪贴板。

; 热键列表 
; Del 删除文件
; Ctrl+a 全选行
; z 复制元数据（仅当启用时）
; x 粘贴元数据（仅当启用时）

; 应用名称
appname1 = PDF元数据管理器_v0.68

; 类型文件后缀名
TypeFileExt := "pdf"

; 显示缩放DPI值（跟随系统缩放DPI值为96）
GuiDPI = 96

; 是否覆盖原文件
; 屏蔽变量将备份源文件
OverWrite = -overwrite_original_in_place

; 读取出版社列表
gosub, PublisherListLB

; 重命名模版(丛书_书名_作者_出版商_出版日期)
ReNamePattern := "BookSeries_BookTitle_BookAuthor_BookPublisher_BookPubDate"
Sub3EditLinkVar := "_"  ; 字段连接符
; 年月日位数，默认为年份4位数，不用另外设置

; exiftool的cmd路径
ExiftoolPath := A_ScriptDir "\exiftool"  ; exiftool的文件名须是exiftool.exe

; mutool的cmd路径
MutoolPath := A_ScriptDir "\mutool"  ; mutool的文件名须是mutool.exe

; 图像裁剪程序（自定义程序请修改所有CropImage相关变量）
; 程序名称
CropImageTool := "图像裁剪工具v2.0.exe"   ; 注意fileinstall命令处要同步修改
; 程序路径
CropImageToolPath := A_ScriptDir "\" CropImageTool
; 调用程序路径
CropImageCallPath := A_ScriptDir "\命令行调用裁剪工具.exe"

; 预设替换字符串
BatchReplaceStr1=
(Join
￠￠￠^ ￠￠-1￠（￠(￠-1￠）￠)￠-1
￠^\(￠[￠-1￠\)￠]￠1￠【￠[￠-1
￠】￠]￠-1
)

; CLASS LV_InCellEdit改动的项目
; 在CLASS_LV_InCellEdit.ahk中搜索“改动”，定位改动项。
; 目前改动了：1编辑单元格时，按del键为删除输入文字。不编辑时，为删除文件。
;             2双击单元格启动编辑，改为单击单元格。


; 常用可改变量↑————————————————————————————————

; 封装文件(脚本图标、元数据程序、pdf转图像程序、裁剪程序)
FileInstall, Image.png, %A_ScriptDir%\Image.png       ; 如果未指定绝对路径, 则假定文件在相对于脚本的目录中
FileInstall, Imagebg.png, %A_ScriptDir%\Imagebg.png
FileInstall, exiftool.exe, %A_ScriptDir%\exiftool.exe
FileInstall, mutool.exe, %A_ScriptDir%\mutool.exe
FileInstall, 图像裁剪工具v2.0.exe, %CropImageToolPath%
FileInstall, custom_icon.ico, %A_ScriptDir%\custom_icon.ico
FileInstall, 命令行调用裁剪工具.exe, %A_ScriptDir%\命令行调用裁剪工具.exe

OnExit("ExitFunc")

GUIdrawLB:
; 获取系统DPI缩放比例
DPIScaling := A_ScreenDPI / 96 * GuiDPI / 96

; 设置图像分辨率
; GUI会跟随系统DPI缩放，但指定了宽高数字的控件不会，要多算一步
Image1Height := 16 * DPIScaling
Image2Width := 230 * DPIScaling
DDLPublisherHeight1 := (13 * DPIScaling) + ((GuiDPI - 96) / 96 / .25)  ; 出版商窗口框高度。和dpi()函数的计算结果可能有小数差别，因此补数调整
DDLPublisherHeight2 := 7 * DPIScaling                                  ; 出版商下拉列表每行高度

dpi(,GuiDPI)  ; 自定义DPI要先调用一次

Gui,main: Default
Gui,main: Color, 202020 
Gui,main: +LastFound +Resize hwndGUIMainHD +OwnDialogs ;AlwaysOnTop


; 创建一些按钮:
Gui, main: Font, % dpi("s9 cFF9900 bold")
Gui, main: Add, Edit, % dpi("x15 y5 H15 W65 HwndYearInputHD Section number")
; 0x80 省略伙伴控件中正常出现在每三位十进制数间的千位分隔符
Gui, main: Add, UpDown, % dpi("Range0-3000 gCopyDateLB 0x80"), 2010    

Gui, main: Font,% dpi("norm"), 微软雅黑
Gui, main: Add, Text, % dpi("x+1 ys"), 年

Gui, main: Font,% dpi("bold"), SimSun
Gui, main: Add, Edit, % dpi("x+6 ys H15 W50 HwndMonthInputHD number")
Gui, main: Add, UpDown, % dpi("Range1-12 Wrap gCopyDateLB")
GuiControl,, %MonthInputHD%, 01

Gui, main: Font,% dpi("norm"), 微软雅黑
Gui, main: Add, Text, % dpi("x+1 ys"), 月
Gui, main: Font,% dpi("s8")
Gui, main: Add, Button, % dpi("x+6 ys H16 W60 gCopyDateLB"), 复制日期

Gui, main: Font, % (FontOptions := "s" DDLPublisherHeight2), % (FontName := "") 
OD_Colors.SetItemHeight(FontOptions, FontName) ; needed by the class to colour the DDL
;if (GuiDPI > 96)
; Gui, main: Font,% dpi("s7"), 微软雅黑
;Else
 Gui, main: Font,% dpi("s9 bold"), 微软雅黑
Gui, main: Add, DropDownList, % dpi("x+16 ys w189 HwndDDLPublisherHD gCopyPublisherLB AltSubmit +0x0210")
, %PublisherList%
; CB_SETITEMHEIGHT = 0x0153
PostMessage, 0x0153, -1, %DDLPublisherHeight1%,, ahk_id %DDLPublisherHD%  ; 设置选区字段的高度.

Gui, main: Font, % dpi("s8 norm")
Gui, main: Add, Button, % dpi("x+10 ys H16 W70 hwndButtonCopyPubHD gCopyPublisherLB"), 复制出版商

Gui, main: Font, % dpi("s11 cFF9900")
Gui, main: Add, Picture, % dpi("x+12 ys H16 w62 hwndSquareHD"),%A_ScriptDir%\Imagebg.png  ;▓▓▓▓▓▓ 
Gui, main: Add, Picture, % dpi("xp+24 ys w-1 h16 hwndImage1HD BackgroundTrans Section"), %A_ScriptDir%\Image.png

; 创建地址栏
Gui,main: Font, % dpi("s10")
DownloadsDir := StrReplace(A_Desktop, "Desktop", "Downloads")
Gui,main:Add, ComboBox,% dpi("x15 w500 HwndEditPathHD gButtonLoadFolderLB AltSubmit Section")
, %A_Desktop%|%DownloadsDir%|选择目录

; 地址栏获取焦点
ControlFocus,, ahk_id %EditPathHD%

Gui,main: Font, % dpi("s9")
Gui,main: Add, Button,% dpi("x+12 H21 w62 HwndButtSaveRenameHD gButtSaveRenameLB Section"), 更改-更名

; 创建封面栏
Gui, main: Font, % dpi("s8")  ; ↓x+11控制图片和listview间隙，下行y+-5控制listview上边y值
Gui, main: Add, Text,% dpi("x+11 ys+10 h25 w230 hwndButtChangeImageHD Right gButtChangeImageLB"), 更 换 封 面
Gui, main: Add, Picture, % dpi("y+-5 hwndImage2HD h-1 w230 gCallCropImageLB Section")
GuiControl, Hide, %ButtChangeImageHD%

Gui,main: Font, % dpi("s11")
ButtLastWidth := ((230 * A_ScreenDPI / 96) / 2 - 15 ) / A_ScreenDPI * 96
Gui, main: Add, Text, % dpi("xs y430 w" ButtLastWidth " Right hwndButtLastImageHD gButtLastImageLB"), <<<
Gui, main: Add, Text, % dpi("x+15 hwndButtNextImageHD gButtNextImageLB"), >>>
GuiControl, Hide, %ButtLastImageHD%
GuiControl, Hide, %ButtNextImageHD%

; 通过 Gui Add 创建 ListView 及其列:
Gui,main: Add, ListView,% dpi("x15 ys w574 r21 vListViewVar hwndListViewHD gListViewLB Background202020 AltSubmit -ReadOnly +LV0x010000 Section")
, 编号|书名|作者|出版日期|丛书|出版商|格式|路径|大小(MB)|页数|制作者|生成工具1(Producer)|生成工具2(Creator Tool)|生成日期|修改日期|元数据日期|关键词|描述              ; exif项目相关命令


; 为了排序, 表示 Size/Page 列中的内容是整数.
LV_ModifyCol(9, "Integer")  
LV_ModifyCol(10, "Integer")  

; 创建图像列表, 这样 ListView 才可以显示图标:
ImageListID1 := IL_Create(2)

; 关联图像列表到 ListView, 然而它就可以显示图标了:
LV_SetImageList(ImageListID1)

; 创建作为上下文菜单的弹出菜单:
Menu, MyContextMenu, Add, 打开, ContextOpenFileLB
Menu, MyContextMenu, Add, 模板重命名文件, RenameFileLB
Menu, MyContextMenu, Add, 删除文件, ContextDelFileLB
Menu, MyContextMenu, Add, 
Menu, MyContextMenu, Add, 删除选中行, ContextClearRowsLB
Menu, MyContextMenu, Add, 编辑选中行, SelectedRowModifyLB
Menu, MyContextMenu, Add, 编辑整列, ModifyColumnLB
Menu, MyContextMenu, Add, 清空整列, EmptyColumnLB
Menu, MyContextMenu, Add, 替换选中行字符, SelectedRowReplaceLB
Menu, MyContextMenu, Add, 预设替换整列, BatchReplaceLB
Menu, MyContextMenu, Add,
Menu, MyContextMenu, Add, 复制元数据, CopyExifLB
Menu, MyContextMenu, Add, 粘贴元数据, PasteExifLB
Menu, MyContextMenu, Add, 撤销粘贴, UndoPasteExifLB
Menu, MyContextMenu, Add, 导入文件名, FileNameAsExifNameLB
Menu, MyContextMenu, Add,
Menu, MyContextMenu, Add, 保存选中行, SaveSelectedRowLB
Menu, MyContextMenu, Add, 保存所有行, SaveEditingLB
Menu, MyContextMenu, Add,
Menu, MyContextMenu, Add, 自动列宽, AutoColSizeLB
Menu, MyContextMenu, Add, 最小列宽, SmallColSizeLB
Menu, MyContextMenu, Add, 窗口置顶, ontopLB
Menu, MyContextMenu, Add,
Menu, MyContextMenu, Add, 全选, SelectAllLB
Menu, Lv2menu1, Add, 说明, InfoLB
Menu, Lv2menu1, Add, 重启, ReloadLB
Menu, Lv2menu1, Add, 查看元数据, ShowExifLB
Menu, Lv2menu1, Add, 元数据热键, ExifHotkeyLB
Menu, Lv2menu1, Add, 生成封面文件, GenerateCoverLB
Menu, Lv2menu1, Add, 设置重命名模板, RenameFileSetupLB
Menu, Lv2menu1, Add, 设置预设替换模板, BatchReplaceSetupLB
Menu, Lv2menu1, Add, 在资源浏览器中打开, ContextOpenInExplorerLB

Menu, Lv3menu1, Add, 100`%, SetGUIdpiLB
Menu, Lv3menu1, Add, 125`%, SetGUIdpiLB
Menu, Lv3menu1, Add, 150`%, SetGUIdpiLB
Menu, Lv3menu1, Add, 175`%, SetGUIdpiLB
Menu, Lv2menu1, Add, 窗口缩放, :Lv3menu1

Menu, MyContextMenu, Add, 其他,:Lv2menu1


; 利用空行微调窗口总高度
Gui,main:Font, % dpi("s1")
Gui,main:Add,Text,% dpi("h17 HwndTextCtrlYHD")  ; 通过加减h，微调窗口总高度
GuiControl, Hide, %TextCtrlYHD%

; 显示窗口
Gui,main:Show, % dpi("w845"), %appname1%

; 修改控件颜色，RGB模式
Gui,main: Color,, 202020
OD_Colors.Attach(DDLPublisherHD,{T: 0x64ffda, B: 0x202020})
CtlColors.Attach(EditPathHD, "202020", "FF9900")

; 经测试，edit加了g标签后，绘制gui时会自动跳转几十至几百次,
; 因此显示后再加g标签，避免额外的复制提示。
GuiControl, +gCopyDateLB, %YearInputHD%
GuiControl, +gCopyDateLB, %MonthInputHD%

; 确定窗口高度后，再显示底行（状态栏）
Gui,main:Font, % dpi("s11")
Gui,main:Add,Text,% dpi("w180 yp-3 HwndGUIattribute1 ")
;显示选择多少项及显示搜索中
Gui,main:Add,Text,% dpi("w180 x+m HwndGUIattribute2")
Gui,main:Add,Text,% dpi("w180 x+m HwndGUIattribute3")
; 启用单元格编辑
LV_InCellEdit.OnMessage()
If !(LV_InCellEdit.Attach(ListViewHD, True, True))
 MsgBox, % "注册ListView句柄失败: " . ErrorLevel

; 设置哪些列可以编辑
LV_InCellEdit.SetColumns(ListViewHD, [ 2, 3, 4, 5, 6, 11, 12, 13, 14, 15, 16, 17, 18])      ; exif项目相关命令

; 注册热键
; 在Class_LV_InCellEdit.ahk里已注册
;Hotkey, Del, ContextDelFileLB, On


gosub, BatchReplaceReadLB
return


; 按钮_复制日期
CopyDateLB:
GuiControlGet, YearInput,, %YearInputHD%
GuiControlGet, MonthInput,, %MonthInputHD%

; 如果数字是个位数，在前面加上 0
if (MonthInput < 10)
{
 MonthInput := Format("{:02}", MonthInput)
 GuiControl,, %MonthInputHD%, %MonthInput%
}
 
CopyDateTime := A_TickCount

if !CopyDateTipIDN
{
 SetTimer, CopyDateTipLB, -1000
 CopyDateTipIDN = 1
}
return

; 1秒后检查变化
CopyDateTipLB:
if ((A_TickCount - CopyDateTime) > 990)
{
 Clipboard := YearInput MonthInput
 ToolTip, 已复制
 SetTimer, ToolTipoffLB, -700
 CopyDateTipIDN =
} else 
 SetTimer, CopyDateTipLB, -1000
return


; 按钮_复制出版商
CopyPublisherLB:
GuiControlGet, DDLPublisher,,%DDLPublisherHD%, Text
DDLPublisher := SubStr(DDLPublisher,3)
Clipboard := DDLPublisher
ToolTip, 已复制
SetTimer, ToolTipoffLB, -700
return


; 按钮_前往目录
ButtonLoadFolderLB:
Thread, Priority, 1  ; 读取过程不被按钮打断
Gui,main:Default

; 1 如果有编辑尚未保存，显示提示
if UnSaveIDN
{
 MsgBox, 4100,, 有编辑尚未保存，确定打开路径？
 IfMsgBox, No
 {
  ; 恢复输入栏原路径
  gosub, RefreshEditpathLB
  return
 }
 else
 {
  LV_InCellEdit.Changed.Remove(ListViewHD, "")
  UnSaveIDN := ""
 }
}

; 2 正常打开路径（保存编辑后刷新列表则跳过此段）
if !RefreshIDN
{
 ; 获取输入栏路径
 GuiControlGet, EditPath,,%EditPathHD%, Text
 
 ; 如果是正常路径（非"选择目录"）
 if !(EditPath = "选择目录") and EditPath
 {
  ;检测输入路径是否为文件夹
  if !InStr(FileExist(EditPath),"D")    
  {
   ToolTip, 输入路径不是文件夹,请检查
   SetTimer, ToolTipoffLB, -2000
   return
  }
  else
   LoadFolder := EditPath
 }
 ; 如果是"选择目录"
 else
 {
  Gui +OwnDialogs    ; 强制用户解除此对话框后才可以操作主窗口.
  FileSelectFolder, EditPath,, 3, Select a folder to read:
  if !EditPath {     ; 用户取消了对话框.
   ; 恢复输入栏原路径
   gosub, RefreshEditpathLB
   return
  } else
   LoadFolder := EditPath
 }
}

; 3 更新输入的路径到地址栏
; 1) 检查文件夹名称的最后一个字符是否为反斜杠, 对于根目录则会如此,
; 例如 C:\. 如果是, 则移除这个反斜杠以避免之后出现两个反斜杠.
LastChar := SubStr(LoadFolder, 0)
if (LastChar = "\")
 LoadFolder := SubStr(LoadFolder, 1, -1)  ; 移除尾随的反斜杠.
; 2) 更新
gosub, RefreshEditpathLB

; 4 任务栏显示处理提示
GuiControl,,%GUIattribute1%, 已读取 共0项

; 5 清理环境
; 1) 清空 ListView, 但为了简化保留了图标缓存.
LV_Delete()
GuiControl,,%GUIattribute3%, |

; 2) 关闭裁剪程序
SetTimer, CheckCropImageFinishLB, Off
WinClose, ahk_exe %CropImageToolPath%
GuiControl,, %Image2HD%
GuiControl, Hide, %ButtChangeImageHD%
GuiControl, Hide, %ButtLastImageHD%
GuiControl, Hide, %ButtNextImageHD%

; 6 检查文件数量
; 注意，FileNo要先赋值0。
; 当下行loop files时，如果没有loop到文件，将不会更改FileNo原来的值，导致可能用到旧值，因此先赋值0。
; 整个线程遍历文件三次，第一次检查PDF文件数量，第二次检查文件夹数量，第三次配置图标及PDF元数据
; 1）先检查pdf文件数量
FileNo = 0
Loop, Files, %LoadFolder%\*.%TypeFileExt%, F
{
 ; 检查文件数量
 FileNo := A_Index
}

; 2）如果没有PDF文件，则置空后缀名变量，准备显示文件夹
if !FileNo
{
 Loop, Files, %LoadFolder%\*.*, D     ; 检查文件夹数量
  FileNo := A_Index
 LoopFileExt := "*"
 FileExt =                            ; 文件后缀置空
 LoopFileMode := "D"                  ; 设置遍历模式
}
; 2) 如果有PDF文件，则赋值后缀名变量
else
{
 LoopFileExt := FileExt := TypeFileExt
 LoopFileMode := "F"                  ; 设置遍历模式
}
 
; 3) 如果既没有PDF文件，也没有文件夹，则显示0
if !FileNo
{
 FileNo = 0
 gosub, PostSearchLB
 return
}

; 7 设置进度条分段
GuiControl,,%GUIattribute1%, 需读取%FileNo%项
SetTaskbarProgress(0, "N", GUIMainHD)
ProgressNo := 100 / FileNo

; 8 计算 SHFILEINFO 结构需要的缓存大小.
sfi_size := A_PtrSize + 8 + (A_IsUnicode ? 680 : 340)
VarSetCapacity(sfi, sfi_size)

; 9 处理一些变量内容
ClipSaved := ClipboardAll   ; 把剪贴板的所有内容保存到您选择的变量中.
ErrorLog := ""              ; 清空错误记录变量

; 10 获取所选择文件夹中的文件名列表并添加到 ListView:
;GuiControl, -Redraw, ListViewVar  ; 在加载时禁用重绘来提升性能.
Loop, Files, %LoadFolder%\*.%LoopFileExt%, %LoopFileMode%
{
 ; 任务栏标签显示进度
 SetTaskbarProgress(A_Index * ProgressNo)
 
 ; 1）每次循环开始时的一些准备
 ; 如果只是刷新列表，则仅读取刚处理的文件
 if RefreshIDN
 {
  if A_LoopFileLongPath not in %FilePathList%     ; 对于 "in" 运算符, 需要准确匹配列表中的某项.
   continue
 }
 
 GuiControl,,%GUIattribute2%, |  正在读取第%A_Index%项
 
 ; 先清空一些列项目之前的值
 BookTitle := BookAuthor := BookPubDate := BookSeries := BookPublisher := BookPageCount 
 := BookCreator := BookProducer := BookCreatorTool := BookCreateDate := BookModifyDate 
 := BookMetadataDate := BookKeyWords := BookDescription := ""     ; exif项目相关命令

 ; 2）下面提取文件图标
 ExtID := 0  ; 进行初始化来处理为更短的扩展名.
 Loop 7      ; 限制扩展名为 7 个字符, 这样之后计算的结果才能存放到 64 位值.
 {
  ExtChar := SubStr(FileExt, A_Index, 1)
  if not ExtChar  ; 没有更多字符了.
   break
  ; 通过给每个字符分配一个不同的比特位置, 来得到唯一 ID:
  ExtID := ExtID | (Asc(ExtChar) << (8 * (A_Index - 1)))
 }
 ; 检查此文件扩展名的图标是否已经在图像列表中. 如果是,
 IconNumber := IconArray%ExtID%
 
 if !IconNumber  ; 扩展名还没有相应的图标, 所以进行加载.
 {
  ; 获取与此文件扩展名关联的高质量小图标:
  DllCall("Shell32\SHGetFileInfo" . (A_IsUnicode ? "W":"A"), "Str", A_LoopFileLongPath
  , "UInt", 0, "Ptr", &sfi, "UInt", sfi_size, "UInt", 0x101)  ; 0x100也可 0x101 为 SHGFI_ICON+SHGFI_SMALLICON
  
  ; 从结构中提取 hIcon 成员:
  hIcon := NumGet(sfi, 0)
  
  ; 下面加上 1 来把返回的索引从基于零转换到基于一:
  IconNumber := DllCall("ImageList_ReplaceIcon", "Ptr", ImageListID1, "Int", -1, "Ptr", hIcon) + 1
  
  ; 现在已经把它复制到图像列表, 所以应销毁原来的:
  DllCall("DestroyIcon", "Ptr", hIcon)
  
  ; 缓存图标来节省内存并提升加载性能:
  IconArray%ExtID% := IconNumber
 }

 ; 3）这段仅应用于pdf文件

 if (LoopFileExt = TypeFileExt)
 {
  FileNameNoExt := StrReplace(A_LoopFileName, "." TypeFileExt)
  
  ; 如果没有同名封面，则提取第一页为封面
  if !FileExist(LoadFolder "\" FileNameNoExt ".jpg")
  {
   RunWait, %ComSpec% /c %MutoolPath% convert -o "%LoadFolder%\%FileNameNoExt%.jpg" -O width=800 "%A_LoopFileLongPath%" 1 ,, Hide
   FileMove, %LoadFolder%\%FileNameNoExt%1.jpg, %LoadFolder%\%FileNameNoExt%.jpg
  }
  
  ; 获取书名
  SubCode = 
  (join`s Comment
  -Title -Author -Date -Subject -Publisher -PageCount -Creator -Producer -CreatorTool    ; exif项目相关命令
  -CreateDate -ModifyDate -XMP:MetadataDate -KeyWords -Description "%A_LoopFileLongPath%"
  )
  
  ; 运行exiftool
  gosub, RunCMDLB
  
  ; 如果有错误，则记录
  Loop, parse, OutputVar, `n, `r
  {
   if A_LoopField contains error
    ErrorLog .= OutputVar "`n"
  }  
  
  Loop, Parse, OutputVar, `n, `r   ; exif项目相关命令
  {
   if InStr(A_LoopField, "Title                           : ")
    BookTitle := StrReplace(A_LoopField, "Title                           : ")
   if InStr(A_LoopField, "Author                          : ")
    BookAuthor := StrReplace(A_LoopField, "Author                          : ")
   if InStr(A_LoopField, "Date                            : ")
    BookPubDate := StrReplace(A_LoopField, "Date                            : ")
   if InStr(A_LoopField, "Subject                         : ")
    BookSeries := StrReplace(A_LoopField, "Subject                         : ")
   if InStr(A_LoopField, "Publisher                       : ")
    BookPublisher := StrReplace(A_LoopField, "Publisher                       : ")
   if InStr(A_LoopField, "Page Count                      : ")
    BookPageCount := StrReplace(A_LoopField, "Page Count                      : ")
   if InStr(A_LoopField, "Creator                         : ")
    BookCreator := StrReplace(A_LoopField, "Creator                         : ")   
   if InStr(A_LoopField, "Producer                        : ")
    BookProducer := StrReplace(A_LoopField, "Producer                        : ")   
   if InStr(A_LoopField, "Creator Tool                    : ")
    BookCreatorTool := StrReplace(A_LoopField, "Creator Tool                    : ")
   if InStr(A_LoopField, "Create Date                     : ")
    BookCreateDate := StrReplace(A_LoopField, "Create Date                     : ")   
   if InStr(A_LoopField, "Modify Date                     : ")
    BookModifyDate := StrReplace(A_LoopField, "Modify Date                     : ")   
   if InStr(A_LoopField, "Metadata Date                   : ")
    BookMetadataDate := StrReplace(A_LoopField, "Metadata Date                   : ")
   if InStr(A_LoopField, "Keywords                        : ")
    BookKeyWords := StrReplace(A_LoopField, "Keywords                        : ")
   if InStr(A_LoopField, "Description                     : ")
    BookDescription := StrReplace(A_LoopField, "Description                     : ")   
  }
 }

 ; 4）收尾处理  
 ; 排序编号设为三位数
 OrderNo := Format("{:03}", A_Index)

 ; 在 ListView 中创建新行并把它和上面的图标编号进行关联:
 LV_Add("Icon" IconNumber, OrderNo, BookTitle, BookAuthor, BookPubDate, BookSeries
 , BookPublisher, FileExt, A_LoopFileLongPath, A_LoopFileSizeMB, BookPageCount
 , BookCreator, BookProducer, BookCreatorTool , BookCreateDate, BookModifyDate, BookMetadataDate
 , BookKeyWords, BookDescription)      ; exif项目相关命令
}

; 任务栏标签取消显示进度
SetTaskbarProgress(0)

;GuiControl, +Redraw, ListViewVar  ; 重新启用重绘(上面把它禁用了).

; 11 还原剪贴板及清空缓存的元数据和行索引
Clipboard := ClipSaved      ; 恢复剪贴板为原来的内容. 注意这里使用 Clipboard(不是 ClipboardAll).
ClipSaved := BookExifSave := RowIndex1 := ""  ; 清空“粘贴元数据历史记录”（不管有无）
 
; 12 优化状态显示
; 把列加宽一些以便显示出它的标题.
gosub, AutoColSizeLB

; 搜索后处理
gosub, PostSearchLB

; 13 保存错误记录
if ErrorLog
{
 if FileExist(A_ScriptDir "\ErrorLog.txt")
  FileRecycle, %A_ScriptDir%\ErrorLog.txt
 FormatTime, DateString,, yyyy-MM-dd HH:mm:ss
 FileAppend, %DateString%`n`n%ErrorLog%, %A_ScriptDir%\ErrorLog.txt, UTF-8
 Run, %A_ScriptDir%\ErrorLog.txt
 ToolTip, 请查看错误记录
 SetTimer, ToolTipoffLB, -1500
}
return


; 搜索后处理
PostSearchLB:
; 状态栏消除「正在读取」
GuiControl,,%GUIattribute2%, |

FileNo := LV_GetCount()
GuiControl,,%GUIattribute1%, 已读取 共%FileNo%项

return


; 恢复输入栏原路径
RefreshEditpathLB:
GuiControl,, %EditPathHD%, |  
GuiControl,, %EditPathHD%, %LoadFolder%|%A_Desktop%|%DownloadsDir%|选择目录

if LoadFolder
 GuiControl, Choose, %EditPathHD%, 1
Return


; 按钮_保存更改
; 右键菜单_保存所有行
SaveEditingLB:
Critical
Gui,main:Default

; 下面这段为检查路径是否合法
; 1)路径是否存在
if !LoadFolder
{
 ToolTip, 输入路径为空
 SetTimer, ToolTipOffLB, -1500
 return
}

; 2.1)路径是否为文件（PDF）还是文件夹
; 2.2）路径是否存在
FileNo = 0
Loop, % LV_GetCount()
{
 LV_GetText(CellText, A_Index, 8)
 if FileExist(CellText)
 {
  if !InStr(FileExist(CellText), "D")
   ++ FileNo
 }
 else
  MsgBox, 4096,, 该文件不存在，请删除此行或放入文件`n%CellText%
}

; 没有类型文件则返回
if !FileNo
{
 ToolTip, 未找到%TypeFileExt%文件
 SetTimer, ToolTipoffLB, -1500
 return
}

; 关闭裁剪程序
SetTimer, CheckCropImageFinishLB, Off
WinClose, ahk_exe %CropImageToolPath%

; 任务栏显示处理提示
GuiControl,,%GUIattribute3%, |

; 设置进度条分段
SetTaskbarProgress(0, "N", GUIMainHD)
ProgressNo := 100 / FileNo

; 处理一些变量内容
ErrorLog := ""              ; 清空错误记录变量

Loop, % LV_GetCount()
{
 ; 任务栏标签显示进度
 SetTaskbarProgress(A_Index * ProgressNo)
 
 GuiControl,,%GUIattribute2%, |  正在保存第%A_Index%项
 
 ; 获取文件路径
 LV_GetText(FilePath, A_Index, 8)
 
 ; 建立保存文件名单，当RefreshIDN=1时用
 FilePathList .= FilePath ","

 ; 检查要写入的pdf是否已打开
 WinGet, ActiveWindowList, List                       ; 获取所有窗口的句柄
 Loop, %ActiveWindowList%                             ; 遍历所有窗口句柄一次
 {
  ActiveWindow := ActiveWindowList%A_Index%           ; 逐个赋值句柄
  WinGetTitle, WindowTitle, % "ahk_id " ActiveWindow  ; 获取句柄的窗口标题
  
  SplitPath, FilePath, OpenFileName                   ; 获取当前pdf文件的带后缀标题
  if (InStr(windowTitle, OpenFileName))               ; 如果窗口标题含有上述标题，则提示
   MsgBox, 4096, 正在写入元数据, 请关闭文件：`n%windowTitle%
 }

 ; 先清空一些列项目之前的值
 BookTitle := BookAuthor := BookPubDate := BookSeries := BookPublisher := BookPageCount 
 := BookCreator := BookProducer := BookCreatorTool := BookCreateDate := BookModifyDate 
 := BookMetadataDate := BookKeyWords := BookDescription := ""     ; exif项目相关命令
 
 ; 读取列表信息
 RowNumber := A_Index                   ; 当前行号
 Loop, 18                               ; 根据有多少列，循环多少次   ; exif项目相关命令
 {
  LV_GetText(CellText, RowNumber, A_Index)
  if (A_Index = 2)
   BookTitle := CellText
  if (A_Index = 3)
   BookAuthor := CellText
  if (A_Index = 4)
   BookPubDate := ConvertToYearMonth(CellText)
  if (A_Index = 5)
   BookSeries := CellText
  if (A_Index = 6)
   BookPublisher := CellText
  if (A_Index = 8)
   FilePath := CellText                ; 获取路径 
  if (A_Index = 11)
   BookCreator := CellText
  if (A_Index = 12)
   BookProducer := CellText
  if (A_Index = 13)
   BookCreatorTool := CellText  
  if (A_Index = 14)
   BookCreateDate := CellText 
  if (A_Index = 15)
   BookModifyDate := CellText
  if (A_Index = 16)
   BookMetadataDate := CellText
  if (A_Index = 17)
   BookKeyWords := CellText 
  if (A_Index = 18)
   BookDescription := CellText  
 }
 Sleep, 200

 ; 导入到目标文件EXIF
 ; 脚本借用主题（Subject）存入书籍系列
 
 ; 设置MetadataDate为当前时间的命令例子：
 ; exiftool -XMP:MetadataDate=now example.pdf

 subcode=     ; exif项目相关命令
 (Join`s Comment
 "-Title=%BookTitle%"                      ; 书名
 "-Author=%BookAuthor%"                    ; 作者
 "-Date=%BookPubDate%"                     ; 出版日期
 "-Subject=%BookSeries%"                   ; -书籍系列←借用为（主题）
 "-Publisher=%BookPublisher%"              ; 出版商
 "-Creator=%BookCreator%"                  ; 制作者
 "-Producer=%BookProducer%"                ; 生成工具1
 "-CreatorTool=%BookCreatorTool%"          ; 生成工具2
 "-CreateDate=%BookCreateDate%"            ; 生成日期
 "-ModifyDate=%BookModifyDate%"            ; 修改日期
 "-XMP:MetadataDate=%BookMetadataDate%"    ; 元数据日期记录的是元数据本身的最后修改时间。
 "-Keywords=%BookKeywords%"                ; 关键词
 "-Description=%BookDescription%"          ; 描述 
 "%FilePath%"
 )

 ; 运行ExifTool
 gosub, RunCMDLB

 ; 如果有错误，则显示提示
 Loop, parse, OutputVar, `n, `r
 {
  if A_LoopField contains error
   ErrorLog .= OutputVar "`n"
 }
}

; 任务栏标签取消显示进度
SetTaskbarProgress(0)

; 恢复任务栏显示
GuiControl,,%GUIattribute2%, |

; 清空编辑历史记录
UnSaveIDN := ""
LV_InCellEdit.Changed.Remove(ListViewHD, "")

; 清空缓存的元数据和行索引
BookExifSave := RowIndex1 := ""  ; 清空“粘贴元数据历史记录”（不管有无）

; 保存错误记录
if ErrorLog
{
 if FileExist(A_ScriptDir "\ErrorLog.txt")
  FileRecycle, %A_ScriptDir%\ErrorLog.txt
 FormatTime, DateString,, yyyy-MM-dd HH:mm:ss
 FileAppend, %DateString%`n`n%ErrorLog%, %A_ScriptDir%\ErrorLog.txt, UTF-8
 Run, %A_ScriptDir%\ErrorLog.txt
 ToolTip, 请查看错误记录
 SetTimer, ToolTipoffLB, -1500
}

; RefreshIDN = 1表示由保存线程引起的列表刷新，仅刷新保存的文件
RefreshIDN = 1
if !ErrorLog                ; 如果没有错误，则刷新列表
 gosub, ButtonLoadFolderLB

; 清空一些变量值
RefreshIDN := FilePathList := ""
Return


; 按钮_更改-更名
ButtSaveRenameLB:
gosub, SaveEditingLB
; 如果没有错误，则执行更改文件名
if !ErrorLog
{
 SetTimer, SelectAllLB, -300
 TimerPeriod := (FileNo * 40) +500
 SetTimer, RenameFileLB, % "-" TimerPeriod
}
Return

; 按钮_上一个书籍的封面
ButtLastImageLB:
Gui, main : Default
SetTimer, CheckCropImageFinishLB, Off

; 查找选中行.
SelectedRowNumber := LV_GetNext()  
LV_Modify(SelectedRowNumber, "-Select")
LV_Modify(SelectedRowNumber, "-Focus")
--SelectedRowNumber
if (SelectedRowNumber < 1)
 SelectedRowNumber := 1
LV_Modify(SelectedRowNumber, "Select")
LV_Modify(SelectedRowNumber, "Focus")

; 显示封面
LV_GetText(CoverPath, SelectedRowNumber, 8)  ; 获取第8个字段的文本.
CoverPath := StrReplace(CoverPath, ".pdf", ".jpg")
SplitPath,  CoverPath,,,, CoverNameNoExt
GuiControl,, %Image2HD%, *w%Image2Width% *h-1 %CoverPath%

; 调用裁剪程序
gosub, CallCropImageLB
return


; 按钮_下一个书籍的封面
ButtNextImageLB:
Gui, main : Default
SetTimer, CheckCropImageFinishLB, Off

; 查找选中行.
SelectedRowNumber := LV_GetNext()
FileNo := LV_GetCount()
LV_Modify(SelectedRowNumber, "-Select")
LV_Modify(SelectedRowNumber, "-Focus")
++SelectedRowNumber
if (SelectedRowNumber > FileNo)
 SelectedRowNumber := FileNo
LV_Modify(SelectedRowNumber, "Select")
LV_Modify(SelectedRowNumber, "Focus")

; 显示封面
LV_GetText(CoverPath, SelectedRowNumber, 8)  ; 获取第8个字段的文本.
CoverPath := StrReplace(CoverPath, ".pdf", ".jpg")
SplitPath,  CoverPath,,,, CoverNameNoExt
GuiControl,, %Image2HD%, *w%Image2Width% *h-1 %CoverPath%

; 调用裁剪程序
gosub, CallCropImageLB
return


; 打开图像裁剪程序
CheckCropWinLB:
Thread, Priority, 1  ; 打开过程不被上一个或下一个封面按钮打断
; 打开程序
if !FileExist(CropImageToolPath)
{
 MsgBox, 4096,, %CropImageTool%不存在，请检查
 Exit
} Else

run, % CropImageToolPath

Loop
{
 if !WinActive("ahk_exe" CropImageToolPath)
  WinActivate, ahk_exe %CropImageToolPath%
 Else
  break
 Sleep, 200
}

WinSet, AlwaysOnTop, On, ahk_exe %CropImageToolPath%
return


; 调用裁剪程序
CallCropImageLB:
SetTimer, CheckCropImageFinishLB, Off                 ; 关闭裁剪进度检查

if !WinExist("ahk_exe" CropImageToolPath)
 gosub, CheckCropWinLB                                ; 如果窗口不存在，则打开程序
FileGetTime, FileModifiedTime1, %CoverPath%           ; 获取文件时间
run, % CropImageCallPath " """ CoverPath """",, Hide  ; 打开图像

; 当有裁剪则刷新封面显示
SetTimer, CheckCropImageFinishLB, 500                 ; 定时裁剪进度检查
Return


; 裁剪进度检查
CheckCropImageFinishLB:
FileGetTime, FileModifiedTime2, %CoverPath%

; 文件时间不同，更新显示
if (FileModifiedTime1 != FileModifiedTime2)
{
 FileModifiedTime1 := FileModifiedTime2
 GuiControl,, %Image2HD%, *w%Image2Width% *h-1 %CoverPath%
}

; 如果关闭裁剪程序，则关闭裁剪进度检查
if !WinExist("ahk_exe" CropImageToolPath)
 SetTimer, CheckCropImageFinishLB, Off
Return


; 按钮_更换封面
ButtChangeImageLB:
; 打开一个文件选择对话框来选择图像文件
FileSelectFile, ImageFile, 3, %CoverPath%, 选择一个图像文件, 图片文件 (*.png; *.jpg; *.bmp)
if (ImageFile = "")
 return

FileCopy, %ImageFile%, %CoverPath%, 1
GuiControl,, %Image2HD%, *w%Image2Width% *h-1 %CoverPath%
return


; 清空ListView
ButtonClearLB:
LV_Delete()  ; 清理 ListView, 但为了简化保留了图标缓存.
return


; 左键动作
ListViewLB:
Gui, main: Default
; 双击左键
if (A_GuiEvent = "DoubleClick")                  ; 脚本还可以检查许多其他的可能值.
{
 LV_GetText(FilePath, A_EventInfo, 8)            ; 获取第8个字段的文本.
 if InStr(FileExist(FilePath),"D")               ; 如果选中路径路径是文件夹
  gosub, RedirectDirLB                           ; 应用内打开路径
 else {
  Run "%FilePath%",, UseErrorLevel
 if ErrorLevel
  MsgBox,4096,提示,无法打开“%FilePath%”
 }
 return
}

; 如果显示的是文件夹，则返回
if (LoopFileMode = "D") or !LoadFolder
 Return

; 检查是否有编辑未保存
if !UnSaveIDN      ; 未保存辨别符，「空值」时检查是否有编辑未保存
{
 If LV_InCellEdit.Changed.HasKey(ListViewHD)
 {
  GuiControl,,%GUIattribute2%, |  未保存          ; 状态栏显示未保存
  ++UnSaveIDN
 }
}

; 仅鼠标点击或键盘输入时执行
if (A_GuiEvent = "Normal") or (A_GuiEvent = "RightClick") or (A_GuiEvent = "DoubleClick") or (A_GuiEvent = "K") 
{
SelectedRowNumber := LV_GetNext()                    ; 查找选中行.
if !SelectedRowNumber                                ; 没有选中行或读取中
{
 GuiControl,, %Image2HD%           ; 不显示封面图像
 GuiControl, Hide, %ButtChangeImageHD%
 GuiControl, Hide, %ButtLastImageHD%
 GuiControl, Hide, %ButtNextImageHD%
 GuiControl,,%GUIattribute3%, |
 return                            ; 返回
}

if A_EventInfo in 16,17,18  ; 如果按键是Shift，Control或Alt
 Return

; 显示封面图及路径链接
LV_GetText(CoverPath, SelectedRowNumber, 8)      ; 获取第8个字段的文本.
CoverPath := StrReplace(CoverPath, ".pdf", ".jpg")
SplitPath, CoverPath,,,, CoverNameNoExt
GuiControl,, %Image2HD%, *w%Image2Width% *h-1 %CoverPath%
GuiControl, Show, %ButtChangeImageHD%
GuiControl, Show, %ButtLastImageHD%
GuiControl, Show, %ButtNextImageHD%

LV_GetText(FileSize, SelectedRowNumber, 9)            ; 获取第9个字段的文本.
LV_GetText(BookPageCount, SelectedRowNumber, 10)      ; 获取第10个字段的文本.
GuiControl,, %GUIattribute3%, |  %FileSize%MB    共%BookPageCount%页

if WinExist("ahk_exe" CropImageToolPath)
 gosub, CallCropImageLB
}
return


; 应用内转到文件夹
RedirectDirLB:
Gui,main:Default
; 仅对选中行进行操作而不是所有选择的行:
SelectedRowNumber := LV_GetNext()  ; 查找选中行.
if !SelectedRowNumber              ; 没有选中行.
 return

LV_GetText(FilePath, SelectedRowNumber, 8)  ; 获取第8个字段的文本.
if !InStr(FileExist(FilePath),"D")          ; 如果选中路径路径不是文件夹，则返回
 Return

GuiControl,, %EditPathHD%, %FilePath%
GuiControl, Choose, %EditPathHD%, 5
gosub, ButtonLoadFolderLB
return


; 单击右键显示右键菜单
mainGuiContextMenu:                        ; 运行此标签来响应右键点击或按下 Apps 键.
if (A_GuiControl != "ListViewVar")         ; 仅在 ListView 中点击时才显示菜单.
 return

; 获取列号
ColumnNo := LV_SubItemHitTest(ListViewHD)

; 在提供的坐标处显示菜单, A_GuiX 和 A_GuiY. 应该使用这些
; 因为即使用户按下 Apps 键, 它们也会提供正确的坐标:
Menu, MyContextMenu, Show, % dpi(A_GuiX "," A_GuiY)
return


; 右键菜单_打开
ContextOpenFileLB:  ; 用户在上下文菜单中选择了 "打开".
Gui,main:Default
; 为了简化, 仅对焦点行进行操作而不是所有选择的行:
SelectedRowNumber := LV_GetNext()  ; 查找选中行.
if !SelectedRowNumber              ; 没有选中行.
{
 gosub, ButtonLoadFolderLB
 return
}

LV_GetText(FilePath, SelectedRowNumber, 8)      ; 获取第8个字段的文本.
if InStr(FileExist(FilePath),"D")               ; 如果选中路径路径是文件夹
 gosub, RedirectDirLB                           ; 应用内打开路径
else {
 Run %FilePath%,, UseErrorLevel
 if ErrorLevel
  MsgBox, 4096,, 未能打开"%FilePath%"
}
return


;右键菜单_在资源浏览器中打开
ContextOpenInExplorerLB:
Gui, main : Default
SelectedRowNumber := LV_GetNext()  ; 查找选中行.
if !SelectedRowNumber              ; 没有选中行.
 return
LV_GetText(FilePath, SelectedRowNumber, 8) ; 获取第8个字段的文本.

Run,% "explorer.exe /select," """" FilePath """"
return


; 右键菜单_重命名文件
RenameFileLB:
; 路径是否存在
if !LoadFolder
{
 ToolTip, `        输入路径为空`        `
 SetTimer, ToolTipOffLB, -1500
 return
}

if !ReNamePattern
{
 ToolTip, `        重命名模式为空，请先设置`        `
 SetTimer, ToolTipOffLB, -1500
 return
}

Gui, main : Default

SelectedRowNumber := LV_GetNext()  ; 查找选中行.
if !SelectedRowNumber              ; 没有选中行.
{
 ToolTip, `        没有选中行`        `
 SetTimer, ToolTipOffLB, -1500
 return
}

; 重命名选中书籍及对应的封面文件
SelectedRowNumber := 0  ; 这会使得首次循环从顶部开始搜索.

; 设置字段替换顺序，避免重名字段替换出错。
; 例如：先替换"BookCreator"，会将"BookCreatorTool"也替换掉。exif项目相关命令
ArrayKeyOrder := ["BookTitle", "BookAuthor", "BookPubDate", "BookSeries"
, "BookPublisher", "BookPageCount", "BookProducer", "BookCreatorTool"
, "BookCreateDate", "BookModifyDate", "BookMetadataDate", "BookKeyWords"
, "BookDescription", "BookCreator"]

; 设置年月日位数（默认为年份4位）
DateDigit := Radio1Sub3Var ? 6 : (Radio2Sub3Var ? 8 : 4)

; 暂存重命名模板
ReNamePatternTemp := ReNamePattern

; 设置进度条分段
FileNo := LV_GetCount()
SetTaskbarProgress(0, "N", GUIMainHD)
ProgressNo := 100 / FileNo

Loop  ; exif项目相关命令
{
 SelectedRowNumber := LV_GetNext(SelectedRowNumber)
 if !SelectedRowNumber  ; 上面返回零, 所以没有更多选择的行了.
  break
 
 ; 任务栏标签显示进度
 if !Sub3GuiTestIDN  ; 如果不是测试，则显示进度
  SetTaskbarProgress(A_Index * ProgressNo)
 
 ; 恢复重命名模板
 ReNamePattern := ReNamePatternTemp

 ; 创建一个对象来存储键值对
 BookInfoArray := {}
 Loop, 18
 {
  LV_GetText(CellText, SelectedRowNumber, A_Index)
  switch A_Index
  {
   case 2:
    BookInfoArray["BookTitle"] := CellText
   case 3:
    BookInfoArray["BookAuthor"] := CellText
   case 4:                                    ; 出版日期做补0处理
    BookInfoArray["BookPubDate"] := SubStr(ConvertToYearMonth(CellText), 1, DateDigit)
   case 5:
    BookInfoArray["BookSeries"] := CellText
   case 6:
    BookInfoArray["BookPublisher"] := CellText
   case 8:
    FilePath := CellText
   case 10:
    BookInfoArray["BookPageCount"] := CellText
   case 11:
    BookInfoArray["BookCreator"] := CellText
   case 12:
    BookInfoArray["BookProducer"] := CellText
   case 13:
    BookInfoArray["BookCreatorTool"] := CellText
   case 14:
    BookInfoArray["BookCreateDate"] := CellText
   case 15:
    BookInfoArray["BookModifyDate"] := CellText
   case 16:
    BookInfoArray["BookMetadataDate"] := CellText
   case 17:
    BookInfoArray["BookKeyWords"] := CellText
   case 18:
    BookInfoArray["BookDescription"] := CellText
  }
 }
 
 ; 将命名模板转为对应的元数据
 for index, value in ArrayKeyOrder
 {
  if InStr(ReNamePattern, value)    ; 在模板中查找字段
   ; 如找到，则替换字段为对应的元数据，没有则删除字段
   ReNamePattern := StrReplace(ReNamePattern, value, BookInfoArray[value])
 }

 ; 多个连续的连接符替换成一个，删除最后的连接符
 ReNamePattern := RTrim(RegExReplace(ReNamePattern, Sub3EditLinkVar "{2,}", Sub3EditLinkVar), Sub3EditLinkVar)
 
 ; 删除文件名不规范字符
 ReNamePattern := RegExReplace(ReNamePattern, "[\\/:*?""<>|]")

 ; 将空格替换为指定变量
 ;ReNamePattern := StrReplace(ReNamePattern, A_Space, Sub3EditLinkVar)
 
 ; 如测试则中止
 if Sub3GuiTestIDN
 {
  MsgBox, 4096,, % "文件重命名预览：`n`n" ReNamePattern ".pdf"
  ; 恢复重命名模板
  ReNamePattern := ReNamePatternTemp
  break
 }
 
 ; 重命名
 FileMove, %FilePath%, %LoadFolder%\%ReNamePattern%.pdf, 1       ; 1为覆盖源文件
 FilePath := SubStr(FilePath, 1,-4)                              ; 路径删掉.pdf
 FileMove, %FilePath%.jpg, %LoadFolder%\%ReNamePattern%.jpg, 1
}

; 恢复重命名模板
ReNamePattern := ReNamePatternTemp

; 如果不是测试，则
if !Sub3GuiTestIDN
{
 SetTaskbarProgress(0)  ; 取消显示进度
 ToolTip, `        已重命名文件`        ` ; 显示提示
 SetTimer, ToolTipOffLB, -1500
}

gosub, ButtonLoadFolderLB
return


; 右键菜单_设置重命名文件
RenameFileSetupLB:
Gui, Sub3: Color, 202020 
Gui, Sub3: +AlwaysOnTop
Gui, Sub3: Font, % dpi("s11 cFF9900"),微软雅黑 
Gui, Sub3: add, Text, % dpi("y15 HwndSub3Text1HD"), 选择需要命名的元数据字段
GuiControlGet, Sub3Text1Pos, Pos, %Sub3Text1HD%
Sub3Text1PosX := ((952 * GuiDPI / 96) - Sub3Text1PosW) / 2 
GuiControl, move, %Sub3Text1HD%, x%Sub3Text1PosX%
Gui, Sub3: Font, % dpi("s8")
Gui, Sub3: add, Text, % dpi("x20 y+16 Section"), 字段1:
Gui, Sub3: add, DropDownList, % dpi("xs w72 vExifRename1Var")
, 无|书名|作者|出版日期|丛书||出版商|页数|制作者|生成工具1|生成工具2|关键词|描述

Gui, Sub3: add, Text, % dpi("ys Section"), 字段2:
Gui, Sub3: add, DropDownList, % dpi("xs w72 vExifRename2Var")
, 无|书名||作者|出版日期|丛书|出版商|页数|制作者|生成工具1|生成工具2|关键词|描述

Gui, Sub3: add, Text, % dpi("ys Section"), 字段3:
Gui, Sub3: add, DropDownList, % dpi("xs w72 vExifRename3Var")
, 无|书名|作者||出版日期|丛书|出版商|页数|制作者|生成工具1|生成工具2|关键词|描述

Gui, Sub3: add, Text, % dpi("ys Section"), 字段4:
Gui, Sub3: add, DropDownList, % dpi("xs w72 vExifRename4Var")
, 无|书名|作者|出版日期|丛书|出版商||页数|制作者|生成工具1|生成工具2|关键词|描述

Gui, Sub3: add, Text, % dpi("ys Section"), 字段5:
Gui, Sub3: add, DropDownList, % dpi("xs w72 vExifRename5Var")
, 无|书名|作者|出版日期||丛书|出版商|页数|制作者|生成工具1|生成工具2|关键词|描述

Gui, Sub3: add, Text, % dpi("ys Section"), 字段6:
Gui, Sub3: add, DropDownList, % dpi("xs w72 vExifRename6Var")
, 无||书名|作者|出版日期|丛书|出版商|页数|制作者|生成工具1|生成工具2|关键词|描述

Gui, Sub3: add, Text, % dpi("ys Section"), 字段7:
Gui, Sub3: add, DropDownList, % dpi("xs w72 vExifRename7Var")
, 无||书名|作者|出版日期|丛书|出版商|页数|制作者|生成工具1|生成工具2|关键词|描述

Gui, Sub3: add, Text, % dpi("ys Section"), 字段8:
Gui, Sub3: add, DropDownList, % dpi("xs w72 vExifRename8Var")
, 无||书名|作者|出版日期|丛书|出版商|页数|制作者|生成工具1|生成工具2|关键词|描述

Gui, Sub3: add, Text, % dpi("ys Section"), 字段9:
Gui, Sub3: add, DropDownList, % dpi("xs w72 vExifRename9Var")
, 无||书名|作者|出版日期|丛书|出版商|页数|制作者|生成工具1|生成工具2|关键词|描述

Gui, Sub3: add, Text, % dpi("ys Section"), 字段10:
Gui, Sub3: add, DropDownList, % dpi("xs w72 vExifRename10Var")
, 无||书名|作者|出版日期|丛书|出版商|页数|制作者|生成工具1|生成工具2|关键词|描述

Gui, Sub3: add, Text, % dpi("ys Section"), 字段11:
Gui, Sub3: add, DropDownList, % dpi("xs w72 vExifRename11Var")
, 无||书名|作者|出版日期|丛书|出版商|页数|制作者|生成工具1|生成工具2|关键词|描述

Gui, Sub3: Font, % dpi("s10")
Gui, Sub3: add, Text, % dpi("x20 y+35"), 出版日期格式：
Gui, Sub3: add, Radio, % dpi("x+10 Checked"), 年
Gui, Sub3: add, Radio, % dpi("x+10 vRadio1Sub3Var"), 年月
Gui, Sub3: add, Radio, % dpi("x+10 vRadio2Sub3Var"), 年月日
Gui, Sub3: add, Text, % dpi("x+50"), 连接符：

Gui, Sub3: Font, % dpi("Bold")
Gui, Sub3: add, Edit, % dpi("x+10 W52 H22 vSub3EditLinkVar"), _

Gui, Sub3: Font, % dpi("Norm")
Gui, Sub3: add, Text, % dpi("x20"), 重命名模板：
Gui, Sub3: add, button, % dpi("xs yp w72 h19 gSub3GUISubmitLB Section"), 生成字段
Gui, Sub3: add, button, % dpi("xp-83 w72 h19 gSub3GUITestLB"), 测试

Gui, Sub3: Font, % dpi("c99FF33")
Gui, Sub3: add, Edit, % dpi("x20 w924 h25 HwndEditExifRenameHD vReNamePattern"), %ReNamePattern%

Gui, Sub3: add, button, % dpi("xs y+34 w72 h22 gSub3guiclose Section"), 取消
Gui, Sub3: add, button, % dpi("xp-83 w72 h22 gSub3GUIYesLB"), 确定

; 修改控件颜色，RGB模式
Gui, Sub3: Color,, 202020
Gui, Sub3: Show, % dpi("w963 h280")
Return


; Sub3按钮确定
Sub3GUIYesLB:
GuiControlGet, ReNamePattern,, %EditExifRenameHD%
gosub, Sub3guiclose
Return


; Sub3按钮生成字段
Sub3GUISubmitLB:
Gui, Sub3: Submit, NoHide
ReNamePattern := ""
Loop, 11  ; exif项目相关命令
{
  ReNamePatternVar := "ExifRename" A_Index "Var"
  
  if (%ReNamePatternVar% = "书名")
   ReNamePattern .= "BookTitle" Sub3EditLinkVar
  if (%ReNamePatternVar% = "作者")
   ReNamePattern .= "BookAuthor" Sub3EditLinkVar
  if (%ReNamePatternVar% = "出版日期")
   ReNamePattern .= "BookPubDate" Sub3EditLinkVar
  if (%ReNamePatternVar% = "丛书")
   ReNamePattern .= "BookSeries" Sub3EditLinkVar
  if (%ReNamePatternVar% = "出版商")
   ReNamePattern .= "BookPublisher" Sub3EditLinkVar
  if (%ReNamePatternVar% = "页数")
   ReNamePattern .= "BookPageCount" Sub3EditLinkVar
  if (%ReNamePatternVar% = "制作者")
   ReNamePattern .= "BookCreator" Sub3EditLinkVar
  if (%ReNamePatternVar% = "生成工具1")
   ReNamePattern .= "BookProducer" Sub3EditLinkVar
  if (%ReNamePatternVar% = "生成工具2")
   ReNamePattern .= "BookCreatorTool" Sub3EditLinkVar
  if (%ReNamePatternVar% = "关键词")
   ReNamePattern .= "BookKeyWords" Sub3EditLinkVar
  if (%ReNamePatternVar% = "描述")
   ReNamePattern .= "BookDescription" Sub3EditLinkVar  
}

ReNamePattern := RTrim(ReNamePattern, Sub3EditLinkVar)
GuiControl,, %EditExifRenameHD%, %ReNamePattern%
Return


; Sub3按钮测试
Sub3GUITestLB:
Sub3GuiTestIDN = 1
Gui, Sub3: Submit, NoHide
gosub, RenameFileLB
Sub3GuiTestIDN =
return


; 关闭Sub3窗口
Sub3guiEscape:
Sub3guiclose:
GuiControlGet, ReNamePattern,, %EditExifRenameHD%
Gui, Sub3: Destroy
Return


;右键菜单_设置预设替换模板
BatchReplaceSetupLB:
Gui, main : Default
; 右键点击ListView时会记录鼠标所在列号，如果列号是1789或10列
if !ColumnNo or (RegExMatch(ColumnNo, "^(?:1|7|8|9|10)$"))  ; exif项目相关命令
{
 ToolTip, 没有选择有效列
 SetTimer, ToolTipoffLB, -1500
 return
}

Gui, Sub4: Color, 202020 
Gui, Sub4: +AlwaysOnTop
Gui, Sub4: Font, % dpi("s11 cFF9900"),微软雅黑 
Gui, Sub4: add, Text, % dpi("x330 y25"), 预设批量替换的字符
Gui, Sub4: Font, % dpi("s8")
Gui, Sub4: add, Text, % dpi("x240 Section"), 替换使用正则表达式
Gui, Sub4: Font, % dpi("s11")
Gui, Sub4: add, Text, % dpi("ys-3"), \ . * ? + [ { | ( ) ^ $ 务必转义
Gui, Sub4: Font, % dpi("s8")
Gui, Sub4: add, Text, % dpi("x40 y110 w50 Right Section"), 原字符:
Gui, Sub4: add, Text, % dpi("w50 Right"), 新字符:
Gui, Sub4: add, Text, % dpi("w50 Right"), 替换次数:
Gui, Sub4: add, Text, % dpi("y+0 w50 Right"), (-1为无限)

Gui, Sub4: Font, % dpi("cFF9900")
Gui, Sub4: add, Text, % dpi("x100 y85 Section"), 字符1:
Gui, Sub4: Font, % dpi("c99FF33")
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit1SouVar HwndSub4Edit1SouHD"),
Gui, Sub4: add, Edit, % dpi("w50 vSub4Edit1RepVar HwndSub4Edit1RepHD"),
Gui, Sub4: add, Edit, % dpi("w50 vSub4Edit1LimitVar HwndSub4Edit1LimitHD"),
Gui, Sub4: Font, % dpi("cFF9900")
Gui, Sub4: add, Text, % dpi("ys Section"), 字符2:
Gui, Sub4: Font, % dpi("c99FF33")
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit2SouVar HwndSub4Edit2SouHD"),
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit2RepVar HwndSub4Edit2RepHD"),
Gui, Sub4: add, Edit, % dpi("w50 vSub4Edit2LimitVar HwndSub4Edit2LimitHD"),
Gui, Sub4: Font, % dpi("cFF9900")
Gui, Sub4: add, Text, % dpi("ys Section"), 字符3:
Gui, Sub4: Font, % dpi("c99FF33")
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit3SouVar HwndSub4Edit3SouHD"),
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit3RepVar HwndSub4Edit3RepHD"),
Gui, Sub4: add, Edit, % dpi("w50 vSub4Edit3LimitVar HwndSub4Edit3LimitHD"),
Gui, Sub4: Font, % dpi("cFF9900")
Gui, Sub4: add, Text, % dpi("ys Section"), 字符4:
Gui, Sub4: Font, % dpi("c99FF33")
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit4SouVar HwndSub4Edit4SouHD"),
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit4RepVar HwndSub4Edit4RepHD"),
Gui, Sub4: add, Edit, % dpi("w50 vSub4Edit4LimitVar HwndSub4Edit4LimitHD"),
Gui, Sub4: Font, % dpi("cFF9900")
Gui, Sub4: add, Text, % dpi("ys Section"), 字符5:
Gui, Sub4: Font, % dpi("c99FF33")
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit5SouVar HwndSub4Edit5SouHD"),
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit5RepVar HwndSub4Edit5RepHD"),
Gui, Sub4: add, Edit, % dpi("w50 vSub4Edit5LimitVar HwndSub4Edit5LimitHD"),
Gui, Sub4: Font, % dpi("cFF9900")
Gui, Sub4: add, Text, % dpi("ys Section"), 字符6:
Gui, Sub4: Font, % dpi("c99FF33")
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit6SouVar HwndSub4Edit6SouHD"),
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit6RepVar HwndSub4Edit6RepHD"),
Gui, Sub4: add, Edit, % dpi("w50 vSub4Edit6LimitVar HwndSub4Edit6LimitHD"),
Gui, Sub4: Font, % dpi("cFF9900")
Gui, Sub4: add, Text, % dpi("ys Section"), 字符7:
Gui, Sub4: Font, % dpi("c99FF33")
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit7SouVar HwndSub4Edit7SouHD"),
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit7RepVar HwndSub4Edit7RepHD"),
Gui, Sub4: add, Edit, % dpi("w50 vSub4Edit7LimitVar HwndSub4Edit7LimitHD"),
Gui, Sub4: Font, % dpi("cFF9900")
Gui, Sub4: add, Text, % dpi("ys Section"), 字符8:
Gui, Sub4: Font, % dpi("c99FF33")
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit8SouVar HwndSub4Edit8SouHD"),
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit8RepVar HwndSub4Edit8RepHD"),
Gui, Sub4: add, Edit, % dpi("w50 vSub4Edit8LimitVar HwndSub4Edit8LimitHD"),
Gui, Sub4: Font, % dpi("cFF9900")
Gui, Sub4: add, Text, % dpi("ys Section"), 字符9:
Gui, Sub4: Font, % dpi("c99FF33")
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit9SouVar HwndSub4Edit9SouHD"),
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit9RepVar HwndSub4Edit9RepHD"),
Gui, Sub4: add, Edit, % dpi("w50 vSub4Edit9LimitVar HwndSub4Edit9LimitHD"),
Gui, Sub4: Font, % dpi("cFF9900")
Gui, Sub4: add, Text, % dpi("ys Section"), 字符10:
Gui, Sub4: Font, % dpi("c99FF33")
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit10SouVar HwndSub4Edit10SouHD"),
Gui, Sub4: add, Edit, % dpi("xs w50 vSub4Edit10RepVar HwndSub4Edit10RepHD"),
Gui, Sub4: add, Edit, % dpi("w50 vSub4Edit10LimitVar HwndSub4Edit10LimitHD"),

Gui, Sub4: add, button, % dpi("x290 y+25 w60 h25 gSub4GUISubmitLB Section"), 确定
Gui, Sub4: add, button, % dpi("x+20 ys w60 h25 gSub4guiclose"), 取消
Gui, Sub4: add, button, % dpi("x+20 ys w60 h25 gBatchReplaceReadLB"), 读取存档
Gui, Sub4: Font, % dpi("s2")
Gui, Sub4: add, Text,
Gui, Sub4: Color,, 202020
Gui, Sub4: Show, % dpi("w780")
gosub, BatchReplaceReadLB
return


; 读取存档
BatchReplaceReadLB:
Loop, parse, BatchReplaceStr1, ￠
{
 Segment := (A_Index - 1) // 3 + 1
 HwndNAME1 := "Sub4Edit" Segment "SouHD"
 HwndNAME1 := %HwndNAME1%
 HwndNAME2 := "Sub4Edit" Segment "RepHD"
 HwndNAME2 := %HwndNAME2%
 HwndNAME3 := "Sub4Edit" Segment "LimitHD"
 HwndNAME3 := %HwndNAME3%
 
 VarNAME1 := "Sub4Edit" Segment "SouVar"
 VarNAME2 := "Sub4Edit" Segment "RepVar"
 VarNAME3 := "Sub4Edit" Segment "LimitVar"

 if (A_Index = 1 || Mod(A_Index - 1, 3) = 0)
 {
  GuiControl,, %HwndNAME1%, %A_LoopField%
  %VarNAME1% := A_LoopField
 }
 if (A_Index = 2 || Mod(A_Index + 1, 3) = 0)
 {
  GuiControl,, %HwndNAME2%, %A_LoopField%
  %VarNAME2% := A_LoopField
 }
 if (Mod(A_Index, 3) = 0)
 {
  GuiControl,, %HwndNAME3%, %A_LoopField%
  %VarNAME3% := A_LoopField
 }
}
Return


; Sub4窗口-确定按钮
Sub4GUISubmitLB:
Gui, main : Default

Gui, Sub4: Submit
gosub, Sub4guiclose
BatchReplaceStr1 := "
(Join`s
" Sub4Edit1SouVar "￠" Sub4Edit1RepVar "￠" Sub4Edit1LimitVar "￠"
Sub4Edit2SouVar "￠" Sub4Edit2RepVar "￠" Sub4Edit2LimitVar "￠"
Sub4Edit3SouVar "￠" Sub4Edit3RepVar "￠" Sub4Edit3LimitVar "￠"
Sub4Edit4SouVar "￠" Sub4Edit4RepVar "￠" Sub4Edit4LimitVar "￠"
Sub4Edit5SouVar "￠" Sub4Edit5RepVar "￠" Sub4Edit5LimitVar "￠"
Sub4Edit6SouVar "￠" Sub4Edit6RepVar "￠" Sub4Edit6LimitVar "￠"
Sub4Edit7SouVar "￠" Sub4Edit7RepVar "￠" Sub4Edit7LimitVar "￠"
Sub4Edit8SouVar "￠" Sub4Edit8RepVar "￠" Sub4Edit8LimitVar "￠"
Sub4Edit9SouVar "￠" Sub4Edit9RepVar "￠" Sub4Edit9LimitVar "￠"
Sub4Edit10SouVar "￠" Sub4Edit10RepVar "￠" Sub4Edit10LimitVar "
)"
;Clipboard:=BatchReplaceStr1
Return


; 一键替换
BatchReplaceLB:
MsgBox, 4100,, `                 是否批量替换整列？
IfMsgBox, No
 Return

Gui, main : Default
; 右键点击ListView时会记录鼠标所在列号，如果列号是1789或10列
if !ColumnNo or (RegExMatch(ColumnNo, "^(?:1|7|8|9|10)$"))  ; exif项目相关命令
{
 ToolTip, 没有选择有效列
 SetTimer, ToolTipoffLB, -1500
 return
}
; 整列逐行替换
Loop, % LV_GetCount()
{
 LV_GetText(CellText, A_Index, ColumnNo)
 Loop, 10
 {
  BatchReplaceSou := "Sub4Edit" A_Index "SouVar"
  BatchReplaceSou := %BatchReplaceSou%
  BatchReplaceRep := "Sub4Edit" A_Index "RepVar"
  BatchReplaceRep := %BatchReplaceRep%
  BatchReplaceLimit := "Sub4Edit" A_Index "LimitVar"
  BatchReplaceLimit := %BatchReplaceLimit%
  
  CellText := RegExReplace(CellText, BatchReplaceSou, BatchReplaceRep,,BatchReplaceLimit)
 }
 
 LV_Modify(A_Index, "Col" ColumnNo, CellText)  ; 更改第ColumnNo列
}

; 状态栏显示未保存
++UnSaveIDN
GuiControl,,%GUIattribute2%, |  未保存
Return

; 关闭Sub4窗口
Sub4guiEscape:
Sub4guiclose:
Gui, Sub4: Destroy
Return


;右键菜单_删除文件
ContextDelFileLB:
Gui, main : Default
MsgBox, 4100,, `                      确定删除文件？
IfMsgBox, NO
 return
else
{
 SelectedRowNumber := LV_GetNext()  ; 查找选中行.
 if !SelectedRowNumber              ; 没有选中行.
 {
  MsgBox, 4096,, `                     没有选中行
  return
 }
}

RowNumber := 0  ; 这会使得首次循环从顶部开始搜索.
Loop
{
 RowNumber := LV_GetNext(RowNumber - 1)
 if not RowNumber  ; 上面返回零, 所以没有更多选择的行了.
  break
 LV_GetText(FilePath, RowNumber, 8)
 FileRecycle, %FilePath%
 if ErrorLevel
  MsgBox, 4096,, 无法删除 “%FilePath%”
 
 ; 同时删除封面
 CoverPath := RegExReplace(FilePath, "\.pdf$", ".jpg")
 if FileExist(CoverPath)
 {
  FileRecycle, %CoverPath%
  if ErrorLevel
   MsgBox, 4096,, 无法删除 “%CoverPath%”
 }
 
 LV_Delete(RowNumber)  ; 从 ListView 中删除行.
}
FileNo := LV_GetCount()
GuiControl,, %GUIattribute1%, 已读取 共%FileNo%项

; 隐藏之前显示的封面
GuiControl,, %Image2HD%
GuiControl, Hide, %ButtChangeImageHD%
GuiControl, Hide, %ButtLastImageHD%
GuiControl, Hide, %ButtNextImageHD%

if WinExist("ahk_exe" CropImageToolPath)
 run, % CropImageCallPath,, Hide  ; 打开图像
Return


; 右键菜单_删除所选行
ContextClearRowsLB:    ; 用户在上下文菜单中选择了 "Clear".
Gui,main:Default
RowNumber := 0         ; 这会使得首次循环从顶部开始搜索.
Loop
{
 ; 由于删除了一行使得此行下面的所有行的行号都减小了,
 ; 所以把行号减 1, 以免有些行不在范围内，导致搜索遗漏
 RowNumber := LV_GetNext(RowNumber - 1)
 if not RowNumber      ; 上面返回零, 所以没有更多选择的行了.
  break
 LV_Delete(RowNumber)  ; 从 ListView 中删除行.
}
FileNo := LV_GetCount()
GuiControl,, %GUIattribute1%, 已读取 共%FileNo%项
return


; 右键菜单_编辑整列
ModifyColumnLB:
Gui,main:Default
; 右键点击ListView时会记录鼠标所在列号，如果列号是1789或10列
if !ColumnNo or (RegExMatch(ColumnNo, "^(?:1|7|8|9|10)$"))  ; exif项目相关命令
{
 ToolTip, 没有选择有效列
 SetTimer, ToolTipoffLB, -1500
 return
}

LV_GetText(HeaderText, 0, ColumnNo) ; 获取列标题

Gui +OwnDialogs
InputBox, ColumnModifyText, 编辑整列
, 鼠标所在为第 %ColumnNo% 列，列标题为「%HeaderText%」，请输入整列要更改的文字
,, 480, 120,,,,, %Clipboard%

if ErrorLevel   ; 当用户按下取消按钮时 ErrorLevel 值被设置为 1, 按下确定时值为 0
{
 ColumnModifyText =
 return
}

LV_Modify( 0, "Col" ColumnNo, ColumnModifyText)  ; 更改第ColumnNo列

ColumnModifyText := ColumnNo := ""

; 状态栏显示未保存
++UnSaveIDN
GuiControl,,%GUIattribute2%, |  未保存
Return


; 右键菜单_清空整列
EmptyColumnLB:
Gui,main:Default
; 右键点击ListView时会记录鼠标所在列号，如果列号是1789或10列
if !ColumnNo or (RegExMatch(ColumnNo, "^(?:1|7|8|9|10)$"))  ; exif项目相关命令
{
 ToolTip, 没有选择有效列
 SetTimer, ToolTipoffLB, -1500
 return
}

LV_Modify( 0, "Col" ColumnNo, "")  ; 更改第ColumnNo列
ColumnModifyText := ColumnNo := ""

; 状态栏显示未保存
++UnSaveIDN
GuiControl,,%GUIattribute2%, |  未保存
return


; 右键菜单_编辑选中行
SelectedRowModifyLB:
Gui,main:Default

; 右键点击ListView时会记录鼠标所在列号，如果列号是1789或10列
if !ColumnNo or (RegExMatch(ColumnNo, "^(?:1|7|8|9|10)$"))  ; exif项目相关命令
{
 ToolTip, 没有选择有效列
 SetTimer, ToolTipoffLB, -1500
 return
}

SelectedRowNumber := LV_GetNext()  ; 查找选中行.
if !SelectedRowNumber              ; 没有选中行.
{
 ToolTip, `        没有选中行`        `
 SetTimer, ToolTipOffLB, -1500
 return
}

LV_GetText(HeaderText, 0, ColumnNo)

Gui +OwnDialogs
InputBox, ColumnModifyText, 编辑多行
, 鼠标所在为第 %ColumnNo% 列，列标题为「%HeaderText%」，请输入选中行要更改的文字
,, 480, 120,,,,, %Clipboard%

if ErrorLevel   ; 当用户按下取消按钮时 ErrorLevel 值被设置为 1, 按下确定时值为 0
{
 ColumnModifyText =
 return
}

RowNumber := 0  ; 这会使得首次循环从顶部开始搜索.
Loop
{
 RowNumber := LV_GetNext(RowNumber)
 if !RowNumber  ; 上面返回零, 所以没有更多选择的行了.
  break
 LV_Modify(RowNumber, "Col" ColumnNo, ColumnModifyText)
}

ColumnModifyText := ColumnNo := ""

; 状态栏显示未保存
++UnSaveIDN
GuiControl,,%GUIattribute2%, |  未保存
return


; 右键菜单_替换选中行字符
SelectedRowReplaceLB:
Gui,main:Default
; 右键点击ListView时会记录鼠标所在列号，如果列号是1789或10列
if !ColumnNo or (RegExMatch(ColumnNo, "^(?:1|7|8|9|10)$"))  ; exif项目相关命令
{
 ToolTip, 没有选择有效列
 SetTimer, ToolTipoffLB, -1500
 return
}

SelectedRowNumber := LV_GetNext()  ; 查找选中行.
if !SelectedRowNumber              ; 没有选中行.
{
 ToolTip, `        没有选中行`        `
 SetTimer, ToolTipOffLB, -1500
 return
}

LV_GetText(HeaderText, 0, ColumnNo) ; 获取列标题

Gui, Sub1: Color, 202020
Gui, Sub1: +AlwaysOnTop
Gui, Sub1: Font, % dpi("s11"),微软雅黑 
Gui, Sub1: Font, cFF9900
Gui, Sub1: add, Text, HwndSub1Text1HD, 鼠标所在为第 %ColumnNo% 列，列标题为「%HeaderText%」
GuiControlGet, Sub1Text1Pos, Pos, %Sub1Text1HD%
Sub1Text1PosX := ((500* GuiDPI / 96) - Sub1Text1PosW) / 2 
GuiControl, move, %Sub1Text1HD%, x%Sub1Text1PosX%
Gui, Sub1: Font, % dpi("s9")
Gui, Sub1: add, Text, % dpi("x40 vSub1ButtRegex1Var gSub1ButtRegexLB Section"), 删除前3字符
Gui, Sub1: add, Text, % dpi("ys vSub1ButtRegex2Var gSub1ButtRegexLB"), 删除后3字符
Gui, Sub1: add, Text, % dpi("ys vSub1ButtRegex3Var gSub1ButtRegexLB"), 替换[
Gui, Sub1: add, Text, % dpi("ys vSub1ButtRegex4Var gSub1ButtRegexLB"), 替换]
Gui, Sub1: add, Text, % dpi("ys vSub1ButtRegex5Var gSub1ButtRegexLB"), 替换【
Gui, Sub1: add, Text, % dpi("ys vSub1ButtRegex6Var gSub1ButtRegexLB"), 替换】
Gui, Sub1: add, Text, % dpi("ys vSub1ButtRegex7Var gSub1ButtRegexLB"), 前增
Gui, Sub1: add, Text, % dpi("ys vSub1ButtRegex8Var gSub1ButtRegexLB"), 后增

Gui, Sub1: add, Text, % dpi("x40 y+30 section"), 查找字符：
Gui, Sub1: Font, % dpi("s11")
Gui, Sub1: add, Checkbox, % dpi ("ys-3 vSub1Checkbox1Var hwndSub1Checkbox1HD checked1"), 使用原义字符
Gui, Sub1: Font, % dpi("s9")
Gui, Sub1: Font, c99FF33
Gui, Sub1: add, Edit, % dpi("x40 w418 r1 HwndSub1EditSearchHD vSub1EditSearchVar"), %Clipboard%
Gui, Sub1: Font, cFF9900
Gui, Sub1: add, Text,, 替换字符：
Gui, Sub1: Font, c99FF33
Gui, Sub1: add, Edit, % dpi("w418 HwndSub1EditReplaceHD vSub1EditReplaceVar")
Gui, Sub1: add, button, % dpi("x180 y+20 w60 h25 gSub1GUISubmitLB Section"), 替换
Gui, Sub1: add, button, % dpi("x+20 ys w60 h25 gSub1guiclose"), 关闭
Gui, Sub1: Font, % dpi("s1")
Gui, Sub1: add, Text,
; 修改控件颜色，RGB模式
Gui, Sub1: Color,, 202020
Gui, Sub1: Show, % dpi("w500")
Return


; 快捷输入正则
Sub1ButtRegexLB:
CoordMode, ToolTip, window
if (A_GuiControl = "Sub1ButtRegex1Var")
{
 GuiControl,, %Sub1EditSearchHD%, ^...
 GuiControl,, %Sub1EditReplaceHD%,
 GuiControl,, %Sub1Checkbox1HD%, 0
 ToolTip, 提示：一个点"."代表一个字符, 0, 0
 SetTimer, ToolTipoffLB, -1500
}

if (A_GuiControl = "Sub1ButtRegex2Var")
{
 GuiControl,, %Sub1EditSearchHD%, ...$
 GuiControl,, %Sub1EditReplaceHD%,
 GuiControl,, %Sub1Checkbox1HD%, 0
 ToolTip, 提示：一个点"."代表一个字符, 0, 0
 SetTimer, ToolTipoffLB, -1500
}

if (A_GuiControl = "Sub1ButtRegex3Var")
{
 GuiControl,, %Sub1EditSearchHD%, [
 GuiControl,, %Sub1EditReplaceHD%, (
 GuiControl,, %Sub1Checkbox1HD%, 1
}

if (A_GuiControl = "Sub1ButtRegex4Var")
{
 GuiControl,, %Sub1EditSearchHD%, ]
 GuiControl,, %Sub1EditReplaceHD%, )
 GuiControl,, %Sub1Checkbox1HD%, 1
}

if (A_GuiControl = "Sub1ButtRegex5Var")
{
 GuiControl,, %Sub1EditSearchHD%, 【
 GuiControl,, %Sub1EditReplaceHD%, %A_Space%
 ToolTip, 提示：替换为空格, 0, 0
 SetTimer, ToolTipoffLB, -1500
}

if (A_GuiControl = "Sub1ButtRegex6Var")
{
 GuiControl,, %Sub1EditSearchHD%, 】
 GuiControl,, %Sub1EditReplaceHD%, %A_Space%
 ToolTip, 提示：替换为空格, 0, 0
 SetTimer, ToolTipoffLB, -1500
}

if (A_GuiControl = "Sub1ButtRegex7Var")
{
 GuiControl,, %Sub1EditSearchHD%, ^
 GuiControl,, %Sub1EditReplaceHD%, %Clipboard%
 GuiControl,, %Sub1Checkbox1HD%, 0
 ToolTip, 提示：^表示在字符串前面, 0, 0
 SetTimer, ToolTipoffLB, -1500
}

if (A_GuiControl = "Sub1ButtRegex8Var")
{
 GuiControl,, %Sub1EditSearchHD%, $
 GuiControl,, %Sub1EditReplaceHD%, %Clipboard%
 GuiControl,, %Sub1Checkbox1HD%, 0
 ToolTip, 提示：$表示在字符串后面, 0, 0
 SetTimer, ToolTipoffLB, -1500
}

CoordMode, ToolTip, Screen
return


; Sub1替换按钮
Sub1GUISubmitLB:
Gui, main: Default
Gui, Sub1: Submit, NoHide

; 右键点击ListView时会记录鼠标所在列号，如果列号是1789或10列
if !ColumnNo or (RegExMatch(ColumnNo, "^(?:1|7|8|9|10)$"))  ; exif项目相关命令
{
 ToolTip, 没有选择有效列
 SetTimer, ToolTipoffLB, -1500
 return
}

SelectedRowNumber := LV_GetNext()  ; 查找选中行.
if !SelectedRowNumber              ; 没有选中行.
{
 ToolTip, `        没有选中行`        `
 SetTimer, ToolTipOffLB, -1500
 return
}

; 给正则中的特殊字符转义（使用原义）
if Sub1Checkbox1Var
 Sub1EditSearchVar := "\Q" Sub1EditSearchVar "\E"

SelectedRowNumber := 0  ; 这会使得首次循环从顶部开始搜索.
Loop
{
 SelectedRowNumber := LV_GetNext(SelectedRowNumber)
 if !SelectedRowNumber  ; 上面返回零, 所以没有更多选择的行了.
  break
 LV_GetText(CellText, SelectedRowNumber, ColumnNo)
 CellText := RegExReplace(CellText, Sub1EditSearchVar, Sub1EditReplaceVar)
 LV_Modify(SelectedRowNumber, "Col" ColumnNo, CellText)
}

; 状态栏显示未保存
++UnSaveIDN
GuiControl,,%GUIattribute2%, |  未保存
Return

; 关闭Sub1窗口
Sub1guiEscape:
Sub1guiclose:
Gui, Sub1: Destroy
Return


; 复制元数据
CopyExifLB:
Gui, main : Default

SelectedRowNumber := LV_GetNext()  ; 查找选中行.
if !SelectedRowNumber              ; 没有选中行.
{
 ToolTip, `        没有选中行`        `
 SetTimer, ToolTipOffLB, -1500
 return
}

BookExifSave := ""
Loop, 18                           ; 根据有多少列，循环多少次      ; exif项目相关命令
{
 LV_GetText(CellText, SelectedRowNumber, A_Index)
 BookExifSave .= CellText "`n"
}

Clipboard := RTrim(BookExifSave, "`n")
BookExifSave := ""
Return


; 粘贴元数据
PasteExifLB:
Gui, main : Default

SelectedRowNumber := LV_GetNext()  ; 查找选中行.
if !SelectedRowNumber              ; 没有选中行.
{
 ToolTip, `        没有选中行`        `
 SetTimer, ToolTipOffLB, -1500
 return
}

if !DllCall("IsClipboardFormatAvailable", "uint", 1)
{
 ToolTip, ` `n  剪贴板没有文字  `n `
 SetTimer, ToolTipoffLB, -1500
 return
}

; 先缓存该行元数据
BookExifSave := ""
Loop, 18                           ; 根据有多少列，循环多少次       ; exif项目相关命令
{
 LV_GetText(CellText, SelectedRowNumber, A_Index)
 BookExifSave .= CellText "`n"
}

TempVar := UnSaveIDN               ; 暂存UnSaveIDN值，后面比对

; 逐列粘贴元数据
Loop, Parse, Clipboard, `n, `r
{
 if A_Index not in 1,7,8,9         ; 跳过一些文件属性列
  LV_Modify(SelectedRowNumber, "Col" A_Index, A_LoopField)
 
 if (A_Index = 2)                  ; 如果轮到第二列修改，则证明是有效粘贴
 {
 ; 状态栏显示未保存
 ++UnSaveIDN                       ; 仅当空变量在一行中单独使用时, 运算符 ++ 和 -- 才把它们视为零;
 GuiControl,,%GUIattribute2%, |  未保存
 }
}

; 如果判断符和原来的值一样，则证明没有粘贴
if (TempVar = UnSaveIDN)
{
 ToolTip, ` `n  未粘贴  `n `
 SetTimer, ToolTipoffLB, -1500
}
; 有粘贴则记录行索引（行号会变，但行索引即第一列序号唯一）
else
{
 LV_GetText(RowIndex1, SelectedRowNumber)    ; 撤销粘贴元数据会检查RowIndex1
 gosub, AutoColSizeLB
}
Return


; 撤销粘贴元数据
UndoPasteExifLB:
Gui,main:Default
; 如果行索引为空，则返回
if !RowIndex1
 Return

; 查找原来的行索引，并获取新行号
Loop, % LV_GetCount()
{
 LV_GetText(RowIndex2, A_Index)
 if (RowIndex2 = RowIndex1)
 {
  SelectedRowNumber := A_Index
  break
 }
}

; 逐列粘贴元数据
Loop, Parse, BookExifSave, `n, `r
{
 if A_Index not in 1,7,8,9         ; 跳过一些文件属性列
  LV_Modify(SelectedRowNumber, "Col" A_Index, A_LoopField)
}

; 清空缓存的元数据和行索引
BookExifSave := RowIndex1 := RowIndex2 := ""

; 判断符减一，如果判断符为0，则清空未保存提示
--UnSaveIDN
if !UnSaveIDN
 GuiControl,,%GUIattribute2%, |
return


; 文件名导入名称栏
FileNameAsExifNameLB:
Gui,main:Default
SelectedRowNumber := 0   ; 这会使得首次循环从顶部开始搜索.
Loop
{
 SelectedRowNumber := LV_GetNext(SelectedRowNumber)
 if !SelectedRowNumber   ; 上面返回零, 所以没有更多选择的行了.
  break
 LV_GetText(FilePath, SelectedRowNumber, 8)
 SplitPath, FilePath,,,, FileName
 
 LV_Modify(SelectedRowNumber, "Col" 2, FileName)
}

gosub, AutoColSizeLB

++UnSaveIDN              ; 判断符加一，如果判断符为0，则清空未保存提示
GuiControl,,%GUIattribute2%, |  未保存
return


; 显示元数据
ShowExifLB:
Gui,main:Default
SelectedRowNumber := LV_GetNext()  ; 查找选中行.
if !SelectedRowNumber              ; 没有选中行.
 return

LV_GetText(FilePath, SelectedRowNumber, 8)
SplitPath, FilePath, FileName
SubCode := "-a -u -g1 " """" FileName """"
HideCmdIDN = 1

gosub, RunCMDLB
HideCmdIDN =
return


; 切换元数据热键
ExifHotkeyLB:
if !ExifHotkeyIDN
{
ExifHotkeyIDN = 1 
Hotkey, z, CopyExifLB, On
hotkey, x, PasteExifLB, On
Menu, Lv2menu1, Rename, 元数据热键, 元数据热键√ z/x
} else {
ExifHotkeyIDN =
Hotkey, z, CopyExifLB, Off
Hotkey, x, PasteExifLB, Off
Menu, Lv2menu1, Rename, 元数据热键√ z/x, 元数据热键
}
return


; 右键菜单_生成封面
GenerateCoverLB:
Gui, main : Default

SelectedRowNumber := LV_GetNext()  ; 查找选中行.
if !SelectedRowNumber              ; 没有选中行.
{
 ToolTip, `        没有选中行`        `
 SetTimer, ToolTipOffLB, -1500
 return
}

; 路径是否存在
if !LoadFolder
{
 ToolTip, 输入路径为空
 SetTimer, ToolTipOffLB, -1500
 return
}

LV_GetText(BookTitle, SelectedRowNumber, 2)  ; 获取第2个字段的文本.
LV_GetText(FilePath, SelectedRowNumber, 8)  ; 获取第8个字段的文本.
Gui, Sub2: Color, 202020 
Gui, Sub2: +AlwaysOnTop
Gui, Sub2: Font, % dpi("s10"),微软雅黑 
Gui, Sub2: Font, cFF9900
Gui, Sub2: add, Text, % dpi("y15 HwndSub2Text1HD"), 书名为「%BookTitle%」
GuiControlGet, Sub2Text1Pos, Pos, %Sub2Text1HD%
Sub2Text1PosX := ((385* GuiDPI / 96) - Sub2Text1PosW) / 2 
GuiControl, move, %Sub2Text1HD%, x%Sub2Text1PosX%
Gui, Sub2: add, Text, % dpi("x133"), 请输入要提取第几页：
Gui, Sub2: Font, c99FF33
Gui, Sub2: add, Edit, % dpi("x163 w60 vPageNoVar number"), 
; 0x80 省略伙伴控件中正常出现在每三位十进制数间的千位分隔符
Gui, Sub2: Add, UpDown, % dpi("Range1-9999 0x80"), 1
Gui, Sub2: add, button, % dpi("x122 y+20 w60 h25 gSub2GUISubmitLB Section"), 确定
Gui, Sub2: add, button, % dpi("x+20 ys w60 h25 gSub2guiclose"), 取消
; 修改控件颜色，RGB模式
Gui, Sub2: Color,, 202020
Gui, Sub2: Show, % dpi("w385 h150")
Return

; Sub2确定按钮
Sub2GUISubmitLB:
Gui, Sub2: Submit
gosub, Sub2guiclose
RunWait, %ComSpec% /c %MutoolPath% convert -o "%CoverPath%" -O width=800 "%FilePath%" %PageNoVar% ,, Hide
FileMove, %LoadFolder%\%CoverNameNoExt%1.jpg, %CoverPath%, 1
GuiControl,, %Image2HD%, *w%Image2Width% *h-1 %CoverPath%
Return


; 关闭Sub2窗口
Sub2guiEscape:
Sub2guiclose:
Gui, Sub2: Destroy
Return


; 右键菜单_保存选中行
SaveSelectedRowLB:
Gui, main : Default

SelectedRowNumber := LV_GetNext()  ; 查找选中行.
if !SelectedRowNumber              ; 没有选中行.
{
 ToolTip, `        没有选中行`        `
 SetTimer, ToolTipOffLB, -1500
 return
}


; 第1种方法：删除未选中行，跳过选中行
RowNumber := RowNumberTotal := 0
RowToDelNo = 1

Loop, % LV_GetCount()
{
 ; 从上一个选中行号开始搜索下一个选中行号
 RowNumber := LV_GetNext(RowNumber)
 ;MsgBox % RowNumber "  " RowToDelNo
 
 ; 删除行号和选中行号相同，则删除行号加1，选中行号累积数加1（可以理解为：选中行累积置顶行数）
 ; 跳过删除
 if (RowToDelNo = RowNumber)
 {
  RowToDelNo++
  RowNumberTotal++
  continue
 }
 ; 删除行号和选中行号不同，选中行号等于累积数（下次从累计行号开始搜索选中行）
 ; 进行删除
 else
 {
  RowNumber := RowNumberTotal
  ;MsgBox % "else   " RowNumber "  " RowToDelNo
  LV_Delete(RowToDelNo)
 }
}

FileNo := LV_GetCount()
GuiControl,, %GUIattribute1%, 已读取 共%FileNo%项

; 执行保存（界面可能要刷新时间，此处设等待0.5秒）
SetTimer, SaveEditingLB, -500
return


;右键菜单_自动列宽
AutoColSizeLB:
Gui,main:Default
LV_ModifyCol(1, 60)
LV_ModifyCol(2, "AutoHdr") 
LV_ModifyCol(3, "AutoHdr") 
LV_ModifyCol(4, "90") 
LV_ModifyCol(5, "AutoHdr") 
LV_ModifyCol(6, "AutoHdr") 
LV_ModifyCol(7, "60 Center")
LV_ModifyCol(8, "Auto Left") 
LV_ModifyCol(9, "80 Right")
LV_ModifyCol(10, "60 Right") 
LV_ModifyCol(11, "AutoHdr") 
LV_ModifyCol(12, "AutoHdr") 
LV_ModifyCol(13, "AutoHdr") 
LV_ModifyCol(14, "AutoHdr") 
LV_ModifyCol(15, "AutoHdr") 
LV_ModifyCol(16, "AutoHdr") 
LV_ModifyCol(17, "AutoHdr") 
LV_ModifyCol(18, "400") 
return


;右键菜单_最小列宽
SmallColSizeLB:
Gui,main:Default
LV_ModifyCol(1, 60)
LV_ModifyCol(2, 60)
LV_ModifyCol(3, 60)
LV_ModifyCol(4, 60)
LV_ModifyCol(5, 60) 
LV_ModifyCol(6, 60) 
LV_ModifyCol(7, 60)
LV_ModifyCol(8, 60)
LV_ModifyCol(9, 60)
LV_ModifyCol(10, 60) 
LV_ModifyCol(11, 60) 
LV_ModifyCol(12, 60)
LV_ModifyCol(13, 60)
LV_ModifyCol(14, 60) 
LV_ModifyCol(15, 60) 
LV_ModifyCol(16, 60)
LV_ModifyCol(17, 60)
LV_ModifyCol(18, 60)
return


;右键菜单_窗口置顶
ontopLB:
if ontopidn
{
Gui,main : -AlwaysOnTop
ontopidn =
Menu, MyContextMenu, Rename, 窗口置顶√, 窗口置顶
return
}

if !ontopidn
{
Gui,main : +AlwaysOnTop
ontopidn = 1
Menu, MyContextMenu, Rename, 窗口置顶, 窗口置顶√
return
}


;重启
ReloadLB:
if UnSaveIDN
 {
  MsgBox, 4100,, 有编辑尚未保存，确定重启？
  IfMsgBox, No
   return
 }
Reload
Return


; 热键_全选
^a::
SelectAllLB:
Gui, main : Default
;高亮全部项目
if !LoadFolder
 Return

FileNo := LV_GetCount()
Loop
{
 LV_Modify(A_Index, "Select")
 if (A_Index = FileNo )
  break
}
GuiControl,,%GUIattribute3%, |
Return


; 说明
InfoLB:
InfoText=
(
保存时，「生成日期」，「制作日期」和「元数据日期」列填写“now”，可自动生成当前时间。`n`n
制作者原文为Creator，生成工具1原文为Producer，生成工具2原文为Creator Tool。`n`n
热键：Del 删除文件 
`           Ctrl+a 全选行
`           z 复制元数据（仅当启用时） 
`           x 粘贴元数据（仅当启用时）
)
MsgBox, 4096,, %InfoText%
return


; 右键菜单_窗口缩放
setGuiDPILB:
GuiDPIVar := ((A_ThisMenuItemPos-1) * .25 + 1) * 96    ; 根据右键菜单选项确定dpi数值

; 如果和当前dpi相同，则返回
if (GuiDPI = GUIdpivar)
{
 SetTimer, ToolTipoffLB, -1500  
 ToolTip, DPI和当前相同
 return
}

; 如果有编辑未保存
if UnSaveIDN
{
 MsgBox, 4100,,`          有编辑尚未保存，确定重启窗口？
 IfMsgBox, No
  return
 Else {
 UnSaveIDN := BookExifSave := RowIndex1 := ""  ; 清空原来的“粘贴元数据历史记录”
 LV_InCellEdit.Changed.Remove(ListViewHD, "")
 }
} else {
 MsgBox,4100,提示,`          即将重启窗口，是否确定？
 IfMsgBox, No
  Return 
}

GuiDPI := GUIdpivar
Menu, MyContextMenu, Delete     ; 销毁右键菜单

Gui, main: Destroy              ; 销毁原来窗口并重绘窗口
gosub, GUIdrawLB
return


; 主窗口缩放
mainGuiSize:  ; 扩大或缩小 ListView 来响应用户对窗口大小的改变.
if (A_EventInfo = 1)  ; 窗口被最小化了. 无需进行操作.
 return

;调整listview宽高
listviewwidminus:=271*GuiDPI/96   ; 30       ;减数要适配不同dpi，30为96dpi时的减数，即放大率为1时减数为27
listviewhgtminus:=91*GuiDPI/96   ; 66
GuiControl, Move, ListViewVar,% "W" (A_GuiWidth - listviewwidminus) " H" (A_GuiHeight - listviewhgtminus)

;调整「出版商列表」宽度
DDLPublisherwidminus := (656*GuiDPI/96)      ; 624
GuiControl, Move, %DDLPublisherHD%, % "W" (A_GuiWidth - DDLPublisherwidminus)

;调整「按钮_复制出版商」横向位置
ButtonCopyPubXminus := 400*GuiDPI/96
GuiControl, Move, %ButtonCopyPubHD%, % " x" (A_GuiWidth - ButtonCopyPubXminus)

;调整「编辑栏-地址栏」宽度
EditPathwidminus := 345*GuiDPI/96       ; 有前往目录按钮时为188
GuiControl, Move, %EditPathHD%, % "W" (A_GuiWidth - EditPathwidminus)

;调整「按钮_保存更改」横向位置
ButtSaveRenameXminus := 318*GuiDPI/96
GuiControl, Move, %ButtSaveRenameHD%, % " x" (A_GuiWidth - ButtSaveRenameXminus)

;调整「补位方块」横向位置
GuiControl, Move, %SquareHD%, % "x" (A_GuiWidth - ButtSaveRenameXminus)

;调整「图像1」横向位置
Image1Plus := 25*GuiDPI/96
GuiControl, Move, %Image1HD%, % "x" (A_GuiWidth - ButtSaveRenameXminus + Image1Plus)
GuiControl,, %Image1HD%, *w-1 *h%Image1Height% %A_ScriptDir%\Image.png

;调整「图像2」横向位置
Image2Minus := (247 * GuiDPI / 96)
GuiControl, Move, %Image2HD%, % "x" (A_GuiWidth - Image2Minus)

;调整「按钮_上一张」横向位置
GuiControl, Move, %ButtLastImageHD%, % "x" (A_GuiWidth - Image2Minus)

;调整「按钮_下一张」横向位置
GuiControl, Move, %ButtNextImageHD%, % "x" (A_GuiWidth - Image2Minus + ButtLastWidth + 30)

;调整「按钮_更改图像」横向位置
GuiControl, Move, %ButtChangeImageHD%, % "x" (A_GuiWidth - Image2Minus)

;调整「底部状态文字1」纵向位置
GUIattribute1Yminus := (27*GuiDPI/96)
GuiControl, Move, %GUIattribute1%, % " y" (A_GuiHeight - GUIattribute1Yminus)

;调整「底部状态文字2」纵向位置
GuiControl, Move, %GUIattribute2%, % " y" (A_GuiHeight - GUIattribute1Yminus)

;调整「底部状态文字2」纵向位置
GuiControl, Move, %GUIattribute3%, % " y" (A_GuiHeight - GUIattribute1Yminus)

; 重画窗口
WinSet, Redraw ,, ahk_id %GUIMainHD%
return


;主窗口关闭
mainGuiClose:  ; 当窗口关闭时, 退出脚本
ExitApp
return


; 运行ExifTool
RunCMDLB:
; 合成命令
if !CMDcode
{
 CMDcode=
 (join&
 cd /d %LoadFolder%
 %ExiftoolPath% %OverWrite% %SubCode%
 )
}
; 注意，要先cd到目标文件夹下再执行exiftool，因为其对中文（或非英文语种）支持不佳。
; 先cd到目标文件夹，则不必指定长路径，因此路径带中文也不会有问题。文件名则由脚本临时改名。

; 运行命令
OutputVar := ""

if !HideCmdIDN
 ;RunWait, %ComSpec% /c %CMDcode% ,,Hide
 OutputVar := StdoutToVar_CreateProcess(Comspec " /c " CMDcode)
else
 Run, %ComSpec% /k %CMDcode%

CMDcode := ""

; 查找并关闭正在运行的 ExifTool 进程
Process, Close, exiftool.exe
Return


; 关闭tooltip
ToolTipoffLB:
ToolTip
return



; 函数名称：退出功能
; 脚本退出时，执行的命令
ExitFunc(ExitReason)
{
 global
 if ExitReason not in Reload
 {
  if UnSaveIDN
  {
   MsgBox, 4100,,`      有编辑尚未保存，确定退出？
   IfMsgBox, No
    return 1
  }
  WinClose, ahk_exe %CropImageToolPath%
 }
}


; 函数名称：转换字符为日期格式的纯数字
ConvertToYearMonth(num) {
 ; 将-、_、、:、/或空格等分隔符转为0
 pattern := "(\d{4})([-/_: ])(\d{1,2})([-/_: ])?(\d{1,2})?"
 if RegExMatch(num, pattern, Matchlist)
 {
  if Matchlist3 and (strlen(Matchlist3) = 1)
   deli1 = 0
  if Matchlist5 and (strlen(Matchlist5) = 1)
   deli2 = 0
  
  num := RegExReplace(num, pattern, "$1" deli1 "$3" deli2 "$5")
 }
 
 ; 使用正则表达式删除除数字外的字符
 num := RegExReplace(num, "\D")

 return num
}


; 函数名称：定制DPI
; 窗口DPI不随系统变化，而是根据用户指定显示
/*
Name             : DPI
Purpose          : Return scaling factor or calculate position/values for AHK controls (font size, position (x y), width, height)
Version          : 0.31
Source           : https://github.com/hi5/dpi
AutoHotkey Forum : https://autohotkey.com/boards/viewtopic.php?f=6&t=37913
License          : see license.txt (GPL 2.0)
Documentation    : See readme.md @ https://github.com/hi5/dpi

History:

* v0.31: refactored "process" code, just one line now
* v0.3: - Replaced super global variable ###dpiset with static variable within dpi() to set dpi
        - Removed r parameter, always use Round()
        - No longer scales the Rows option and others that should be skipped (h-1, *w0, hwnd etc)
* v0.2: public release
* v0.1: first draft

*/

DPI(in="",setdpi=1)
	{
	 static dpi:=1
	 if (setdpi <> 1)
		dpi:=setdpi
	 RegRead, AppliedDPI, HKEY_CURRENT_USER, Control Panel\Desktop\WindowMetrics, AppliedDPI
	 ; If the AppliedDPI key is not found the default settings are used.
	 ; 96 is the default value.
	 if (ErrorLevel=1) OR (AppliedDPI=96)
		AppliedDPI:=96
	 if (dpi <> 1)
		AppliedDPI:=dpi
	 factor:=AppliedDPI/96
	 if !in
		Return factor
	 Loop, parse, in, %A_Space%%A_Tab%
		{
		 option:=A_LoopField
		 if RegExMatch(option,"i)(w0|h0|h-1|xp|yp|xs|ys|xm|ym)$") or RegExMatch(option,"i)(icon|hwnd)") ; these need to be bypassed
			out .= option A_Space
		 else if RegExMatch(option,"i)^\*{0,1}(x|xp|y|yp|w|h|s)[-+]{0,1}\K(\d+)",number) ; should be processed
			out .= StrReplace(option,number,Round(number*factor)) A_Space
		 else ; the rest can be bypassed as well (variable names etc)
			out .= option A_Space
		}
	 Return Trim(out)
	}

LV_SubitemHitTest(ListViewHD) {
   ; To run this with AHK_Basic change all DllCall types "Ptr" to "UInt", please.
   ; ListViewHD - ListView's HWND
   Static LVM_SUBITEMHITTEST := 0x1039
   VarSetCapacity(POINT, 8, 0)
   ; Get the current cursor position in screen coordinates
   DllCall("User32.dll\GetCursorPos", "Ptr", &POINT)
   ; Convert them to client coordinates related to the ListView
   DllCall("User32.dll\ScreenToClient", "Ptr", ListViewHD, "Ptr", &POINT)
   ; Create a LVHITTESTINFO structure (see below)
   VarSetCapacity(LVHITTESTINFO, 24, 0)
   ; Store the relative mouse coordinates
   NumPut(NumGet(POINT, 0, "Int"), LVHITTESTINFO, 0, "Int")
   NumPut(NumGet(POINT, 4, "Int"), LVHITTESTINFO, 4, "Int")
   ; Send a LVM_SUBITEMHITTEST to the ListView
   SendMessage, LVM_SUBITEMHITTEST, 0, &LVHITTESTINFO, , ahk_id %ListViewHD%
   ; If no item was found on this position, the return value is -1
   If (ErrorLevel = -1)
      Return 0
   ; Get the corresponding subitem (column)
   Subitem := NumGet(LVHITTESTINFO, 16, "Int") + 1
   Return Subitem
}


; SetTaskbarProgress  -  Windows 7+
;  by lexikos, modified by gwarble for U64,U32,A32 compatibility
;
; pct    -  A number between 0 and 100 or a state value (see below).
; state  -  "N" (normal), "P" (paused), "E" (error) or "I" (indeterminate).
;           If omitted (and pct is a number), the state is not changed.
; hwnd   -  The hWnd of the window which owns the taskbar button.
;           If omitted, the Last Found Window is used.
;
SetTaskbarProgress(pct, state="", hwnd="") {
 static tbl, s0:=0, sI:=1, sN:=2, sE:=4, sP:=8
 if !tbl
  Try tbl := ComObjCreate("{56FDF344-FD6D-11d0-958A-006097C9A090}"
                        , "{ea1afb91-9e28-4b86-90e9-9e9f8a5eefaf}")
  Catch 
   Return 0
 If hwnd =
  hwnd := WinExist()
 If pct is not number
  state := pct, pct := ""
 Else If (pct = 0 && state="")
  state := 0, pct := ""
 If state in 0,I,N,E,P
  DllCall(NumGet(NumGet(tbl+0)+10*A_PtrSize), "uint", tbl, "uint", hwnd, "uint", s%state%)
 If pct !=
  DllCall(NumGet(NumGet(tbl+0)+9*A_PtrSize), "uint", tbl, "uint", hwnd, "int64", pct*10, "int64", 1000)
Return 1
}


; 出版社列表
; 使用延续片段无法生成总长度超过 16,383 字符的行，
; 解决此问题的一种方法是把一系列内容连接到变量中。
PublisherListLB:
PublisherList =
(
A 安徽美术出版社||A 安徽科学技术出版社|B 北京大学出版社|B 北京大学医学出版社|B 北京工艺美术出版社|B 北京联合出版公司|B 北京邮电大学出版社|B 北方文艺出版社|C 长春出版社|C 重庆大学出版社|D 第二军医大学出版社|D 第四军医大学出版社|G 高等教育出版社|G 光明日报出版社|G 广西美术出版社|G 广西师范大学出版社|G 贵州人民出版社|H 海洋出版社|H 河北美术出版社|H 合記圖書出版社|H 河南科学技术出版社|H 湖北美术出版社|H 湖南科学技术出版社|H 湖南美术出版社|H 湖南人民出版社|H 海南出版社|H 化学工业出版社|J 吉林科学技术出版社|J 吉林美术出版社|J 金城出版社|J 江苏凤凰科学技术出版社|J 江苏科学技术出版社|J 江苏美术出版社|J 江西美术出版社|J 机械工业出版社|J 军事医学科学出版社|K 科学出版社|K 科学普及出版社|L 蓝天出版社|L 辽宁科学技术出版社|L 辽宁美术出版社|L 岭南美术出版社|R 人民教育出版社|R 人民军医出版社|R 人民美术出版社|R 人民卫生出版社|R 人民邮电出版社|S 山东画报出版社|S 山东科学技术出版社|S 陕西科学技术出版社|S 商务印书馆|S 上海交通大学出版社|S 上海教育出版社|S 上海科学技术出版社|S 上海科学技术文献出版社|S 上海人民出版社|S 上海人民美术出版社|S 上海三联书店|S 上海社会科学院出版社|S 世界图书出版公司|S 四川大学出版社|T 天津科技翻译出版有限公司|T 天津科学技术出版社|T 天津人民出版社|T 天津人民美术出版社|T 天津杨柳青画社|W 文物出版社|X 新星出版社|Y 云南科技出版社|Z 浙江大学出版社|Z 浙江科学技术出版社|Z 浙江人民美术出版社|Z 知识产权出版社|Z 中国长安出版社|Z 中国电力出版社|Z 中国电影出版社|Z 中国建筑工业出版社|Z 中国连环画出版社|Z 中国青年出版社|Z 中国文联出版社|Z 中国协和医科大学出版社|Z 中国医药科技出版社|Z 中国传媒大学出版社|Z 中南大学出版社|Z 中山大学出版社|Z 中信出版社|Z 中央编译出版社|Z 中国水利水电出版社|Z 浙江摄影出版社
)
return