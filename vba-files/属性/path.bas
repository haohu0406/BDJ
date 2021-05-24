thisworkbook.path '当前工作簿路径
Application.Path '返回应用程序完整路径
Application.AutoRecover.Path '返回/设置Excel存储"自动恢复"临时文件的完整路径
Application.DefaultFilePath '选项>常规中的默认工作目录
Application.LibraryPath '返回库文件夹的路径
Application.NetworkTemplatesPath '返回保存模板的网络路径

Application.RecentFiles.Item(1).Path
'返回最近使用的某个文件路径,Item(1)=第一个文件

Application.StartupPath 'Excel启动文件夹的路径
Application.TemplatesPath '返回模板所存储的本地路径
Application.UserLibraryPath '返回用户计算机上 COM 加载宏的安装路径

ebug.Print Application.PathSeparator '路径分隔符 "\"
CurDir '默认工作目录
Excel.Parent.DefaultFilePath '默认工作目录