FileSystemObject对象位于File System对象模型的最高层 ，并且是该层次中惟一可以
在外部创建的对象 ，也就是说它是惟一能使用New关键字的对象 。
FileSystemObject对象有许多用来操作文件系统的方法和属性 。下面先看一个例子 ，
如下面的代码 ：

Sub FileInfo()
    Dim fs As Object
    Dim objFile As Object
    Dim strMsg As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set objFile = fs.GetFile("C:\Windows\System.ini")
    strMsg = "文件名:" & objFile.Name & VbCrLf
    strMsg = strMsg & "硬盘:" & objFile.Drive & VbCrLf
    strMsg = strMsg & "创建日期:" & objFile.DateCreated & VbCrLf
    strMsg = strMsg & "修改日期:" & objFile.DateLastModified & VbCrLf
    MsgBox strMsg, , "文件信息"
End Sub
FileInfo过程首先使用CreateObject函数创建一个FileSystemObject对象 ，用来访问
计算机的文件系统 。然后 ，使用GetFile方法创建一个File对象并返回对System.ini文件
的引用 。接着 ，利用File对象的Name属性 、Drive属性 、DateCreated属性 、
DateLastModified属性返回文件的相应信息 。
'下面详细介绍FileSystemObject对象的方法和属性 。

BuildPath方法
其语法为 ：

oFileSysObj.BuildPath(Path, Name)
其中 ，oFileSysObj为任何能够返回FileSystemObject对象的对象变量 。参数Path必需
，指定驱动器或文件夹路径 ，String类型 ，可以是绝对路径也可以是相对路径 ，不一
定要包含驱动器名 。参数Name必需 ，指定附加在Path后的文件夹或文件路径 ，
String类型 。参数Path或Name都不一定要求是当前已经存在的路径或文件夹 。
BuildPath方法通过合并参数Path和文件夹或文件名生成一个字符串 ，并且在必要的
地方加上正确的主机系统路径分隔符 。该方法不能检验新的文件夹或文件名的有效性 。
与人工合并两个字符串相比 ，使用BuildPath函数的惟一好处就是它能够选择正确的路
径分隔符 。

FileExists方法
FileExists方法用于判断指定的文件是否存在 ，若存在则返回True 。其语法为 ：
oFileSysObj.FileExists(FileSpec)
其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象 。参数FileSpec
必需 ，代表文件的完整路径 ，String类型 ，不能包含有通配符 。
如果用户有充分的权限 ，FileSpec可以是网络路径或共享名 ，例如 ：

If ofs.FileExists("\\TestPath\Test.txt") Then
    示例
    
    Sub IfFileExists()
        Dim fs As Object
        Dim strFile As String
        Set fs = CreateObject("Scripting.FileSystemObject")
        strFile = InputBox("请输入文件的完整名称:")
        If fs.FileExists(strFile) Then
            MsgBox strFile & "已经找到."
        Else
            MsgBox "该文件不存在."
        End If
    End Sub
    GetFile方法
    GetFile方法用来返回一个File对象 。其语法为 ：
    
    oFileSysObj.GetFile(FilePath)
    其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参数
    FilePath必需 ，指定路径和文件名 ，String类型 。可以是绝对路径或相对路径 。
    如果FilePath是一个共享名或网络路径 ，GetFile确认该驱动器或共享是File对象创
    建进程的一部分 。如果参数FilePath指定的路径的任何部分不能连接或不存在 ，就
    会产生错误 。
    GetFile方法返回的是File对象 ，而不是TextStream对象 。File对象不是打开的文件    
    ，主要是用来完成如复制或移动文件和询问文件的属性之类的方法 。尽管不能对File
    对象进行写或读操作 ，但可以使用File对象的OpenAsTextStream方法获得
    TextStream对象 。
    要获得所需的FilePath字符串 ，首先应该使用GetAbsolutePathName方法 。如果
    FilePath包含网络驱动器或共享 ，可以在调用GetFile方法之前用DriveExists方法来
    检验所需的驱动器是否可用 。
    因为在FilePath指定的文件不存在时会产生错误 ，所以应该在调用GetFile之前调用
    FileExists方法确定文件是否存在 。
    必须用Set语句将File对象赋给一个局部对象变量 。
    
    GetFileName方法
    GetFileName方法返回给定路径的文件名称部分 。其语法为 ：
    oFileSysObj.GetFileName(Path)
    其中 ，oFileSysObj表示任何能够返回FileSystemObject对象的对象变量 。参数
    Path必需 ，指定路径说明 ，String类型 。如果不能从给定的Path确定文件名 ，则
    返回一个零长字符串 （””）。Path可以为绝对路径或相对路径 。
    GetFileName方法不能检验Path中是否存在指定的文件 。Path可以为网络驱动器或
    共享 。GetFileName本身不具有智能 ，它认为字符串中不属于驱动器说明的最后部分就是一个文件名 ，更像是一个字符串处理函数而不是对象处理方法 。
    GetFileVersion方法
    GetFileVersion方法返回文件的版本 。
    CopyFile方法
    CopyFile方法用来复制文件 ，将文件从一个文件夹复制到另一个文件夹 。其语法为
    ：
    
    oFileSysObj.CopyFile Source, Destination[, OverwriteFiles]
    其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参数
    Source必需 ，指定要复制的文件的路径和名称 ，String类型 。参数Destination必
    需 ，代表复制文件的目标路径和文件名 （可选 ），String类型 。参数
    OverwriteFiles可选 ，表示是否覆盖一个现有文件的标志 ，True表示覆盖 ，False
    表示不覆盖 ，Boolean类型 ，默认值为True 。
    参数source中源路径可以是绝对路径或相对路径 ，源文件名可包含通配符但源路径不
    能 。在参数Destination中不能包含通配符 。
    如果目标路径或文件设置为只读 ，则无论OverwriteFiles参数的值如何 ，都将无法完
    成CopyFile方法 。如果参数OverwriteFiles设置为False且Destination指定的文件已经
    存在 ，则会产生一个运行时错误 “文件已经存在 ”。如果在复制多个文件时出现错误    
    ，CopyFile方法将立即停止复制操作 ，该方法不具有撤销错误前文件复制操作的返回
    功能 。如果用户有充分的权限 ，那么source或destination可以是网络路径或共享名    
    ，例如 ：
    
    CopyFile "\\NTSERV1\RootTest\test.txt", "C:\RootOne"
    CopyFile方法可以复制一个保存在特定文件夹中的文件 。如果文件夹本身有包含文件
    的子文件夹 ，则使用CopyFile方法不能复制这些文件 ，应该使用CopyFolder方法 。
    例如 ：
    
    Sub CopyFile()
        Dim fs As Object
        Dim strFile As String
        Dim strNewFile As String
        strFile = "C:\test.doc"
        strNewFile = "C:\Program Files\test.doc"
        Set fs = CreateObject("Scripting.FileSystemObject")
        fs.CopyFile strFile, strNewFile
        MsgBox "已经创建了指定文件的副本."
        Set fs = Nothing
    End Sub
    
    CreateTextFile方法
    CreateTextFile方法创建一个新的文件并返回其TextStream对象 ，其语法为 ：
    oFileSysObj.CreateTextFile Filename[, Overwrite[, Unicode]]    
    
    其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参数
    Filename必需 ，代表任何有效文件名 ，String类型 。在Filename中不允许使用通配
    符 。Filename可以是相对路径也可以是绝对路径 ，如果没有指定路径 ，则使用应用
    程序的当前驱动器或文件夹作为路径 。如果指定的路径不存在 ，则该方法将失败 。
    参数Overwrite可选 ，作为一个标志 ，指定是否覆盖一个具有相同文件名的现有文件    
    ，Boolean类型 。默认值为False 。
    参数Unicode可选 ，作为一个标志 ，指明用Unicode格式还是ASC Ⅱ格式写文件 ，
    Boolean类型 。如果设置为True ，则以Unicode格式创建文件 ，否则创建一个
    Asc Ⅱ文本文件 。默认值为False 。
    只有写操作才能使新创建的文本文件自动打开 ，如果以后希望读取该文件 ，则必须选
    关闭它再以读模式重新打开该文件 。
    如果参数Filename中指定的路径设置为只读 ，则不论参数Overwrite的值如保 ，
    CreateTextFile方法都将失败 。
    如果用户有充分的权限 ，那么参数Filename可以是网络路径或共享名 ，例如 ：
    
    CreateTextFile "\\NTSERV1\RootTest\myFile.doc"
    必须使用Set语句将TextStream对象赋值给局部对象变量 。
    
    MoveFile方法
    MoveFile方法用来移动文件 ，将文件从一个文件夹移动到另一个文件夹 。其语法为
    oFileSysObj.MoveFile source, destination
    其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参数
    source必需 ，指定要移动的文件的路径 ，String类型 。参数destination必需 ，指
    定文件移动操作中的目标位置的路径 ，String类型 。
    如果Source包含通配符或者destination以路径分隔符结尾 ，则认为destination是
    一个路径 ，否则认为destination的最后一部分是文件名 。
    如果目标文件已经存在 ，则将出现一个错误 。
    source可以包含通配符 ，但只能出现在它的最后一部分中 。
    destination参数不能包含通配符 。
    source或destination可以是相对路径或绝对路径 ，可以是网络路径或共享名 。
    MoveFile方法在开始操作前先解析source和destination这两个参数 。
    
    DeleteFile方法
    DeleteFile方法删除指定的一个或多个文件 。其语法为 ：
    oFileSysObj.DeleteFile FileSpec[, Force]
    其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参数
    FileSpec必需 ，代表要删除的单个文件或多个文件的名称和路径 ，String类型 ，可
    以在路径的最后部分包含通配符 ，可以为相对路径或绝对路径 。如果在FileSpec中
    只有文件名 ，则认为该文件在应用程序的当前驱动器和文件夹中 。参数Force可选    
    ，如果将其设置为True ，则忽略文件的只读标志并删除该文件 ，Boolean类型 ，
    默认值为False 。
    如果指定要删除的文件已经打开 ，该方法将失败并出现一个 “Permission Denied    
    ”错误 。如果找不到指定的文件 ，则该方法失败 。
    如果在删除多个文件的过程中出现错误 ，DeleteFile方法将立即停止删除操作 ，即
    不能删除余下的文件部分 。该方法不具有撤销产生错误前文件删除操作的返回功能 。
    如果用户有充分的权限 ，源路径或目标路径可以是网络路径或共享名 。例如 ：
    
    DeleteFile “ \ \NTSERV1 \ RootTest \ MyFile.doc ”
    DeleteFile方法永久性地删除文件 ，并不把这些文件移到回收站中 。
    示例
    
    Sub DeleteFile()
        Dim fs As FileSystemObject
        Set fs = New FileSystemObject
        fs.DeleteFile "C:\test.doc"
        MsgBox "删除了该文件."
    End Sub
    DriveExists方法
    DriveExists方法用来判断在本地计算机或者网络上是否存在指定的磁盘 ，若存在则
    返回True 。其语法为 ：
    
    oFileSysObj.DriveExists(DriveSpec)
    其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参数
    DriveSpec必需 ，代表路径或驱动器名 ，String类型 。如果DriveSpec是一个
    Windows驱动器名 ，则其后面不需要跟冒号 ，例如 “C ”和 “C ：”是一样的 。
    DriveExists方法不能返回可移动驱动器的当前状态 ，要实现这一目的 ，必须使用
    指定驱动器的IsReady属性 。
    如果用户有充分的权限 ，DriveSpec可以是网络路径或共享名 ，例如 ：
    
    If ofs.DriveExists("\\NTESERV1\d$") Then
        在调用位于某驱动器上一个远程ActiveX服务器中的函数前 ，最好先使用
        DriveExists方法检测网络上是否存在该驱动器 。
        示例
        
        Function DriveExists(disk)
            Dim fs As Object
            Dim strMsg As String
            Set fs = CreateObject("Scripting.FileSystemObject")
            If fs.DriveExists(disk) Then
                strMsg = "驱动器" & UCase(disk) & "盘已存在."
            Else
                strMsg = UCase(disk) & "盘未找到."
            End If
            DriveExists = strMsg
        End Function
        GetDrive方法
        GetDrive方法返回Drive对象 ，即获得对指定驱动器的Drive对象的引用 。其语法
        为 ：
        
        oFileSysObj.GetDrive(drivespecifier)
        其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参数
        drivespecifier必需 ，代表驱动器名 、共享名或网络路径 ，String类型 。如果
        drivespecifier是一个共享名或网络路径 ，GetDrive确认它是Drive对象创建进程
        的一部分 ，否则会产生运行时错误 “找不到路径 ”。如果指定的驱动器没有连接上
        或者不存在 ，则会出现运行时错误 “设备不可用 ”。
        如果要从路径中导出drivespecifier字符串 ，应该首先用GetAbsolutePathName
        来确保驱动器是路径的一部分 ，然后在调用GetDriveName从全限定路径中提取
        出驱动器之前 ，用FolderExists方法检验路径是否有效 ，例如 ：
        
        Dim oFileSys As New FileSystemObject
        Dim oDrive As Drive
        sPath = oFileSys.GetAbsolutePathName(sPath)
        If oFileSys.FolderExists(sPath) Then
            Set oDrive = oFileSys.GetDrive(oFileSys.GetDriveName(sPath))
        End If
        如果driverspecifier是一个网络驱动器或共享 ，在调用GetDrive方法之前 ，应该
        用DriveExists方法检验所需的驱动器是否可用 。
        必须用Set语句将Drive对象赋给局部对象变量 。
        示例
        
        Sub DriveInfo()
            Dim fs, disk, infoStr, strDiskName
            strDiskName = InputBox("输入驱动器盘符:", "驱动器名称", "C:\")
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set disk = fs.GetDrive(fs.GetDriveName(strDiskName))
            infoStr = "驱动器:" & UCase(strDiskName) & VbCrLf
            infoStr = infoStr & "驱动器盘符:" & UCase(disk.DriveLetter) & VbCrLf
            infoStr = infoStr & "驱动器类型:" & disk.DriveType & VbCrLf
            infoStr = infoStr & "驱动文件系统:" & disk.FileSystem & VbCrLf
            infoStr = infoStr & "驱动器系列号:" & disk.SerialNumber & VbCrLf
            infoStr = infoStr & "字节的总大小:" & FormatNumber(disk.TotalSize / 1024, 0) & "kb" & VbCrLf
            infoStr = infoStr & "驱动器中的自由空间:" & FormatNumber(disk.FreeSpace / 1024, 0) & "kb" & VbCrLf
            MsgBox infoStr, vbInformation, "驱动器信息"
        End Sub
        
        GetDriveName方法
        GetDriveName方法返回一个包含硬盘名称或者网络共享名称的字符串 。
        即返回给定路径的驱动器名 ，如果从给定的路径中不能确定驱动器名 ，则返回
        一个零长字符串 （””）。其语法为 ：
        oFileSysObj.GetDriveName(Path)
        其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参数
        Path必需 ，指定路径 ，String类型 。
        GetDriveName不能检验Path中是否存在指定的驱动器 。Path可以是网络驱动器
        或共享 。
        示例
        
        Function DriveName(disk)
            Dim fs As Object
            Dim strDiskName As String
            Set fs = CreateObject("Scripting.FileSystemObject")
            strDiskName = fs.GetDriveName(disk)
            DriveName = strDiskName
        End Function
        
        GetExtensionName方法
        返回给定路径中文件的扩展名 。其语法为 ：
        oFileSysObj.GetExtensionName(Path)
        其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参数
        Path必需 ，表示路径说明 ，String类型 。如果不能确定Path中的扩展名 ，则返
        回一个零长字符串 。
        GetExtensionName方法不能检验Path是否有效 ，Path可以为网络路径或共享 。
        GetExtensionName没有智能功能 ，它简单地解析一个字符串 ，并返回Path最
        后部分中最后一个点后的文本 。
        FolderExists方法
        FolderExists方法可以判断指定的文件夹是否存在 ，若存在则返回True 。其语法
        为 ：
        
        oFileSysObj.FolderExists(FolderSpec)
        其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参数
        FolderSpec指定文件夹的完整路径 ，String类型 ，不能包含通配符 。
        如果用户有充分的权限 ，FolderSpec可以是网络路径或共享名 ，例如 ：
        
        If ofs.FileExists("\\NTSERV1\d$\TestPath\") Then
            示例
            
            Sub IfFolderExists()
                Dim fs As Object
                Set fs = CreateObject("Scripting.FileSystemObject")
                MsgBox fs.FolderExists("C:\Program Files")
            End Sub
            GetAbsolutePathName方法
            将相对路径转变为一个全限定路径 （包括驱动器名 ），返回一个字符串 ，包
            含一个给定的路径说明的绝对路径 。其语法为 ：
            
            oFileSysObj.GetAbsolutePathName(Path)
            其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参
            数Path必需 ，代表路径说明 ，String类型 。
            “.”返回当前文件夹的驱动器名和完整路径 。“..”返回当前文件夹的父文件夹的
            驱动器名和路径 。“filename ”返回当前文件夹中的文件的驱动器名 、路径及
            文件名 。
            所有相对路径名均以当前文件夹为基准 。
            如果没有明确地提供驱动器作为Path的一部分 ，就以当前驱动器作为Path参数
            中的驱动器 。在Path中可以包含任意个通配符 。
            对于映射网络驱动器和共享而言 ，这种方法不能返回完整的网络地址 ，而是
            返回全限定的本地路径和本地驱动器名 。
            GetAbsolutePathName不能检验指定路径中是否存在某个给定的文件或文件夹
            
            GetBaseName方法
            返回路径的最后部分的名称 ，不包含扩展名 。其语法为 ：
            oFileSysObj.GetBaseName(Path)
            其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参
            数Path必需 ，代表路径说明 ，String类型 。Path中最后部分的文件扩展名不
            包含在返回的字符串中 。
            GetBaseName方法不能检验Path中是否存在给定的文件或文件夹 。
            GetBaseName方法没有舍去文件扩展名并返回Path的基本名称的智能功能 。
            也就是说 ，它不能识别路径的最后部分是路径还是文件名 。如果最后部分包括
            一个或多个点 “.”，它仅仅删除最后一个占以及该点后的文本 。所以如果Path
            为 “.”，GetBaseName方法返回一个空字符串 ；如果Path为 “..”，
            GetBaseName方法返回 “.”。换句话说 ，它只不过是一个字符串处理函数 ，而
            不是一个文件函数 。
            
            GetFolder方法
            GetFolder方法返回Folder对象 。其语法为 ：
            oFileSysObj.GetFolder(FolderPath)
            其中 ，oFileSysObj代表任何能返回FileSystemObject对象的对象变量 。参数
            FolderPath必需 ，指定所需文件夹的路径 ，String类型 ，可以为相对路径或绝
            对路径 。如果FolderPath是共享名或网络路径 ，GetFolder确认该驱动器或共
            享是File对象创建进程的一部分 。如果FolderPath的任何部分不能连接或不存在
            ，就会产生一个错误 。
            要获得所需的Path字符串 ，首先应该使用GetAbsolutePathName方法 。如果
            FolderPath包含一个网络驱动器或共享 ，可以在调用GetFolder方法之前使用
            DriveExists方法确认指定的驱动器是否可用 。由于GetFolder方法要求
            FolderPath是一个有效文件夹的路径 ，所以应调用FolderExists方法来检验
            FolderPath是否存在 。
            必须使用Set语句将Folder对象赋给一个局部对象变量 。
            示例
            
            Sub FilesInFolder()
                Dim fs As Object
                Dim objFolder As Object
                Dim objFile As Object
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set objFolder = fs.GetFolder("C:\")
                Workbooks.Add
                For Each objFile In objFolder.Files
                    ActiveCell.Select
                    Selection.Formula = objFile.Name
                    ActiveCell.Offset(0, 1).Range("A1").Select
                    Selection.Formula = objFile. Type
                    ActiveCell.Offset(1,  - 1).Range("A1").Select
                Next
                Columns("A:B").Select
                Selection.Columns.AutoFit
            End Sub
            
            GetParentFolderName方法
            返回给定路径中最后部分前的文件夹名 ，其语法为 ：
            oFileSysObj.GetParentFolderName(Path)
            其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。
            参数Path必需 ，指定路径说明 ，String类型 。
            如果从Path中不能确定父文件夹名 ，就返回一个零长字符串 （””）。Path可以
            为相对路径或绝对路径 。可以是网络驱动器或共享 。
            GetParentFolderName方法不能检验Path的某个部分是否存在 。
            GetParentFolderName方法认为Path中不属于驱动器说明的那部分字符串除
            了最后一部分外余下的字符串就是父文件夹 。除此之外它不做任何其他检测 ，
            更像是一个字符串解析和处理例程而不是与对象处理有关的例程 。
            
            GetSpecialFolder方法
            GetSpecialFolder方法返回操作系统文件夹路径 ，其中0代表Windows文件夹 ，
            1 代表System （系统 ）文件夹 ，2 代表Temp （临时 ）文件夹 。其语法为 ：
            oFileSysObj.GetSpecialFolder(SpecialFolder)
            其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参
            数SpecialFolder必需 ，为特殊的文件夹常数 ，表示三种特殊系统文件夹中其
            中一个的值 。
            可以使用Set语句将Folder对象赋给一个局部对象变量 ，但是如果只对检索特
            殊的文件夹感兴趣 ，就可以使用下列语句来实现 ：
            
            sPath = oFileSys.GetSpecialFolder(iFolderConst)
            或 ：
            sPath = oFileSys.GetSpecialFolder(iFolderConst).Path
            由于Path属性是Folder对象的缺省属性 ，所认第一个语句有效 。因为不是给一
            个对象变量赋值 ，所以赋给sPath的值是缺省的Path属性值 ，而不是对象引用
            示例
            
            Sub SpecialFolders()
                Dim fs As Object
                Dim strWindowsFolder As String
                Dim strSystemFolder As String
                Dim strTempFolder As String
                Set fs = CreateObject("Scripting.FileSystemObject")
                strWindowsFolder = fs.GetSpecialFolder(0)
                strSystemFolder = fs.GetSpecialFolder(1)
                strTempFolder = fs.GetSpecialFolder(2)
                MsgBox strWindowsFolder & VbCrLf & strSystemFolder & VbCrLf _
                         & strTempFolder, vbInformation + vbOKOnly, "Special Folders"
            End Sub
            
            GetTempName方法
            返回系统创建的一个临时文件或文件夹名 。其语法为 ：
            oFileSysObj.GetTempName
            其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。
            GetTempName方法不能创建临时文件或文件夹 ，它仅仅提供一个可用于
            CreateTextFile方法的文件或文件夹名 。
            一般来说 ，不必创建自已的临时文件名 。Windows在Windows API中提供
            了一种算法来创建特殊的临时文件或文件夹名 ，这样Windows才能识别它们            
            。GetTempName很好地包装了GetTempFilename API函数 。
            CreateFolder方法
            CreateFolder方法用于在指定的路径下创建一个新文件夹 ，并返回其Folder
            对象 。其语法为 ：
            
            oFileSysObj.CreateFolder(Path)
            其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参数
            Path必需 ，为一个返回要创建的新文件夹名的表达式 ，String类型 。Path
            指定的路径可以是相对路径也可以是绝对路径 ，如果没有指定路径则使用当前
            驱动器和目录作为路径 。在新的文件夹名中不能使用通配符 。
            如果参数Path指定的路径为只读 ，则CreateFolder方法将失败 ；如果参数Path
            指定的文件夹已经存在 ，就会产生运行时错误 “文件已经存在 ”。如果用户有充
            分的权限 ，则参数Path可以指定为网络路径或共享名 ，例如 ：
            
            CreateFolder "\\NTSERV1\RootTest\newFolder"
            在实际使用时 ，必须使用Set语句将Folder对象赋给对象变量 ，例如 ：
            
            Dim oFileSys As New FileSystemObject
            Dim oFolder As Folder
            Set oFolder = oFileSys.CreateFolder("MyFolder")
            示例
            
            Sub MakeNewFolder()
                Dim fs, objFolder
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set objFolder = fs.CreateFolder("C:\TestFolder")
                MsgBox "创建了一个名称为" & objFolder.Name & "的文件夹."
            End Sub
            
            
            CopyFolder方法
            CopyFolder方法用于复制文件夹 ，即将一个文件夹的内容 （包括其子文件夹 ）复制到其他位置 。其语法为 ：
            oFileSysObj.CopyFolder Source, Destination[, OverwriteFiles]
            其中 ，参数oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。
            参数Source必需 ，指定要复制的文件夹的路径和文件夹名 ，String类型 ，必须
            使用通配符或者非路径分隔符来结束 。参数Destination必需 ，指定文件夹复
            制操作的目标文件夹的路径 ，String类型 。参数OverwriteFiles可选 ，表示是
            否被覆盖一个现有文件的标志 ，True表示覆盖 ，False表示不覆盖 ，Boolean
            类型 。
            通配符只能在参数Source中使用 ，但是只能放在最后的组件中 。在参数
            Destination中不能使用通配符 。
            除非不允许使用通配符 ，否则就可以把源文件夹中的所有子文件夹和文件都复
            制到Destination指定的文件夹中 ，也就是说CopyFolder方法是递归的 。
            如果参数Destination以一个路径分隔符结束或者参数Source以一个通配符结
            束 ，CopyFolder方法就认为参数Source中的指定的文件夹存在于参数
            Destination中 ，否则就创建这样一个文件夹 。例如 ，假设有如下的文件夹
            结构 ：
            
            
            CopyFolder "C:\Rootone\*", "C:\RootTwo"
            产生如下的文件夹结构 ：
            CopyFolder "C:\Rootone", "C:\RootTwo\"
            产生如下的文件夹结构 ：
            filesystem3
            如果参数Destination指定的目标路径或任意文件被设置成只读属性 ，则不论OverwriteFiles的值如何 ，CopyFolder方法者将失效 。
            如果OverwriterFiles设置为False ，而参数Source指定的源文件夹或任何文件存在于参数Destination中 ，将产生运行时错误 “文件已经存在 ”。
            如果在复制多个文件夹时出现错误 ，CopyFolder方法立即停止复制操作 ，不再复制余下要复制的文件 。该方法不具有撤销产生错误前文件复制操作的返回功能 。
            如果用户有充分的权限 ，source或destination都可以是网络路径或共享名 ，例如 ：
            
            CopyFolder "C:\Rootone", "\\NTSERV1\d$\RootTwo\"
            Sub MakeFolderCopy()
                Dim fs As FileSystemObject
                Set fs = New FileSystemObject
                If fs.FolderExists("C:\TestFolder") Then
                    fs.CopyFolder "C:\TestFolder", "C:\FinalFolder"
                    MsgBox "已复制该文件夹."
                End If
            End Sub
            
            MoveFolder方法
            MoveFolder方法用来移动文件夹 ，将文件夹及其文件和子文件夹一起从某
            位置移动到另一个位置 。其语法为 ：
            oFileSysObj.MoveFolder source, destination
            其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。
            参数Source指定要移动的文件夹的路径 ，String类型 。参数destination指定
            文件夹移动操作中目标位置的路径 ，String类型 。
            Source必须以通配符或非路径分隔符结束 ，可以使用通配符 ，但必须出现在
            最后一部分中 。destination不能使用通配符 。除非不允许使用通配符 ，否则
            源文件夹中所有的子文件夹和文件都被复制到destination指定的位置 ，也就
            是说MoveFolder方法是递归的 。
            如果destination用路径分隔符结束或者source用通配符结束 ，MoveFolder
            就认为source中指定的文件夹存在于destination中 。例如 ，假设有如下文件
            夹结构 ：
            filesystem4
            
            MoveFolder "C:\Rootone\*", "C:\RootTow\"
            产生如下文件夹结构 ：
            filesystem5
            
            MoveFolder "C:\Rootone", "C:\RootTwo\"
            产生如下文件夹结构 ：
            filesystem6
            Source和destination可以为绝对路径或相对路径 ，可以为网络路径或共享名
            MoveFile方法在开始操作前先解析source和destination这两个参数 。
            DeleteFolder方法
            DeleteFolder方法用于删除指定的文件夹及其所有的文件和子文件夹 。其语
            法为 ：
            
            oFileSysObj.DeleteFolder FileSpec[, Force]
            其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参
            数FileSpec必需 ，指定要删除的文件夹的名称和路径 ，String类型 。在参数
            FileSpec中 ，可以在路径的最后部分包含通配符 ，但不能用路径分隔符结束 ，
            可以为相对路径或绝对路径 。
            参数Force可选 ，Boolean类型 ，如果设置为True ，将忽略文件的只读标志并
            删除这个文件 。默认为False 。如果参数Force设置为False并且文件夹中的任
            意一个文件为只读 ，则该方法将失败 。如果找不到指定的文件夹 ，则该方法失
            败 。
            如果指定的文件夹中有文件已经打开 ，则不能完成删除操作 ，且产生一个
            “Permisson Denied ”错误 。DeleteFolder方法删除指定文件夹中的所有内容 ，包括其他文件夹及其内容 。
            如果在删除多个文件或文件夹时出现错误 ，DeleteFolder方法将立即停止删除
            操作 ，即不能删除余下的文件夹或文件 。该方法不具有撤销产生错误前文件夹
            删除操作的返回功能 。
            DeleteFolder方法永久性删除文件夹 ，并不把它们移到回收站中 。
            如果用户有充分的权限 ，源路径和目标路径可以是网络路径或共享名 ，例如 ：
            
            DeleteFolder "\\RootTest"
            示例
            
            Sub RemoveFolder()
                Dim fs As FileSystemObject
                Set fs = New FileSystemObject
                If fs.FolderExists("C:\TestFolder") Then
                    fs.DeleteFolder "C:\TestFolder"
                    MsgBox "该文件夹已经被删除."
                End If
            End Sub
            OpenTextFile方法
            OpenTextFile方法用于打开 （或创建 ）文本文件以进行读取或写入操作 ，返
            回一个TextStream对象 。其语法为 ：
            
            oFileSysObj.OpenTextFile(FileName[, IOMode[, Create[, Format]]])
            其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。参
            数FileName必需 ，指定要打开的文件的路径和文件名 ，String类型 ，
            FileName的路径部分可为相对路径或绝对路径 。参数IOMode可选 ，指定文
            件打开模式的一个常数 （参见前文的表格 ），默认设置为ForReading （1 ）
            。参数Create可选 ，一个Boolean型标志 ，说明如果在给定的路径中找不到
            文件 ，是否应该创建该文件 。参数Format可选 ，一个Tristate常数 （参见前
            文的表格 ），指定打开文件的格式为ASC Ⅱ或Unicode格式 ，默认设置为
            Asc Ⅱ（False ）。
            如果另一个进程已经打开了指定文件 ，该方法将失败 ，并产生一个
            “Permission Denied ”错误 。
            要保证OpenTextFile方法成功执行 ，可以在调用它之前使用GetAbsolutePath
            和FileExists方法 。
            IOMode的值只能是一个常量值 ，例如 ，下面的方法调用 ：
            
            lMode = ForReading Or ForWriting
            oFileSys.OpenTextStream(strFileName, lMode) '错误
            将产生运行时错误 “无效的过程调用或参数 ”。
            如果用户有充分的权限 ，FileName的路径部分可以是网络路径或共享名 ，
            例如 ：
            
            OpenTextFile "\\NTSERV1\d$\RootTwo\myFile.txt"
            示例
            
            Sub ReadTextFile()
                Dim fs As Object
                Dim objFile As Object
                Dim strContent As String
                Dim strFileName As String
                strFileName = "C:\Windows\System.ini"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set objFile = fs.OpenTextFile(strFileName)
                Do While Not objFile.AtEndOfStream
                    strContent = strContent & objFile.ReadLine & VbCrLf
                Loop
                objFile.Close
                Set objFile = Nothing
                ActiveWorkbook.Sheets(3).Select
                Range("A1").Select
                Selection.Formula = strContent
            End Sub
            
            Drives属性
            Drives属性是FileSystemObject对象唯一的属性 ，返回对硬盘驱动器集合            
            （Drives ）的引用 ，是一个只读属性 。其语法为 ：
            oFileSysObj.Drives
            其中 ，oFileSysObj代表任何能够返回FileSystemObject对象的对象变量 。
            Drives属性返回的集合中的每个成员都是Drive对象 ，表示系统中一个可用的
            驱动器 。可以使用For …Next循环迭代系统中所有驱动器 ，或者使用Drives集
            合的Item方法读取某个Drive对象 （代表系统中的某个驱动器 ）。
            例如 ，下面的代码创建计算机驱动盘清单 ：
            
            Sub DrivesList()
                Dim fs As Object
                Dim colDrives As Object
                Dim Drive As Object
                Dim strDrive As String
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set colDrives = fs.Drives
                For Each Drive In colDrives
                    strDrive = "驱动器" & Drive.DriveLetter & ":"
                    Debug.Print strDrive
                Next
            End Sub