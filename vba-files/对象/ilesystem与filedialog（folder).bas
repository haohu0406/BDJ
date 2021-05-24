With application.filedialog(msofiledialogfolderpicker)
    mypath =  .selecteditems(1)
    
    
    
    Set fso = CreateObject("scripting.filesystemobject")
    Set ff = fso.getfolder(mypath)
    '上一句返回一个文件路径字符串，下一句返回一个文件夹对象，对对象ff的属性files
    （文件 ）可进行遍历文件名